import io
import os
import json
import datetime
import pdfplumber
import anthropic
from flask import Flask, render_template, request, send_file, jsonify
from openpyxl import load_workbook
from openpyxl.styles import Protection, Alignment

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32MB max upload

TEMPLATE_FILE = "BrokersBloc_Medical_Template.xlsx"
GROUP_COLS = [(8,9,10), (11,12,13), (14,15,16), (17,18,19), (20,21,22)]

ROW_MAP = {
    "network":        12,
    "deductible":     13,
    "coinsurance":    14,
    "moop":           15,
    "pcp":            16,
    "telehealth":     17,
    "specialist":     18,
    "inpatient":      19,
    "outpatient":     20,
    "er":             21,
    "urgent_care":    22,
    "lab":            23,
    "xray":           24,
    "imaging":        25,
    "oon_ded":        27,
    "oon_cost":       28,
    "oon_moop":       29,
    "rx_ded":         31,
    "rx_copays":      32,
    "rate_ee":        34,
    "rate_es":        35,
    "rate_ec":        36,
    "rate_ef":        37,
    "rate_guarantee": 38,
}


def extract_pdf_text(file_bytes):
    text_pages = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text_pages.append(t)
    return "\n\n".join(text_pages)


def parse_pdf_with_ai(pdf_text, plan_filter, enrollment):
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    system = """You are a health insurance data extraction specialist.
You read carrier proposal PDFs and extract structured plan benefit data.
You ALWAYS respond with valid JSON only — no explanation, no markdown, no code blocks.
Monthly rates must be numbers (floats). All other values are strings.
If a value is not found, use an empty string."""

    prompt = f"""Extract health insurance plan data from this carrier proposal.

FILTER: Only include plans where the individual in-network deductible matches: {plan_filter}
If filter is "all plans", include all plans found (max 3).

Return ONLY this JSON structure:
{{
  "carrier_name": "Carrier name",
  "commissions": "Commission info or ''",
  "carrier_comments": "Funding type, network, PBM, quote ID, stop-loss, key notes",
  "plans": [
    {{
      "plan_name": "Full plan name",
      "network": "Network name and PBM",
      "deductible": "Ind / Fam in-network",
      "coinsurance": "Member % / Plan % description",
      "moop": "Ind / Fam OOPM in-network",
      "pcp": "PCP cost share",
      "telehealth": "Telehealth cost and vendor",
      "specialist": "Specialist cost share",
      "inpatient": "Inpatient hospitalization",
      "outpatient": "Outpatient facility",
      "er": "Emergency room",
      "urgent_care": "Urgent care",
      "lab": "Laboratory",
      "xray": "X-ray",
      "imaging": "CT/MRI/PET",
      "oon_ded": "OON deductible Ind / Fam",
      "oon_cost": "OON coinsurance",
      "oon_moop": "OON OOPM Ind / Fam",
      "rx_ded": "Rx deductible",
      "rx_copays": "Generic / Brand / Non-pref / Specialty",
      "rate_ee": 0.00,
      "rate_es": 0.00,
      "rate_ec": 0.00,
      "rate_ef": 0.00,
      "rate_guarantee": "12 Months"
    }}
  ]
}}

PDF TEXT:
{pdf_text[:14000]}"""

    message = client.messages.create(
        model="claude-opus-4-5",
        max_tokens=4096,
        system=system,
        messages=[{"role": "user", "content": prompt}]
    )

    raw = message.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.rstrip("`").strip()
    return json.loads(raw)


def sv(ws, row, col, val, number_fmt=None):
    cell = ws.cell(row=row, column=col)
    if type(cell).__name__ == "MergedCell":
        return
    cell.value = val
    cell.protection = Protection(locked=False)
    if number_fmt:
        cell.number_format = number_fmt


def populate_excel(group_name, effective_date, quote_date, enrollment, carrier_groups):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb["Medical"]

    ws["C3"].value = group_name
    ws["C3"].protection = Protection(locked=False)
    ws["C3"].alignment = Alignment(horizontal="left", vertical="center",
                                    wrap_text=False, shrink_to_fit=False)
    sv(ws, 5, 9, f"Effective Date: {effective_date}   |   Quote Date: {quote_date}")

    for row, key in [(34,"ee"),(35,"es"),(36,"ec"),(37,"ef")]:
        for col in [4, 5, 6]:
            try:
                sv(ws, row, col, int(enrollment.get(key, 0)))
            except ValueError:
                sv(ws, row, col, 0)

    for gi, group in enumerate(carrier_groups):
        if not group:
            continue
        c1, c2, c3 = GROUP_COLS[gi]
        sv(ws, 9,  c1, group.get("carrier_name", ""))
        sv(ws, 53, c1, group.get("commissions", ""))
        sv(ws, 39, c1, group.get("carrier_comments", ""))

        plan_cols = [c1, c2, c3]
        for pi, plan in enumerate(group.get("plans", [])[:3]):
            col = plan_cols[pi]
            sv(ws, 10, col, plan.get("plan_name", ""))
            for field, row in ROW_MAP.items():
                val = plan.get(field, "")
                if field.startswith("rate_") and val:
                    try:
                        sv(ws, row, col, float(val), '"$"#,##0.00')
                    except (ValueError, TypeError):
                        sv(ws, row, col, val)
                else:
                    sv(ws, row, col, val)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    try:
        group_name     = request.form.get("group_name", "Quote")
        effective_date = request.form.get("effective_date", "")
        quote_date     = request.form.get("quote_date", str(datetime.date.today()))
        plan_filter    = request.form.get("plan_filter", "all plans")

        enrollment = {
            "ee": request.form.get("enroll_ee", "0"),
            "es": request.form.get("enroll_es", "0"),
            "ec": request.form.get("enroll_ec", "0"),
            "ef": request.form.get("enroll_ef", "0"),
        }

        carrier_groups = []
        for gi in range(5):
            pdf_file = request.files.get(f"pdf_{gi}")
            if pdf_file and pdf_file.filename:
                pdf_bytes  = pdf_file.read()
                pdf_text   = extract_pdf_text(pdf_bytes)
                group_data = parse_pdf_with_ai(pdf_text, plan_filter, enrollment)
                carrier_groups.append(group_data)
            else:
                carrier_groups.append(None)

        if not any(carrier_groups):
            return jsonify({"error": "Please upload at least one carrier PDF."}), 400

        buf = populate_excel(group_name, effective_date, quote_date,
                             enrollment, carrier_groups)

        safe_name = "".join(c for c in group_name if c.isalnum() or c in " _-").strip()
        filename  = f"{safe_name or 'Medical_Quote'}_Comparison.xlsx"

        return send_file(buf, as_attachment=True, download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except json.JSONDecodeError as e:
        return jsonify({"error": f"AI could not parse the PDF response. Try again. ({e})"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True)
