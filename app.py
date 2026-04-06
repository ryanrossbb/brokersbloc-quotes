import io
import datetime
from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook
from openpyxl.styles import Protection, Alignment

app = Flask(__name__)

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


def sv(ws, row, col, val, number_fmt=None):
    cell = ws.cell(row=row, column=col)
    if type(cell).__name__ == "MergedCell":
        return
    cell.value = val
    cell.protection = Protection(locked=False)
    if number_fmt:
        cell.number_format = number_fmt


def safe_float(val):
    try:
        return float(val.replace(",", "").replace("$", "").strip())
    except Exception:
        return val  # leave as string if not numeric


def generate_excel(form):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb["Medical"]

    # ── Header ──────────────────────────────────────────────────────────
    group_name = form.get("group_name", "")
    ws["C3"].value = group_name
    ws["C3"].protection = Protection(locked=False)
    ws["C3"].alignment = Alignment(horizontal="left", vertical="center",
                                    wrap_text=False, shrink_to_fit=False)

    eff = form.get("effective_date", "")
    qdate = form.get("quote_date", str(datetime.date.today()))
    sv(ws, 5, 9, f"Effective Date: {eff}   |   Quote Date: {qdate}")

    # ── Enrollment ───────────────────────────────────────────────────────
    tier_map = [("ee", 34), ("es", 35), ("ec", 36), ("ef", 37)]
    for key, row in tier_map:
        count = safe_float(form.get(f"enroll_{key}", "0"))
        for col in [4, 5, 6]:
            sv(ws, row, col, count)

    # ── Carriers ─────────────────────────────────────────────────────────
    for gi in range(5):
        carrier_name = form.get(f"g{gi}_carrier", "").strip()
        if not carrier_name:
            continue

        c1, c2, c3 = GROUP_COLS[gi]
        sv(ws, 9, c1, carrier_name)
        sv(ws, 53, c1, form.get(f"g{gi}_commissions", ""))

        # First non-empty plan's comments become the carrier comment
        for pi in range(3):
            if form.get(f"g{gi}_p{pi}_plan_name", "").strip():
                sv(ws, 39, c1, form.get(f"g{gi}_p{pi}_comments", ""))
                break

        plan_cols = [c1, c2, c3]
        for pi in range(3):
            col = plan_cols[pi]
            plan_name = form.get(f"g{gi}_p{pi}_plan_name", "").strip()
            if not plan_name:
                continue

            sv(ws, 10, col, plan_name)

            for field, row in ROW_MAP.items():
                raw = form.get(f"g{gi}_p{pi}_{field}", "")
                if field.startswith("rate_") and raw:
                    val = safe_float(raw)
                    sv(ws, row, col, val, '"$"#,##0.00')
                else:
                    sv(ws, row, col, raw)

    # ── Stream to bytes ───────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, group_name


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    buf, group_name = generate_excel(request.form)
    safe_name = "".join(c for c in group_name if c.isalnum() or c in " _-").strip()
    filename = f"{safe_name or 'Medical_Quote'}_Comparison.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
