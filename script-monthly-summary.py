import smartsheet
import pandas as pd
from datetime import datetime
from jinja2 import Environment
from weasyprint import HTML

# --- MONTH INPUT ---
user_input = input("What month would you like to filter before? (Format: yyyy-mm): ")
try:
    datetime.strptime(user_input, "%Y-%m")
    MONTH_FILTER = user_input
except ValueError:
    print("❌ Invalid format. Use yyyy-mm (e.g., 2025-06).")
    exit(1)

# --- CONFIGURATION ---
API_TOKEN = "v6mG3HaOMr1VjYDdTAdvSyIUmTOK2F3SlBBiV"
SHEET_ID = 5933665067421572
TEMP_HTML = f"summary_temp_{MONTH_FILTER}.html"
OUTPUT_PDF = f"Smartsheet_Summary_{MONTH_FILTER}.pdf"
LOGO_PATH = "file:///C:/Users/lpagan.FR/OneDrive - FR Construction Group/Attachments/logo.png"

# --- SMARTSHEET CONNECTION ---
client = smartsheet.Smartsheet(API_TOKEN)
sheet = client.Sheets.get_sheet(SHEET_ID)
columns = {col.title: col.id for col in sheet.columns}

# --- EXTRACT ROWS ---
rows = []
for row in sheet.rows:
    row_data = {}
    for cell in row.cells:
        col_name = next((k for k, v in columns.items() if v == cell.column_id), None)
        row_data[col_name] = cell.value
    rows.append(row_data)

df = pd.DataFrame(rows)

# --- HELPERS ---
def is_before_month(date, target_month):
    try:
        return pd.to_datetime(date).strftime('%Y-%m') < target_month
    except:
        return False

# --- FILTER CASES ---
filtered_cases = []

for _, row in df.iterrows():
    award_type = str(row.get("Award Type Equivalent", "")).strip().lower()

    ntp_payment = 0
    structure_payment = 0
    final_payment = 0
    include = False

    ntp_date = ""
    structure_date = ""
    final_date = ""

    if award_type in ["relocation", "repair"]:
        base_price = 65000

        ntp_raw = row.get("Date of Notice to Proceed")
        final_raw = row.get("Relo or Repair Final Inspection Passed")

        if pd.notnull(ntp_raw) and is_before_month(ntp_raw, MONTH_FILTER):
            ntp_payment = base_price * 0.5
            ntp_date = ntp_raw
            include = True

        if pd.notnull(final_raw) and is_before_month(final_raw, MONTH_FILTER):
            final_payment = base_price * 0.5
            final_date = final_raw
            include = True

    elif award_type == "reconstruction":
        ntp_raw = row.get("Date of Notice to Proceed")
        structure_raw = row.get("Structure Inspection Passed")
        final_raw = row.get("Final Inspection Passed")

        if pd.notnull(ntp_raw) and is_before_month(ntp_raw, MONTH_FILTER):
            ntp_payment = row.get("Payment Notice to Proceed") or 0
            ntp_date = ntp_raw
            include = True

        if pd.notnull(structure_raw) and is_before_month(structure_raw, MONTH_FILTER):
            structure_payment = row.get("Payment Structure") or 0
            structure_date = structure_raw
            include = True

        if pd.notnull(final_raw) and is_before_month(final_raw, MONTH_FILTER):
            final_payment = row.get("Payment Final") or 0
            final_date = final_raw
            include = True

    if include:
        case_info = {
            "CaseID": row.get("Case ID"),
            "Type": award_type.capitalize(),
            "NTP": ntp_date,
            "NTP_Payment": ntp_payment,
            "Structure": structure_date,
            "Structure_Payment": structure_payment,
            "Final": final_date,
            "Final_Payment": final_payment
        }
        filtered_cases.append(case_info)

# --- SUMMARY ---
date_based_summary = {
    "NTP": sum(c["NTP_Payment"] for c in filtered_cases),
    "Structure": sum(c["Structure_Payment"] for c in filtered_cases),
    "Final": sum(c["Final_Payment"] for c in filtered_cases),
    "Total": sum((c["NTP_Payment"] or 0) + (c["Structure_Payment"] or 0) + (c["Final_Payment"] or 0) for c in filtered_cases)
}

# --- HTML TEMPLATE ---
html_template = """
<!DOCTYPE html>
<html>
<head>
    <meta charset=\"UTF-8\">
    <title>Certification 2025-01</title>
    <style>
        body { font-family: Arial, sans-serif; font-size: 10px; }
        table { border-collapse: collapse; width: 100%; font-size: 9px; }
        th, td { border: 1px solid black; padding: 4px; text-align: center; }
        h1, h2 { font-size: 12px; }
    </style>
</head>
<body>
<img src='""" + LOGO_PATH + """' style=\"width:150px; margin-bottom:10px;\">
<h1>Certification No. 2025-01 – May 2025 – All certified cases up to May 31, 2025</h1>

<h2>Individual Cases</h2>
<table>
    <thead>
        <tr>
            <th>Case ID</th><th>Type</th>
            <th>NTP Date</th><th>NTP Payment</th>
            <th>Structure Date</th><th>Structure Payment</th>
            <th>Final Date</th><th>Final Payment</th>
        </tr>
    </thead>
    <tbody>
    {% for c in cases %}
        <tr>
            <td>{{ c.CaseID }}</td><td>{{ c.Type }}</td>
            <td>{{ c.NTP }}</td><td>${{ "%.2f"|format(c.NTP_Payment or 0) }}</td>
            <td>{{ c.Structure }}</td><td>${{ "%.2f"|format(c.Structure_Payment or 0) }}</td>
            <td>{{ c.Final }}</td><td>${{ "%.2f"|format(c.Final_Payment or 0) }}</td>
        </tr>
    {% endfor %}
    </tbody>
</table>
<br>
<h2>Summary Totals:</h2>
<ul>
    <li><strong>NTP:</strong> ${{ date_summary.NTP }}</li>
    <li><strong>Structure:</strong> ${{ date_summary.Structure }}</li>
    <li><strong>Final:</strong> ${{ date_summary.Final }}</li>
    <li><strong style=\"color: green;\">Total:</strong> <strong>${{ date_summary.Total }}</strong></li>
</ul>
</body>
</html>
"""

# --- GENERATE PDF ---
with open(TEMP_HTML, "w", encoding="utf-8") as f:
    f.write(Environment().from_string(html_template).render(
        month=MONTH_FILTER,
        date_summary=date_based_summary,
        cases=filtered_cases
    ))

HTML(TEMP_HTML).write_pdf(OUTPUT_PDF)
print(f"✅ PDF successfully generated: {OUTPUT_PDF}")
