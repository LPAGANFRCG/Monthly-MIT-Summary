import smartsheet
import pandas as pd
from datetime import datetime
from jinja2 import Environment, FileSystemLoader
import pdfkit
import os

# --- CONFIGURACIÓN ---
API_TOKEN = "v6mG3HaOMr1VjYDdTAdvSyIUmTOK2F3SlBBiV"
SHEET_ID = 5933665067421572
MONTH_FILTER = "2025-05"  # Cambia esto si deseas otro mes
TEMP_HTML = r"C:\Users\lpagan.FR\OneDrive - FR Construction Group\Attachments\Desktop\summary_temp.html"
OUTPUT_PDF = rf"C:\Users\lpagan.FR\OneDrive - FR Construction Group\Attachments\Desktop\Smartsheet_Summary_{MONTH_FILTER}.pdf"


# --- CONEXIÓN A SMARTSHEET ---
client = smartsheet.Smartsheet(API_TOKEN)
sheet = client.Sheets.get_sheet(SHEET_ID)
columns = {col.title: col.id for col in sheet.columns}

# --- EXTRACCIÓN DE FILAS ---
rows = []
for row in sheet.rows:
    row_data = {}
    for cell in row.cells:
        col_name = next((k for k, v in columns.items() if v == cell.column_id), None)
        row_data[col_name] = cell.value
    rows.append(row_data)

df = pd.DataFrame(rows)

# --- FUNCIONES AUXILIARES ---
def get_period(date):
    if pd.notnull(date):
        try:
            return pd.to_datetime(date).strftime('%Y-%m')
        except:
            return None
    return None

df["NTP Period"] = df["NTP Date Cleaned"].apply(get_period)
df["Structure Period"] = df["Structure Inspection Passed"].apply(get_period)
df["Final Period"] = df["Final Inspection Passed"].apply(get_period)

# --- FILTRAR CASOS DEL MES ---
filtered_cases = []

for index, row in df.iterrows():
    case_info = {
        "CaseID": row.get("Case ID"),
        "MIT": row.get("Case Number (MIT Invoicing)"),
        "Model": row.get("Model Home Type"),
        "Price": row.get("F&R Contract Price"),
        "NTP": row.get("NTP Date Cleaned"),
        "Structure": row.get("Structure Inspection Passed"),
        "Final": row.get("Final Inspection Passed"),
        "NTP_Payment": row.get("Payment NTP") if row.get("NTP Period") == MONTH_FILTER else 0,
        "Structure_Payment": row.get("Payment Structure") if row.get("Structure Period") == MONTH_FILTER else 0,
        "Final_Payment": row.get("Payment Final") if row.get("Final Period") == MONTH_FILTER else 0,
        "AddOn": row.get("Demo/Relocation Add-on") if row.get("NTP Period") == MONTH_FILTER else 0
    }
    if any([case_info["NTP_Payment"], case_info["Structure_Payment"], case_info["Final_Payment"], case_info["AddOn"]]):
        filtered_cases.append(case_info)

# --- RESUMEN POR TIPO ---
summary_totals = {
    "NTP": sum(c["NTP_Payment"] or 0 for c in filtered_cases),
    "Structure": sum(c["Structure_Payment"] or 0 for c in filtered_cases),
    "Final": sum(c["Final_Payment"] or 0 for c in filtered_cases),
    "AddOn": sum(c["AddOn"] or 0 for c in filtered_cases)
}

# --- CREAR HTML PARA PDF ---
env = Environment(loader=FileSystemLoader("/mnt/data"))
html_template = """
<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"><title>Resumen Mensual</title></head>
<body>
<h1>Resumen Mensual de Certificaciones - {{ month }}</h1>
<h2>Totales por tipo</h2>
<ul>
    <li><strong>NTP:</strong> ${{ totals.NTP }}</li>
    <li><strong>Add-on:</strong> ${{ totals.AddOn }}</li>
    <li><strong>Structure:</strong> ${{ totals.Structure }}</li>
    <li><strong>Final:</strong> ${{ totals.Final }}</li>
</ul>
<h2>Casos individuales</h2>
<table border="1" cellspacing="0" cellpadding="5">
    <thead>
        <tr><th>Case ID</th><th>MIT</th><th>Model</th><th>Price</th><th>NTP</th><th>Structure</th><th>Final</th><th>NTP $</th><th>Add-on $</th><th>Structure $</th><th>Final $</th></tr>
    </thead>
    <tbody>
    {% for c in cases %}
        <tr>
            <td>{{ c.CaseID }}</td><td>{{ c.MIT }}</td><td>{{ c.Model }}</td><td>{{ c.Price }}</td>
            <td>{{ c.NTP }}</td><td>{{ c.Structure }}</td><td>{{ c.Final }}</td>
            <td>${{ "%.2f"|format(c.NTP_Payment or 0) }}</td>
            <td>${{ "%.2f"|format(c.AddOn or 0) }}</td>
            <td>${{ "%.2f"|format(c.Structure_Payment or 0) }}</td>
            <td>${{ "%.2f"|format(c.Final_Payment or 0) }}</td>
        </tr>
    {% endfor %}
    </tbody>
</table>
</body>
</html>
"""

# Guardar HTML temporal
with open(TEMP_HTML, "w") as f:
    f.write(Environment().from_string(html_template).render(
        month=MONTH_FILTER,
        totals=summary_totals,
        cases=filtered_cases
    ))

# Convertir a PDF
pdfkit.from_file(TEMP_HTML, OUTPUT_PDF)

OUTPUT_PDF
