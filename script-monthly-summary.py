#!/usr/bin/env python3
"""
Smartsheet → HTML + PDF + Excel summary hasta 2025-05-31
• Excluye ciertos Stage Status
• Selecciona filas en que NTP, Structure o Final estén entre 2024-10-23 y 2025-05-31 (OR)
• Relocation/Repair: monto fijo $65 000 dividido 50/50/0
• Reconstruction: lee montos reales de las columnas
• Salidas:
    – summary_temp_<YYYY-MM-DD>.html
    – Smartsheet_Summary_until2025-05-31.pdf
    – Smartsheet_Summary_until2025-05-31.xlsx
"""

import os, smartsheet, pandas as pd
from datetime import datetime, date
from weasyprint import HTML
from pandas import ExcelWriter
import openpyxl  # requerido por pandas

# Mostrar TODO en pandas si debugueas
pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", None)

# ─────────────────── CONFIG ───────────────────
API_TOKEN = os.getenv("SMARTSHEET_ACCESS_TOKEN") or "v6mG3HaOMr1VjYDdTAdvSyIUmTOK2F3SlBBiV"
SHEET_ID  = 5933665067421572

INICIO    = pd.to_datetime("2024-10-23")
FIN       = pd.to_datetime("2025-05-31")
HOY       = datetime.today().strftime("%Y-%m-%d")

TMP_HTML  = f"summary_temp_{HOY}.html"
OUT_PDF   = "Smartsheet_Summary_until2025-05-31.pdf"
OUT_XLSX  = "Smartsheet_Summary_until2025-05-31.xlsx"
LOGO_PATH = "fr_logo.png"


# ─────────── CONECTAR Y BAJAR HASTA 10 000 FILAS ───────────
client = smartsheet.Smartsheet(API_TOKEN)
sheet  = client.Sheets.get_sheet(SHEET_ID, page_size=10000, page=1)

col_map = {c.id: c.title.strip().replace("\n"," ") for c in sheet.columns}
records = [
    {col_map[cell.column_id]: cell.value for cell in row.cells}
    for row in sheet.rows
]
df = pd.DataFrame(records)
df.columns = [c.strip().replace("\n"," ") for c in df.columns]


# ─────────── FUNCIONES AUXILIARES ───────────
def parse_fecha(v):
    if pd.isna(v): return pd.NaT
    if isinstance(v, (datetime, pd.Timestamp, date)):
        return pd.to_datetime(v)
    for fmt in ("%m/%d/%y","%m/%d/%Y","%Y-%m-%d","%b/%d/%y","%d/%m/%Y"):
        try:
            return datetime.strptime(str(v), fmt)
        except ValueError:
            pass
    return pd.to_datetime(v, errors="coerce")

to_money = lambda x: float(pd.to_numeric(x, errors="coerce") or 0)


# ─────────── FILTRAR Y CALCULAR ───────────
EXCLUDE = {
    "01 - Initial Scoping",
    "03 - Design & Permitting",
    "00 - Reassigned",
    "02 - Pending Task Order",
    "00 - Assigned Offline",
    "16 - Inactive",
}

casos = []
for _, row in df.iterrows():
    status = str(row.get("Stage Status","")).strip()
    if status in EXCLUDE:
        continue

    f_ntp = parse_fecha(row.get("Date of Notice to Proceed"))
    f_str = parse_fecha(row.get("Structure Inspection Passed"))
    f_fin = parse_fecha(row.get("Final Inspection Passed"))
    if pd.isna(f_fin):
        f_fin = parse_fecha(row.get("Relo or Repair Final Inspection Passed"))

    # OR entre las tres fechas
    if not any([
        pd.notna(f_ntp) and INICIO <= f_ntp <= FIN,
        pd.notna(f_str) and INICIO <= f_str <= FIN,
        pd.notna(f_fin) and INICIO <= f_fin <= FIN,
    ]):
        continue

    tipo  = str(row.get("Award Type Equivalent","")).strip().lower()
    pagos = {"ntp":0,"structure":0,"final":0}
    fechas = {"ntp":"","structure":"","final":""}
    include = False

    # Relocation/Repair: $65 000 split 50/50 (no Structure final)
    if tipo in ("relocation","repair"):
        base = 65000
        if pd.notna(f_ntp) and INICIO <= f_ntp <= FIN:
            pagos["ntp"]      = base*0.5
            fechas["ntp"]     = f_ntp.date()
            include = True
        if pd.notna(f_str) and INICIO <= f_str <= FIN:
            pagos["structure"]   = base*0.5
            fechas["structure"] = f_str.date()
            include = True
        if pd.notna(f_fin) and INICIO <= f_fin <= FIN:
            pagos["final"]    = base*0.5
            fechas["final"]   = f_fin.date()
            include = True

    # Reconstruction: usa datos de columna real
    elif tipo == "reconstruction":
        if pd.notna(f_ntp) and INICIO <= f_ntp <= FIN:
            pagos["ntp"]      = to_money(row.get("Payment Notice to Proceed"))
            fechas["ntp"]     = f_ntp.date()
            include = True
        if pd.notna(f_str) and INICIO <= f_str <= FIN:
            pagos["structure"]   = to_money(row.get("Payment Structure"))
            fechas["structure"] = f_str.date()
            include = True
        if pd.notna(f_fin) and INICIO <= f_fin <= FIN:
            pagos["final"]    = to_money(row.get("Payment Final"))
            fechas["final"]   = f_fin.date()
            include = True

    if include:
        casos.append({
            "Case ID":          row.get("Case ID"),
            "Type":             tipo.capitalize(),
            "NTP Date":         fechas["ntp"],
            "NTP $":            pagos["ntp"],
            "Structure Date":   fechas["structure"],
            "Structure $":      pagos["structure"],
            "Final Date":       fechas["final"],
            "Final $":          pagos["final"],
            "Total Payments":   pagos["ntp"] + pagos["structure"] + pagos["final"],
        })


# ─────────── TOTALES ───────────
tot = {
    "NTP":       sum(c["NTP $"] for c in casos),
    "Structure": sum(c["Structure $"] for c in casos),
    "Final":     sum(c["Final $"] for c in casos),
}
tot["Total"] = sum(tot.values())


# ─────────── EXPORTAR EXCEL ───────────
df_out = pd.DataFrame(casos)
df_out.to_excel(OUT_XLSX, index=False)
print(f"✅ Excel generado: {OUT_XLSX}")

# ─────────── GENERAR PDF ───────────
html_rows = "\n".join(
    f"<tr>"
    f"<td>{c['Case ID']}</td><td>{c['Type']}</td>"
    f"<td>{c['NTP Date']}</td><td>${c['NTP $']:,.2f}</td>"
    f"<td>{c['Structure Date']}</td><td>${c['Structure $']:,.2f}</td>"
    f"<td>{c['Final Date']}</td><td>${c['Final $']:,.2f}</td>"
    f"<td>${c['Total Payments']:,.2f}</td>"
    f"</tr>"
    for c in casos
)

html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
  thead{{display:table-header-group}} tr{{page-break-inside:avoid}}
  @page{{size:A4;margin:10mm}}
  body{{font-family:Arial;font-size:10px}}
  table{{border-collapse:collapse;width:100%}}
  th,td{{border:1px solid #000;padding:4px;text-align:center}}
  h1,h2{{font-size:12px;margin:4px 0}}
</style></head><body>
<img src="{LOGO_PATH}" style="width:120px">
<h1>Resumen hasta 2025-05-31 — {len(casos)} casos</h1>
<table>
  <thead><tr>
    <th>Case ID</th><th>Type</th>
    <th>NTP Date</th><th>NTP $</th>
    <th>Structure Date</th><th>Structure $</th>
    <th>Final Date</th><th>Final $</th>
    <th>Total Payments</th>
  </tr></thead>
  <tbody>{html_rows}</tbody>
  <tfoot><tr>
    <td colspan="3"><strong>Totals:</strong></td>
    <td><strong>${tot['NTP']:,.2f}</strong></td><td></td>
    <td><strong>${tot['Structure']:,.2f}</strong></td><td></td>
    <td><strong>${tot['Final']:,.2f}</strong></td>
    <td><strong>${tot['Total']:,.2f}</strong></td>
  </tr></tfoot>
</table>
</body></html>"""

with open(TMP_HTML, "w", encoding="utf-8") as f:
    f.write(html)
HTML(TMP_HTML).write_pdf(OUT_PDF)
print(f"✅ PDF generado: {OUT_PDF}")
