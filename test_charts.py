import sys
sys.path.insert(0, "scripts")
from openpyxl import Workbook
from excel_charts import add_chart, save_workbook

wb = Workbook()

# --- Sheet 1: line chart, countries, percentages (one decimal)
ws = wb.active
ws.title = "Line"
years = list(range(2008, 2021))            # 13 points
ws["A1"] = "Year"
ws["B1"], ws["C1"], ws["D1"], ws["E1"] = "Chile", "Brasil", "OCDE", "Japon"
for i, y in enumerate(years, start=2):
    ws.cell(row=i, column=1, value=y)
    ws.cell(row=i, column=2, value=40 + i * 2)
    ws.cell(row=i, column=3, value=55 + i * 1)
    ws.cell(row=i, column=4, value=70 + (i % 4))
    ws.cell(row=i, column=5, value=30 + i * 1.5)
add_chart(ws, "line", title="PISA score by country, 2008-2020",
          min_col=2, max_col=5, min_row=1, max_row=14,
          cats_col=1, cats_min_row=2, cats_max_row=14,
          anchor="H2", is_countries=True, value_kind="percent")

# --- Sheet 2: bar chart, non-country series, currency (no decimals)
ws2 = wb.create_sheet("Bar")
cats = ["North", "South", "East", "West"]
ws2["A1"] = "Region"
ws2["B1"], ws2["C1"] = "2023", "2024"
for i, c in enumerate(cats, start=2):
    ws2.cell(row=i, column=1, value=c)
    ws2.cell(row=i, column=2, value=120 + i * 10)
    ws2.cell(row=i, column=3, value=140 + i * 8)
add_chart(ws2, "bar", title="Revenue by region",
          min_col=2, max_col=3, min_row=1, max_row=5,
          cats_col=1, cats_min_row=2, cats_max_row=5,
          anchor="F2", is_countries=False, value_kind="currency")

# --- Sheet 3: single-series bar chart with countries as categories
#     -> each bar (and its label) takes its own country color
ws3 = wb.create_sheet("BarCountries")
ws3.append(["País", "Procedimientos"])
for row in [("Brasil", 14), ("Venezuela", 10), ("Uruguay", 9), ("LATAM", 8),
            ("Chile", 6), ("Perú", 6)]:
    ws3.append(row)
add_chart(ws3, "bar", title="Procedimientos para registrar propiedad (2019)",
          min_col=2, max_col=2, min_row=1, max_row=7,
          cats_col=1, cats_min_row=2, cats_max_row=7,
          anchor="D2", value_kind="currency")

out = "sample.xlsx"
save_workbook(wb, out)
print(f"saved {out}")
