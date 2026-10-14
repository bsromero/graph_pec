---
name: graph_pec
description: "Build native, editable Excel charts that follow our fixed house style (line and bar charts with specific colors, fonts, axes, gridlines, data-label rules; no legend). Use this skill WHENEVER the user asks to add, create, insert, or restyle a chart or graph in an .xlsx file — including phrases like 'graph this', 'chart the data', 'add a line/bar chart', 'plot these countries', or any request to visualize spreadsheet data as a chart. Always use this skill for Excel charts even if the user does not explicitly mention the house style, because the styling is the default expectation. Do NOT use for image/matplotlib charts, Google Sheets, or non-chart spreadsheet work."
---

# graph_pec — Progreso en Cifras house charts

Build **native** Excel charts (real chart objects, editable in Excel) that match a fixed
house style. Never substitute image charts unless the user explicitly asks for one — the
point of this skill is that every chart comes out identical to the rules below.

## How to use

1. Identify the data block in the worksheet: the columns holding series, the header row
   with series names, and the column holding category-axis labels.
2. Decide `chart_type` (`"line"` or `"bar"`) and whether the series are **countries**
   (set `is_countries=True` to apply the country color table). If it's a bar chart with
   **one series and countries on the X axis**, every bar gets its country color — the
   helper detects this on its own (or force it with `color_by="category"`).
3. Get the title from the user — they always supply it.
4. Decide the number format: `value_kind="currency"` (no decimals) or
   `value_kind="percent"` (one decimal). If omitted, the helper infers it from the
   data cells' number format (`%` → percent, otherwise no decimals).
5. Call `add_chart(...)` from `scripts/excel_charts.py`. Place the chart in empty columns
   to the right of the data.
6. Save with **`save_workbook(wb, path)`** (not `wb.save`). This applies the line-chart
   label distribution and leader lines, which require XML post-processing openpyxl can't do.

```python
import sys; sys.path.insert(0, "<path-to-this-skill>/scripts")
from excel_charts import add_chart, save_workbook
from openpyxl import load_workbook

wb = load_workbook("file.xlsx")
ws = wb["Sheet1"]
add_chart(
    ws, "line", title="My title supplied by user",
    min_col=2, max_col=5, min_row=1, max_row=14,   # data incl. header row
    cats_col=1, cats_min_row=2, cats_max_row=14,   # category labels
    anchor="H2",                                    # empty columns to the right
    is_countries=True,
    value_kind="percent",                           # or "currency"
)
save_workbook(wb, "file.xlsx")     # saves AND distributes line labels + leader lines
```

The helper resolves colors, fonts, axis styling, legend removal, and data labels automatically.
Read `scripts/excel_charts.py` if you need to extend it.

> **Previewing:** LibreOffice does **not** render manually-positioned data labels or the
> leader-line extension (it drops them), so a LibreOffice/PDF preview will not show the
> label distribution correctly. Verify label layout in actual Excel. The XML written matches
> what Excel itself produces for these features.

## The house style (authoritative spec)

**Chart types**
- Line chart: line width **5 pt**.
- Bar chart: **gap width 25%**.
- Every series (line or bar) gets a different color.

**Colors — countries.** Whenever countries appear in a chart, use the exact color table in
`scripts/excel_charts.py` (`_COUNTRY_TABLE`). If a country is present but not in the
table, assign a color distinct from every color already used. This applies **both** when
countries are the series (lines or grouped bars) **and** when countries are the
categories of a single-series bar chart (e.g. "Procedimientos para registrar propiedad"
with one bar per country): **each bar takes its own country's color** (Chile red, Brasil
purple, LATAM black, …) and its value label takes the same color. Never paint all the
bars one flat color just because they belong to one series. The helper does this
automatically (`color_by="category"` is inferred when there is one series and the
category labels are countries); pass `color_by` explicitly to override.

**Colors — non-countries.** Use this priority palette in order, then pick further colors
that contrast strongly with those already used:
`#ff0000, #0070C0, #C00000, #8AB3FF, #e65b44, #666666`.

**Fonts.** Everything — title, axis labels, legend, data labels — is **Gotham regular,
size 20, never bold**. (Gotham must be installed on the viewer's machine to display;
otherwise Excel substitutes a fallback. The setting is always written into the file.)

**Title.** The user always provides the title text. Gotham regular size 20, not bold.

**Axes.** Both X and Y: line color black (#000000), weight **0.75 pt**. No axis titles.
No tick marks. Labels in Gotham regular size 20, black, not bold, **horizontal (0°) by
default — even for bar charts**. Number format: **no decimals for currency, one decimal
for percentages**. For the X-axis: **always** rotate tick labels to **270° (vertical)**
and set label position to **low**.

**Gridlines.** None.

**Legend.** **No legend.** (Line charts identify series through the last-point labels.)

**Data labels**
- Line charts: label **only the latest datapoint** of each series, showing **value +
  series name**, in the series color. Labels must be **distributed so they don't overlap**
  — the helper spreads clustered last-point labels vertically (leaving well-separated ones
  in place) and turns on **leader lines colored #A6A6A6** so each moved label still points
  back to its datapoint.
- Bar charts: show **only the value**, in the series color, **horizontal (0°)**.

**Size & placement.** Chart size **8 cm height × 10.6 cm width** (the instructions
file says "pt", but 8 × 10.6 pt would be a postage stamp — cm is the intended unit). Place the chart in
columns that are completely empty, usually just to the right of the data.

## Notes for extending
- Country names are matched case- and accent-insensitively, with a few common aliases
  (USA, EU, OECD, etc.). If a header won't match, add it to `_ALIASES` or `_COUNTRY_TABLE`.
- Label distribution is tuned by `_PLOT_V_SPAN`, `_LABEL_GAP`, `_LABEL_LO`, `_LABEL_HI`
  (fractions of the chart area). Increase `_LABEL_GAP` if labels still feel tight.
- To support more chart types (stacked bar, area), extend `add_chart` following the same
  styling helpers (`_style_axis`, `_rich`, `_bar_labels`, `_line_labels`).
