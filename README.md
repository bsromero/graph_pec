# graph_pec — Progreso en Cifras house charts

A [Claude](https://claude.com/claude-code) **Skill** for building native, editable Excel
charts that follow a fixed house style — line and bar charts with specific colors, fonts,
axes, gridlines, and data-label rules.

The charts are **real Excel chart objects** (not images), so they stay fully editable in
Excel. The whole point of the skill is that every chart comes out identical to the house
style, automatically.

## What it does

- **Line charts** — 5 pt lines, last-datapoint labels (value + series name) distributed so
  they don't overlap, with leader lines back to each point.
- **Bar charts** — 25% gap width, horizontal value-only labels.
- **Country color table** — consistent, named colors whenever countries appear, whether
  as series or as the bars of a single-series bar chart (each bar in its country's color).
- **House typography & axes** — Gotham 20 pt throughout, black 0.75 pt axes, no gridlines,
  no legend, x labels vertical (270°), no decimals for currency / one decimal for percentages.

See [`SKILL.md`](SKILL.md) for the authoritative style spec and full usage.

## Repository layout

```
graph_pec/
├── SKILL.md              # Skill definition + house-style spec (start here)
├── scripts/
│   └── excel_charts.py   # The chart-building helper (add_chart, save_workbook)
└── test_charts.py        # Runnable example that produces sample.xlsx
```

## Install as a Claude Code Skill

Clone this repo into your Claude skills directory:

```bash
git clone https://github.com/bsromero/graph_pec.git \
  ~/.claude/skills/graph_pec
```

Claude Code will pick it up automatically. Ask Claude to "chart this data in the house
style" on any `.xlsx` file and the skill triggers.

## Use the helper directly

Requires [`openpyxl`](https://openpyxl.readthedocs.io/) (`pip install openpyxl`).

```python
import sys; sys.path.insert(0, "scripts")
from excel_charts import add_chart, save_workbook
from openpyxl import load_workbook

wb = load_workbook("file.xlsx")
ws = wb["Sheet1"]
add_chart(
    ws, "line", title="My title",
    min_col=2, max_col=5, min_row=1, max_row=14,   # data incl. header row
    cats_col=1, cats_min_row=2, cats_max_row=14,   # category labels
    anchor="H2",                                    # empty columns to the right
    is_countries=True,
    value_kind="percent",                           # or "currency"
)
save_workbook(wb, "file.xlsx")   # saves AND distributes line labels + leader lines
```

Always save with `save_workbook(...)` rather than `wb.save(...)` — it applies the
line-chart label distribution and leader lines via XML post-processing that openpyxl can't
do on its own.

## Try the example

```bash
pip install openpyxl
python test_charts.py   # writes sample.xlsx with a line chart and a bar chart
```

> **Note:** Gotham must be installed on the viewer's machine to display; otherwise Excel
> substitutes a fallback font. Also, LibreOffice does not render manually-positioned data
> labels or leader lines, so verify label layout in actual Excel.

## License

MIT — see [`LICENSE`](LICENSE).
