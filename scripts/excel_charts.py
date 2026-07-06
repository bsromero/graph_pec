"""
excel_charts.py — Build native (editable) Excel charts to a fixed house style.

Public entry point: add_chart(...). See SKILL.md for the spec this implements.
All charts are real openpyxl chart objects, so they stay editable in Excel.
"""
from openpyxl.chart import LineChart, BarChart, Reference, Series
from openpyxl.chart.title import Title
from openpyxl.chart.text import Text, RichText
from openpyxl.chart.label import DataLabelList, DataLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (
    Paragraph, ParagraphProperties, CharacterProperties, Font as DrawingFont,
    RichTextProperties, RegularTextRun,
)
import unicodedata
import re
import zipfile
import shutil
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

# ---- House style constants (from the spec) -------------------------------
FONT = "Gotham"
SIZE = 2000            # 20 pt, in 1/100 pt
BOLD = False
EMU_PER_PT = 12700
LINE_WIDTH_EMU = 5 * EMU_PER_PT       # 5 pt lines
AXIS_LINE_EMU = int(0.75 * EMU_PER_PT)  # 0.75 pt axis lines
BAR_GAP = 25                          # gap width %
CHART_H_CM = 8.0
CHART_W_CM = 10.6
ROT_VERT = -5400000                   # 90deg clockwise, in 1/60000 deg
NUMFMT = "0"                          # no decimals
ROTATE_THRESHOLD = 10                 # >10 datapoints -> vertical x labels
LEADER_COLOR = "A6A6A6"               # leader-line color
# Label-distribution tuning (fractions of chart area):
_PLOT_V_SPAN = 0.62   # approx. plot-area height as a fraction of the chart
_LABEL_GAP = 0.05     # minimum vertical gap between adjacent labels
_LABEL_LO = -0.04     # keep labels within this band of the plot
_LABEL_HI = 0.66

# Country -> color (hex, no '#'). Keys stored accent/space-insensitive.
_COUNTRY_TABLE = {
    "Argentina": "57b7ce", "Australia": "ffb300", "Bolivia": "047004",
    "Brasil": "a020f0", "Chile": "ff0000", "Colombia": "e45680",
    "Corea del Sur": "003478", "Ecuador": "e2a723", "España": "aa151b",
    "Estados Unidos": "0a3161", "Estonia": "0072ce",
    "Irlanda": "169b62", "LATAM": "000000", "México": "0fbc28",
    "Mundo": "d8c596", "Nueva Zelanda": "0000ff", "OCDE": "00aae6",
    "Perú": "aa0e1d", "Polonia": "dc143c", "Portugal": "006600",
    "Suiza": "997950", "Turquía": "e30a17", "Unión Europea": "ffcc00",
    "Uruguay": "9e9e9e", "Venezuela": "de7c21",
}
# Common aliases people actually type -> canonical key above.
_ALIASES = {
    "usa": "Estados Unidos", "ee uu": "Estados Unidos",
    "eeuu": "Estados Unidos", "estados unidos (ee uu)": "Estados Unidos",
    "united states": "Estados Unidos", "korea": "Corea del Sur",
    "south korea": "Corea del Sur", "oecd": "OCDE", "oecd (ocde)": "OCDE",
    "ue": "Unión Europea", "eu": "Unión Europea", "ue (union europea)": "Unión Europea",
    "brazil": "Brasil", "mexico": "México", "peru": "Perú",
    "spain": "España", "world": "Mundo",
}
DEFAULT_PALETTE = ["ff0000", "0070C0", "C00000", "8AB3FF", "e65b44", "666666"]
# Extra distinct colors to auto-assign when the priority lists run out.
_EXTRA = ["00B050", "7030A0", "FFC000", "00B0F0", "ED7D31", "264478",
          "843C0C", "BF9000", "538135", "C45911", "203864", "7B7B7B"]


def _norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return " ".join(s.lower().replace(".", " ").split())


_COUNTRY_NORM = {_norm(k): v for k, v in _COUNTRY_TABLE.items()}


def _next_distinct(used):
    for c in DEFAULT_PALETTE + _EXTRA:
        if c.upper() not in used:
            return c
    # last resort: deterministic spin through the color wheel
    import colorsys
    i = len(used)
    r, g, b = colorsys.hsv_to_rgb((i * 0.37) % 1.0, 0.65, 0.85)
    return f"{int(r*255):02X}{int(g*255):02X}{int(b*255):02X}"


def resolve_colors(names, is_countries):
    """Return a hex color (no '#') for each series name, per the spec rules."""
    used, out = set(), []
    for i, n in enumerate(names):
        c = None
        if is_countries:
            key = _norm(n)
            c = _COUNTRY_NORM.get(key) or _COUNTRY_NORM.get(_norm(_ALIASES.get(key, "")))
        if not c:
            if not is_countries and i < len(DEFAULT_PALETTE) and DEFAULT_PALETTE[i].upper() not in used:
                c = DEFAULT_PALETTE[i]
            else:
                c = _next_distinct(used)
        out.append(c)
        used.add(c.upper())
    return out


# ---- text helpers --------------------------------------------------------
def _cp(color="000000"):
    return CharacterProperties(latin=DrawingFont(typeface=FONT), sz=SIZE, b=BOLD,
                               solidFill=color)


def _rich(color="000000", rot=None):
    body = RichTextProperties(rot=rot) if rot is not None else RichTextProperties()
    cp = _cp(color)
    return RichText(bodyPr=body,
                    p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def _make_title(text):
    cp = _cp("000000")
    rt = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp),
                               r=[RegularTextRun(rPr=cp, t=str(text))])])
    t = Title(tx=Text(rich=rt))
    t.overlay = False
    return t


def _style_axis(axis, rotate=False, is_x=False):
    axis.delete = False
    axis.majorGridlines = None
    axis.minorGridlines = None
    axis.majorTickMark = "none"
    axis.minorTickMark = "none"
    axis.title = None
    axis.spPr = GraphicalProperties(ln=LineProperties(solidFill="000000", w=AXIS_LINE_EMU))
    axis.number_format = NUMFMT
    axis.txPr = _rich("000000", rot=ROT_VERT if rotate else None)
    if is_x and rotate:
        axis.tickLblPos = "low"


def _bar_labels(color):
    dl = DataLabelList()
    for f in ("showVal",):
        setattr(dl, f, True)
    for f in ("showSerName", "showCatName", "showLegendKey", "showPercent", "showBubbleSize"):
        setattr(dl, f, False)
    dl.numFmt = NUMFMT
    dl.txPr = _rich(color, rot=ROT_VERT)
    return dl


def _line_labels(color, last_idx):
    dl = DataLabelList()
    for f in ("showVal", "showSerName", "showCatName", "showLegendKey",
              "showPercent", "showBubbleSize"):
        setattr(dl, f, False)
    pt = DataLabel(idx=last_idx)
    pt.showVal = True
    pt.showSerName = True
    pt.showCatName = False
    pt.showLegendKey = False
    pt.showPercent = False
    pt.showBubbleSize = False
    pt.numFmt = NUMFMT
    pt.separator = " "
    pt.txPr = _rich(color)
    dl.dLbl = [pt]
    return dl


def add_chart(ws, chart_type, title, min_col, max_col, min_row, max_row,
              cats_col, cats_min_row, cats_max_row, anchor,
              is_countries=False, titles_from_data=True):
    """
    Add a styled native chart to worksheet `ws`.

    chart_type : 'line' or 'bar'
    title      : chart title text (always supplied by the user)
    data block : columns min_col..max_col, rows min_row..max_row.
                 If titles_from_data, min_row is the header row holding series names.
    cats_*     : the category-axis label column and its first/last data row.
    anchor     : top-left cell for the chart, e.g. 'H2' (use empty columns).
    is_countries : True to apply the country color table.
    """
    chart_type = chart_type.lower()
    chart = LineChart() if chart_type == "line" else BarChart()
    if chart_type == "bar":
        chart.type = "col"
        chart.gapWidth = BAR_GAP

    data = Reference(ws, min_col=min_col, max_col=max_col,
                     min_row=min_row, max_row=max_row)
    cats = Reference(ws, min_col=cats_col, min_row=cats_min_row, max_row=cats_max_row)
    chart.add_data(data, titles_from_data=titles_from_data)
    chart.set_categories(cats)

    names = [s.tx.strRef.f if s.tx and s.tx.strRef else (s.tx.v if s.tx else f"S{i}")
             for i, s in enumerate(chart.series)]
    # Prefer actual header text for color lookup.
    header_names = []
    for c in range(min_col, max_col + 1):
        header_names.append(ws.cell(row=min_row, column=c).value)
    colors = resolve_colors(header_names, is_countries)

    n_points = cats_max_row - cats_min_row + 1
    rotate = n_points > ROTATE_THRESHOLD

    for s, color in zip(chart.series, colors):
        s.graphicalProperties = GraphicalProperties()
        if chart_type == "line":
            s.graphicalProperties.line = LineProperties(solidFill=color, w=LINE_WIDTH_EMU)
            s.smooth = False
            s.dLbls = _line_labels(color, n_points - 1)
        else:
            s.graphicalProperties.solidFill = color
            s.graphicalProperties.line = LineProperties(solidFill=color)
            s.dLbls = _bar_labels(color)

    chart.title = _make_title(title)
    _style_axis(chart.x_axis, rotate=rotate, is_x=True)
    _style_axis(chart.y_axis, rotate=False, is_x=False)

    chart.legend.position = "b"
    chart.legend.overlay = False
    chart.legend.txPr = _rich("000000")   # Gotham 20 regular, black

    chart.height = CHART_H_CM
    chart.width = CHART_W_CM
    ws.add_chart(chart, anchor)
    return chart


# ===========================================================================
# Line-chart label distribution (post-processing)
# ---------------------------------------------------------------------------
# openpyxl cannot emit per-label manual offsets or the (Microsoft-extension)
# leader lines, so after saving we edit the line-chart XML directly to:
#   * spread the last-point labels vertically so they don't overlap, and
#   * enable leader lines colored LEADER_COLOR that connect each moved label
#     back to its data point.
# Labels that are already well separated are left untouched.
# ===========================================================================

_LEADER_EXT = (
    '<extLst><ext uri="{{CE6537A1-D6FC-4f65-9D91-7224C49458BB}}" '
    'xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">'
    '<c15:showLeaderLines val="1"/><c15:leaderLines>'
    '<spPr><a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    '<a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></spPr>'
    '</c15:leaderLines></ext></extLst>'
).format(color=LEADER_COLOR)


def _parse_ref(f):
    sheet, _, rng = f.rpartition("!")
    sheet = sheet.strip().strip("'").replace("''", "'")
    minc, minr, maxc, maxr = range_boundaries(rng.replace("$", ""))
    return sheet, minc, minr, maxc, maxr


def _series_values(wb, f):
    sheet, minc, minr, maxc, maxr = _parse_ref(f)
    ws = wb[sheet]
    out = []
    for r in range(minr, maxr + 1):
        v = ws.cell(row=r, column=minc).value
        out.append(v if isinstance(v, (int, float)) else None)
    return out


def _declutter(nat, gap, lo, hi):
    """Spread values so neighbours are >= gap apart, recentring clusters on
    their original mean and keeping everything within [lo, hi]."""
    n = len(nat)
    if n <= 1:
        return list(nat)
    order = sorted(range(n), key=lambda i: nat[i])
    p = [nat[i] for i in order]

    def forward():
        for k in range(1, n):
            if p[k] < p[k - 1] + gap:
                p[k] = p[k - 1] + gap

    forward()
    # recentre each contiguous (gap-spaced) cluster on its original mean
    k = 0
    while k < n:
        j = k
        while j + 1 < n and abs(p[j + 1] - (p[j] + gap)) < 1e-6:
            j += 1
        if j > k:
            orig = sum(nat[order[t]] for t in range(k, j + 1)) / (j - k + 1)
            cur = sum(p[k:j + 1]) / (j - k + 1)
            shift = orig - cur
            for t in range(k, j + 1):
                p[t] += shift
        k = j + 1
    forward()  # fix any overlap reintroduced by recentring
    # clamp into [lo, hi]
    if p[0] < lo:
        p = [v + (lo - p[0]) for v in p]
    if p[-1] > hi:
        p = [v - (p[-1] - hi) for v in p]
        if p[0] < lo:  # not enough room: distribute evenly
            p = [lo + (hi - lo) * t / (n - 1) for t in range(n)]
    res = [0.0] * n
    for slot, i in enumerate(order):
        res[i] = p[slot]
    return res


def _process_line_xml(xml, wb):
    sers = re.findall(r"<ser>.*?</ser>", xml, re.S)
    if not sers:
        return xml
    vals, last_idx, last_val, allv = [], [], [], []
    for b in sers:
        m = re.search(r"<val>.*?<f>(.*?)</f>", b, re.S)
        v = _series_values(wb, m.group(1)) if m else []
        li, lv = None, None
        for i in range(len(v) - 1, -1, -1):
            if v[i] is not None:
                li, lv = i, v[i]
                break
        vals.append(v); last_idx.append(li); last_val.append(lv)
        allv += [x for x in v if x is not None]
    if not allv:
        return xml
    span = (max(allv) - min(allv)) or 1.0
    naturals = [None if lv is None else (lv - min(allv)) / span * _PLOT_V_SPAN
                for lv in last_val]
    present = [i for i, n in enumerate(naturals) if n is not None]
    decl = _declutter([naturals[i] for i in present], _LABEL_GAP, _LABEL_LO, _LABEL_HI)
    offset = {}
    for j, i in enumerate(present):
        offset[i] = naturals[i] - decl[j]   # OOXML y: positive = down

    out = xml
    for i, b in enumerate(sers):
        nb = b
        if last_idx[i] is not None:
            layout = ('<layout><manualLayout><x val="0"/>'
                      '<y val="%.6g"/></manualLayout></layout>' % offset[i])
            nb = re.sub(r'(<dLbl><idx val="%d"/>)' % last_idx[i],
                        r"\1" + layout, nb, count=1)
        if "CE6537A1-D6FC-4f65-9D91-7224C49458BB" not in nb:
            nb = nb.replace("</dLbls>", _LEADER_EXT + "</dLbls>", 1)
        out = out.replace(b, nb, 1)
    return out


def finalize_line_label_layout(path):
    """Post-process a saved .xlsx: distribute line-chart last-point labels and
    add leader lines. Safe to call on any workbook; non-line charts are skipped."""
    wb = load_workbook(path)
    zin = zipfile.ZipFile(path)
    updates = {}
    for n in zin.namelist():
        if re.match(r"xl/charts/chart\d+\.xml$", n):
            xml = zin.read(n).decode("utf-8")
            if "<lineChart>" in xml:
                updates[n] = _process_line_xml(xml, wb)
    zin.close()
    if not updates:
        return
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in updates:
                data = updates[item.filename].encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)


def save_workbook(wb, path):
    """Save the workbook and apply line-label distribution + leader lines.
    Use this instead of wb.save() so charts come out fully styled."""
    wb.save(path)
    finalize_line_label_layout(path)
