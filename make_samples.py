#!/usr/bin/env python3
"""Create a sample model and a 5-slide pitchbook with planted errors.

Writes samples/project_atlas_model.xlsx and samples/project_atlas_pitchbook.pptx.
The company and peers are fictional. PLANTED lists every deliberate error.
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

OUT = Path(__file__).parent / "samples"
MODEL_PATH = OUT / "project_atlas_model.xlsx"
DECK_PATH = OUT / "project_atlas_pitchbook.pptx"

# (slide, Excel cell) of each planted slide vs model error.
PLANTED = {
    (1, "Income Statement!D5"): "FY2024 revenue KPI shows $4.5B, model is 4,213",
    (2, "Income Statement!D11"): "FY2024 EBITDA digits transposed, 985 vs 958",
    (2, "Income Statement!C9"): "FY2023 gross margin 41.2% vs 42.1%",
    (3, "Income Statement!C5"): "chart FY2023 revenue 3,900 vs 3,813",
    (4, "Valuation!B9"): "net debt units error, $18.5B vs 1,850",
    (4, "Valuation!B11"): "EV / LTM EBITDA 12.4x vs 13.3x",
    (4, "Cash Flow!D7"): "FY2024 free cash flow shown negative",
    (5, "Comps!C6"): "Fabrikam P/E 18.5x vs 16.2x",
    (5, "Income Statement!E5"): "FY2025E revenue shows the FY2024 figure, $4.2B",
}

# Slide pairs and item that must be reported as conflicting across slides.
PLANTED_CROSS = {
    ("revenue", "2024", 1, 2), ("revenue", "2024", 1, 3),
    ("revenue", "2023", 2, 3),
    ("revenue", "2025", 2, 5), ("revenue", "2025", 3, 5),
    ("ebitda", "2024", 2, 3),
    ("customer", None, 1, 5),
}

YEARS = ["FY2022A", "FY2023A", "FY2024A", "FY2025E", "FY2026E"]
REVENUE = [3120.4, 3812.6, 4213.0, 4592.2, 5005.5]
GROSS_MARGIN = [0.410, 0.421, 0.428, 0.432, 0.435]
EBITDA = [642.0, 820.5, 958.0, 1075.0, 1196.0]
DA = [-150.0, -165.0, -180.0, -195.0, -210.0]
INTEREST = [-88.0, -85.0, -80.0, -74.0, -68.0]
CAPEX = [-140.0, -165.0, -185.0, -200.0, -215.0]
NWC = [-20.0, -25.0, -30.0, -30.0, -30.0]
TAX_RATE = 0.24
SHARES = 210.0
OFFER, CURRENT = 52.0, 42.5
CASH, DEBT = 350.0, 2200.0

GROSS_PROFIT = [round(r * m, 1) for r, m in zip(REVENUE, GROSS_MARGIN)]
EBIT = [e + d for e, d in zip(EBITDA, DA)]
PRETAX = [e + i for e, i in zip(EBIT, INTEREST)]
TAXES = [round(-p * TAX_RATE, 1) for p in PRETAX]
NET_INCOME = [p + t for p, t in zip(PRETAX, TAXES)]
EPS = [round(n / SHARES, 2) for n in NET_INCOME]
FCF = [round(e + c + w + i + t, 1) for e, c, w, i, t in zip(EBITDA, CAPEX, NWC, INTEREST, TAXES)]
GROWTH = [None] + [round(b / a - 1, 4) for a, b in zip(REVENUE, REVENUE[1:])]
EBITDA_MARGIN = [round(e / r, 4) for e, r in zip(EBITDA, REVENUE)]

EQUITY = OFFER * SHARES
NET_DEBT = DEBT - CASH
EV = EQUITY + NET_DEBT

PEERS = [
    ("Northwind Corp", 14.2, 21.3, 0.241),
    ("Contoso Industries", 12.8, 19.7, 0.218),
    ("Fabrikam Inc", 11.5, 16.2, 0.197),
    ("Tailspin Holdings", 13.6, 20.1, 0.229),
]

MONEY_FMT, PCT_FMT, X_FMT, PS_FMT = "#,##0.0", "0.0%", '0.0"x"', "$0.00"
INK, GREEN, MINT, GREY, WHITE = (
    RGBColor(0x14, 0x2B, 0x26), RGBColor(0x1F, 0x5C, 0x4D), RGBColor(0xE3, 0xF0, 0xEB),
    RGBColor(0x5B, 0x6B, 0x66), RGBColor(0xFF, 0xFF, 0xFF),
)


def build_model() -> None:
    wb = Workbook()
    bold = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="E3F0EB")

    def sheet(ws, title, rows):
        ws["A1"], ws["A1"].font = "Project Atlas: " + title, Font(bold=True, size=14)
        ws["A2"] = "($ in millions, except per share data)"
        ws["A4"] = title
        for col, year in enumerate(YEARS, 2):
            cell = ws.cell(row=4, column=col, value=year)
            cell.font, cell.fill = bold, header_fill
        for r, (label, values, fmt) in enumerate(rows, 5):
            ws.cell(row=r, column=1, value=label)
            for c, v in enumerate(values, 2):
                if v is not None:
                    ws.cell(row=r, column=c, value=v).number_format = fmt
        ws.column_dimensions["A"].width = 28

    ws = wb.active
    ws.title = "Income Statement"
    sheet(ws, "Income Statement", [
        ("Revenue", REVENUE, MONEY_FMT),                                   # row 5
        ("Revenue Growth", GROWTH, PCT_FMT),
        ("Cost of Goods Sold", [round(g - r, 1) for g, r in zip(GROSS_PROFIT, REVENUE)], MONEY_FMT),
        ("Gross Profit", GROSS_PROFIT, MONEY_FMT),
        ("Gross Margin", GROSS_MARGIN, PCT_FMT),                           # row 9
        ("Operating Expenses", [round(e - g, 1) for e, g in zip(EBITDA, GROSS_PROFIT)], MONEY_FMT),
        ("EBITDA", EBITDA, MONEY_FMT),                                     # row 11
        ("EBITDA Margin", EBITDA_MARGIN, PCT_FMT),
        ("Depreciation and Amortization", DA, MONEY_FMT),
        ("EBIT", EBIT, MONEY_FMT),
        ("Interest Expense", INTEREST, MONEY_FMT),
        ("Pre-Tax Income", PRETAX, MONEY_FMT),
        ("Income Taxes", TAXES, MONEY_FMT),
        ("Net Income", NET_INCOME, MONEY_FMT),
        ("Diluted Shares (mm)", [SHARES] * 5, MONEY_FMT),
        ("EPS", EPS, PS_FMT),
    ])

    sheet(wb.create_sheet("Cash Flow"), "Cash Flow", [
        ("Capital Expenditures", CAPEX, MONEY_FMT),                        # row 5
        ("Change in Working Capital", NWC, MONEY_FMT),
        ("Free Cash Flow", FCF, MONEY_FMT),                                # row 7
    ])

    val = wb.create_sheet("Valuation")
    val["A1"], val["A1"].font = "Project Atlas: Transaction Valuation", Font(bold=True, size=14)
    val["A2"] = "($ in millions, except per share data)"
    for r, (label, value, fmt) in enumerate([
        ("Current Share Price", CURRENT, PS_FMT),                          # row 4
        ("Offer Price", OFFER, PS_FMT),
        ("Offer Premium", round(OFFER / CURRENT - 1, 4), PCT_FMT),
        ("Diluted Shares (mm)", SHARES, MONEY_FMT),
        ("Equity Value", EQUITY, MONEY_FMT),
        ("Net Debt", NET_DEBT, MONEY_FMT),                                 # row 9
        ("Enterprise Value", EV, MONEY_FMT),
        ("EV / LTM EBITDA", round(EV / EBITDA[2], 2), X_FMT),              # row 11
        ("EV / NTM EBITDA", round(EV / EBITDA[3], 2), X_FMT),
        ("Net Leverage", round(NET_DEBT / EBITDA[2], 2), X_FMT),
    ], 4):
        val.cell(row=r, column=1, value=label)
        val.cell(row=r, column=2, value=value).number_format = fmt
    val.column_dimensions["A"].width = 28

    comps = wb.create_sheet("Comps")
    comps["A1"], comps["A1"].font = "Project Atlas: Trading Comparables", Font(bold=True, size=14)
    for c, head in enumerate(["Company", "EV / LTM EBITDA", "P/E", "EBITDA Margin"], 1):
        cell = comps.cell(row=3, column=c, value=head)
        cell.font, cell.fill = bold, header_fill
    for r, (name, ev, pe, margin) in enumerate(PEERS, 4):
        comps.cell(row=r, column=1, value=name)
        comps.cell(row=r, column=2, value=ev).number_format = X_FMT
        comps.cell(row=r, column=3, value=pe).number_format = X_FMT
        comps.cell(row=r, column=4, value=margin).number_format = PCT_FMT
    comps.column_dimensions["A"].width = 24
    comps.column_dimensions["B"].width = 18

    wb.save(MODEL_PATH)


def _text(slide, x, y, w, h, paragraphs, size=16, color=INK, bold=False, name=None, align=None, anchor=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        box.name = name
    frame = box.text_frame
    frame.word_wrap = True
    if anchor:
        frame.vertical_anchor = anchor
    for i, line in enumerate([paragraphs] if isinstance(paragraphs, str) else paragraphs):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        para.space_after = Pt(10)
        if align:
            para.alignment = align
        run = para.add_run()
        run.text = line
        run.font.size, run.font.color.rgb, run.font.bold, run.font.name = Pt(size), color, bold, "Calibri"
    return box


def _slide(prs, number, title):
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    t = slide.shapes.title
    t.left, t.top, t.width, t.height = Inches(0.6), Inches(0.35), Inches(12.1), Inches(0.9)
    t.text_frame.text = title
    t.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = t.text_frame.paragraphs[0]
    para.alignment = PP_ALIGN.LEFT
    run = para.runs[0]
    run.font.size, run.font.bold, run.font.color.rgb, run.font.name = Pt(32), True, GREEN, "Cambria"
    _text(slide, 12.2, 7.0, 0.6, 0.35, str(number), size=11, color=GREY, name="Page number", align=PP_ALIGN.RIGHT)
    return slide


def _kpi(slide, x, y, value, label, index):
    card = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(2.85), Inches(1.75))
    card.name = f"KPI card {index}"
    card.fill.solid()
    card.fill.fore_color.rgb = MINT
    card.line.fill.background()
    _text(slide, x + 0.1, y + 0.2, 2.65, 0.9, value, size=36, color=GREEN, bold=True,
          name=f"KPI value {index}", align=PP_ALIGN.CENTER)
    _text(slide, x + 0.1, y + 1.12, 2.65, 0.45, label, size=14, color=INK,
          name=f"KPI label {index}", align=PP_ALIGN.CENTER)


def _table(slide, x, y, w, row_h, rows, first_col, name, size=14):
    shape = slide.shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y), Inches(w), Inches(row_h * len(rows)))
    shape.name = name
    table = shape.table
    table.columns[0].width = Inches(first_col)
    rest = (w - first_col) / (len(rows[0]) - 1)
    for c in range(1, len(rows[0])):
        table.columns[c].width = Inches(rest)
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            cell = table.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = GREEN if r == 0 else (MINT if r % 2 == 0 else WHITE)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            para = cell.text_frame.paragraphs[0]
            para.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT
            run = para.add_run()
            run.text = value
            run.font.size, run.font.name = Pt(size), "Calibri"
            run.font.bold = r == 0
            run.font.color.rgb = WHITE if r == 0 else INK


def build_deck(path: Path = DECK_PATH, planted: bool = True) -> None:
    """Build the deck. With planted=False every figure ties, for false-positive testing."""

    def pick(wrong, right):
        return wrong if planted else right

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    # Slide 1: executive summary. PLANTED: $4.5B revenue, 1,240 customers.
    s = _slide(prs, 1, "Project Atlas: Executive Summary")
    for i, (value, label) in enumerate([
        (pick("$4.5B", "$4.2B"), "FY2024 Revenue"),
        ("22.7%", "FY2024 EBITDA Margin"),
        ("$12.8B", "Enterprise Value"),
        (pick("1,240", "1,420"), "Customers"),
    ]):
        _kpi(s, 0.6 + i * 3.08, 1.5, value, label, i + 1)
    _text(s, 0.6, 3.75, 12.1, 3.0, [
        "Offer price of $52.00 per share, a 22.4% premium to the current $42.50 share price",
        "Net income of $530.5M in FY2024, with EPS of $2.53",
        "Three strategic priorities drive the value creation plan over the next 5 years",
    ], size=20, name="Summary bullets")

    # Slide 2: financial summary. PLANTED: EBITDA 985, FY2023 gross margin 41.2%.
    s = _slide(prs, 2, "Financial Summary")
    fmt = lambda values: [f"{v:,.0f}" for v in values]
    pct = lambda values: ["n/a" if v is None else f"{v:.1%}" for v in values]
    ebitda = fmt(EBITDA)
    ebitda[2] = pick("985", ebitda[2])
    gross_margin = pct(GROSS_MARGIN)
    gross_margin[1] = pick("41.2%", gross_margin[1])
    _table(s, 0.6, 1.5, 12.1, 0.47, [
        ["($ in millions, except per share)"] + YEARS,
        ["Revenue"] + fmt(REVENUE),
        ["Revenue Growth"] + pct(GROWTH),
        ["Gross Profit"] + fmt(GROSS_PROFIT),
        ["Gross Margin"] + gross_margin,
        ["EBITDA"] + ebitda,
        ["EBITDA Margin"] + pct(EBITDA_MARGIN),
        ["Net Income"] + fmt(NET_INCOME),
        ["EPS"] + [f"${v:.2f}" for v in EPS],
        ["Capex"] + fmt([-c for c in CAPEX]),
    ], 4.1, "Financial summary table")
    _text(s, 0.6, 6.45, 11, 0.4, "Source: Project Atlas financial model, FY2022A to FY2026E.", size=12, color=GREY, name="Source")

    # Slide 3: revenue chart. PLANTED: FY2023 chart bar 3,900.
    s = _slide(prs, 3, "Revenue Growth Trajectory")
    data = CategoryChartData(number_format="#,##0")
    data.categories = YEARS
    bars = [round(v) for v in REVENUE]
    bars[1] = pick(3900, bars[1])
    data.add_series("Revenue ($ in millions)", bars)
    chart = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.6), Inches(1.5), Inches(7.4), Inches(5.2), data)
    chart.name = "Revenue chart"
    c = chart.chart
    c.has_legend = False
    c.has_title = True
    c.chart_title.text_frame.text = "Revenue ($ in millions)"
    c.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(16)
    c.value_axis.visible = False
    c.value_axis.has_major_gridlines = False
    c.category_axis.tick_labels.font.size = Pt(14)
    plot = c.plots[0]
    plot.gap_width = 60
    plot.has_data_labels = True
    plot.data_labels.number_format = "#,##0"
    plot.data_labels.number_format_is_linked = False
    plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    plot.data_labels.font.size = Pt(14)
    plot.series[0].format.fill.solid()
    plot.series[0].format.fill.fore_color.rgb = GREEN
    _text(s, 8.4, 1.7, 4.3, 4.8, [
        "FY2024 revenue of $4.2 billion, up 10.5%",
        "FY2024 EBITDA of $958M, a 22.7% margin",
        "Revenue of $4,213mm in FY2024 ties to the audited accounts",
    ], size=20, name="Commentary")

    # Slide 4: valuation. PLANTED: net debt $18.5B, 12.4x, negative FCF.
    s = _slide(prs, 4, "Transaction Valuation")
    _table(s, 0.6, 1.5, 7.6, 0.45, [
        ["($ in millions, except per share)", "Value"],
        ["Offer Price", f"${OFFER:.2f}"],
        ["Offer Premium", "22.4%"],
        ["Diluted Shares (mm)", "210.0"],
        ["Equity Value", f"${EQUITY:,.0f}"],
        ["Net Debt", pick("$18.5B", "$1.85B")],
        ["Enterprise Value", f"${EV:,.0f}"],
        ["EV / LTM EBITDA", pick("12.4x", "13.3x")],
        ["EV / NTM EBITDA", "11.9x"],
        ["Net Leverage", "1.9x"],
        ["FY2024 Free Cash Flow", pick(f"({FCF[2]:,.1f})", f"{FCF[2]:,.1f}")],
    ], 5.4, "Valuation table")
    _kpi(s, 9.2, 1.5, "13.3x", "EV / LTM EBITDA (model)", 5)
    _kpi(s, 9.2, 3.6, "$10.9B", "Equity Value", 6)

    # Slide 5: comps and outlook. PLANTED: P/E 18.5x, $4.2B FY2025E, 1,420 customers, 112% not in model.
    s = _slide(prs, 5, "Peer Benchmarking and Outlook")
    peer_rows = [["Company", "EV / LTM EBITDA", "P/E", "EBITDA Margin"]]
    for name, ev, pe, margin in PEERS:
        shown_pe = pick("18.5x", f"{pe:.1f}x") if name == "Fabrikam Inc" else f"{pe:.1f}x"
        peer_rows.append([name, f"{ev:.1f}x", shown_pe, f"{margin:.1%}"])
    _table(s, 0.6, 1.5, 12.1, 0.47, peer_rows, 4.3, "Peer table")
    _text(s, 0.6, 4.2, 12.1, 2.6, [
        f"FY2025E revenue of {pick('$4.2B', '$4.6B')} supports the growth case",
        "A base of 1,420 customers across 14 countries",
        "Net revenue retention of 112% in the latest cohort",
    ], size=20, name="Outlook")

    prs.save(path)


def main() -> None:
    OUT.mkdir(exist_ok=True)
    build_model()
    build_deck()
    print(f"Wrote {MODEL_PATH}")
    print(f"Wrote {DECK_PATH}")
    print(f"Planted slide vs model errors: {len(PLANTED)}")


if __name__ == "__main__":
    main()
