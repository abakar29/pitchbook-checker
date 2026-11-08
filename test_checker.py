"""Run with: python -m pytest -q"""
import pytest
from openpyxl import Workbook

import make_samples
from pitchbook_checker import (COUNT, Fact, Finding, MONEY, MULTIPLE, Occurrence, PER_SHARE, PERCENT, check, finding_card, format_like,
                               full_period, parse_numbers, read_excel, summary_html, title_case, values_match)


@pytest.fixture(scope="module")
def planted():
    make_samples.OUT.mkdir(exist_ok=True)
    make_samples.build_model()
    make_samples.build_deck()
    return check(make_samples.DECK_PATH, make_samples.MODEL_PATH)


def one(text):
    numbers = parse_numbers(text)
    assert len(numbers) == 1, numbers
    return numbers[0]


def slide_fact(text, note_scale=None):
    p = one(text)
    return Fact("s", "x", "x", None, p.kind, p.value, p.scale, p.tol, p.raw, slide=1, note_scale=note_scale)


def excel_fact(value, scale=1e6, kind=MONEY):
    return Fact("c", "x", "x", None, kind, value, scale, abs(value) * 1e-9, str(value))


def test_every_planted_model_error_is_found_and_nothing_else(planted):
    found = {(i.slide, i.excel_cell) for i in planted.model_issues}
    assert found == set(make_samples.PLANTED)


def test_every_planted_cross_slide_conflict_is_found_and_nothing_else(planted):
    metrics = {}
    for f in planted.slides:
        metrics[(f.slide, f.raw, f.location)] = (f.metric, f.period)
    found = set()
    for i in planted.cross_issues:
        metric, period = metrics[(i.slide, i.slide_value, i.location)]
        found.add((metric, period, i.slide, i.other_slide))
    assert found == make_samples.PLANTED_CROSS


def test_number_missing_from_model_is_listed_for_review(planted):
    assert any(f.raw == "112%" and f.slide == 5 for f in planted.review)


def test_clean_deck_has_no_issues(tmp_path):
    make_samples.OUT.mkdir(exist_ok=True)
    make_samples.build_model()
    deck = tmp_path / "clean.pptx"
    make_samples.build_deck(deck, planted=False)
    result = check(deck, make_samples.MODEL_PATH)
    assert result.model_issues == []
    assert result.cross_issues == []


def test_page_numbers_and_years_are_not_flagged(planted):
    assert not any(f.location == 'text "Page number"' and f.strong for f in planted.slides)
    assert not any(f.raw.startswith("FY") for f in planted.slides)


@pytest.mark.parametrize("text,kind,value,scale", [
    ("$4.2B", MONEY, 4.2, 1e9),
    ("$4.2 billion", MONEY, 4.2, 1e9),
    ("$4,213mm", MONEY, 4213, 1e6),
    ("$958M", MONEY, 958, 1e6),
    ("€1.5bn", MONEY, 1.5, 1e9),
    ("$500K", MONEY, 500, 1e3),
    ("22.7%", PERCENT, 22.7, None),
    ("250 bps", PERCENT, 2.5, None),
    ("12.4x", MULTIPLE, 12.4, None),
    ("($45.2)", MONEY, -45.2, None),
    ("(45.2)", "number", -45.2, None),
    ("-5.0%", PERCENT, -5.0, None),
])
def test_parse(text, kind, value, scale):
    p = one(text)
    assert (p.kind, p.value, p.scale) == (kind, value, scale)


@pytest.mark.parametrize("text", ["FY2024", "2025E", "Q3 2024", "10-K", "S&P 500", "12/31/2024", "COVID-19", "10th"])
def test_not_numbers(text):
    assert parse_numbers(text) == []


def test_parenthetical_in_a_sentence_is_not_negative():
    assert parse_numbers("Net leverage (2.5x) remains low")[0].value == 2.5


def test_range_shares_unit():
    low, high = parse_numbers("$3.6-4.2B")
    assert (low.scale, high.scale) == (1e9, 1e9)


@pytest.mark.parametrize("label,expected", [("FY24", "2024"), ("FY2024A", "2024"), ("2025E", "2025"),
                                            ("Q3'24", "2024Q3"), ("TTM", "LTM")])
def test_periods(label, expected):
    assert full_period(label) == expected


@pytest.mark.parametrize("shown,model,ok", [
    ("$4.2B", 4213.0, True),
    ("$4.2B", 4150.0, True),
    ("$4.2B", 4300.0, False),
    ("$4.21B", 4213.0, True),
    ("$4.22B", 4213.0, False),
    ("$4,213mm", 4213.4, True),
    ("$4,213mm", 4214.0, False),
])
def test_rounding_and_units(shown, model, ok):
    assert values_match(slide_fact(shown), excel_fact(model)) is ok


def test_bare_table_number_uses_slide_units_note():
    assert values_match(slide_fact("4,213", note_scale=1e6), excel_fact(4213.0))
    assert values_match(slide_fact("4,213", note_scale=1e3), excel_fact(4213.0)) is False


def test_uncached_formulas_warn(tmp_path):
    wb = Workbook()
    wb.active["A1"] = "Revenue"
    wb.active["B1"] = "=1+1"
    path = tmp_path / "formulas.xlsx"
    wb.save(path)
    facts, warnings = read_excel(path)
    assert facts == []
    assert any("no saved value" in w for w in warnings)


def test_duplicates_merge_into_one_card_per_item(planted):
    assert len(planted.findings) == 11
    assert planted.html.count("<article") == 11
    revenue = next(f for f in planted.findings if f.item == "Revenue" and f.period == "FY2024")
    assert [o.slide for o in revenue.occurrences] == [1, 2, 3, 3, 3]
    assert [o.status for o in revenue.occurrences][:2] == ["wrong", "ok"]


def test_summary_counts_and_jump_links(planted):
    assert "11 issues found on 5 slides" in planted.html
    for slide in range(1, 6):
        assert f'href="#slide-{slide}"' in planted.html
        assert planted.html.count(f'id="slide-{slide}"') == 1


def test_card_quotes_the_slide_and_highlights_the_number(planted):
    assert 'FY2025E revenue of <mark class="pc-hit">$4.2B</mark> supports the growth case' in planted.html
    assert 'EBITDA, FY2024A: <mark class="pc-hit">985</mark>' in planted.html
    assert "KPI value" not in planted.html and "Financial summary table" not in planted.html


def test_suggested_fixes(planted):
    assert "Change $4.5B to $4.2B on slide 1." in planted.html
    assert "Change 985 to 958 on slide 2." in planted.html
    assert "Change (495.5) to 495.5 on slide 4." in planted.html
    assert "Confirm the correct figure, then use it on slides 1 and 5." in planted.html
    assert all(f.fix for f in planted.findings)


def test_conflict_not_in_model_is_one_card(planted):
    customers = [f for f in planted.findings if f.item == "Customers"]
    assert len(customers) == 1 and customers[0].style == "conflict"


@pytest.mark.parametrize("shown,value,scale,kind,expected", [
    ("$241mm", 214.0, 1e6, MONEY, "$214mm"),
    ("$1.1 billion", 1190.0, 1e6, MONEY, "$1.2 billion"),
    ("$1,100mm", 1190.0, 1e6, MONEY, "$1,190mm"),
    ("2,100", 2400.0, 1, COUNT, "2,400"),
    ("41.2%", 42.1, 1, PERCENT, "42.1%"),
    ("12.4x", 13.33, 1, MULTIPLE, "13.3x"),
    ("(495.5)", 495.5, 1e6, MONEY, "495.5"),
    ("$40.00", 42.5, 1, PER_SHARE, "$42.50"),
    ("250 bps", 3.1, 1, PERCENT, "310 bps"),
])
def test_model_value_written_like_the_slide(shown, value, scale, kind, expected):
    fact = slide_fact(shown)
    fact.metric = "x"
    assert format_like(fact, excel_fact(value, scale, kind)) == expected


def test_money_note_does_not_scale_headcount():
    fact = slide_fact("2,400", note_scale=1e6)
    assert values_match(fact, excel_fact(2400.0, 1, COUNT))


@pytest.mark.parametrize("raw,expected", [("FY2024 ebitda margin", "EBITDA Margin"), ("employees", "Employees"),
                                          ("revenue ($ in millions)", "Revenue"), ("cost of sales", "Cost of Sales")])
def test_title_case(raw, expected):
    assert title_case(raw) == expected


def test_placeholder_sentence_end_to_end(tmp_path):
    from pptx import Presentation

    wb = Workbook()
    ws = wb.active
    for row in [["($ in millions)"], ["", 2023, 2024], ["EBITDA", 190.0, 214.0], ["Employees", 2200, 2400]]:
        ws.append(row)
    model = tmp_path / "m.xlsx"
    wb.save(model)
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Overview ($ in millions)"
    slide.placeholders[1].text_frame.text = "2024 EBITDA of $241mm"
    slide.placeholders[1].text_frame.add_paragraph().text = "Headcount of 2,100 employees in 2024"
    deck = tmp_path / "d.pptx"
    prs.save(deck)

    result = check(deck, model)
    by_item = {f.item: f for f in result.findings}
    assert set(by_item) == {"EBITDA", "Employees"}
    assert by_item["EBITDA"].fix == "Change $241mm to $214mm on slide 1."
    assert by_item["Employees"].model_value == "2,400"
    assert by_item["Employees"].fix == "Change 2,100 to 2,400 on slide 1."
    assert '2024 EBITDA of <mark class="pc-hit">$241mm</mark>' in result.html


def test_all_clear_only_when_every_number_is_verified():
    assert "All clear" in summary_html([], 78)
    open_item = Finding("missing", "X", "no period", [Occurrence(2, "1%", ("", "1%", ""), "unverified")],
                        "1%", "", "n/a", "Check it.")
    page = summary_html([open_item], 78)
    assert "All clear" not in page and "1 issue found on 1 slide" in page
    assert "All clear" not in summary_html([], 0)


def test_html_escapes_deck_text():
    occurrence = Occurrence(1, "$1", ("<script>", "$1", "</script>"), "wrong")
    page = finding_card(Finding("mismatch", "<b>x</b>", "FY2024", [occurrence], "$1", "S!A1", "$2", "<i>fix</i>"))
    assert "<script>" not in page and "<i>fix</i>" not in page and "<b>x</b>" not in page
    assert "&lt;script&gt;" in page
