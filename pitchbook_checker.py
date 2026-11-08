#!/usr/bin/env python3
"""Check a PowerPoint pitchbook against an Excel financial model.

Usage:
    python pitchbook_checker.py deck.pptx model.xlsx [--out report] [--model-units millions] [--no-open]

Every number is extracted with its label and location, then two checks run:
  1. each slide number is compared with the matching Excel cell
  2. the same item shown on different slides is compared across slides

Units and rounding: a displayed number stands for a range. "$4.2B" means
4.15B to 4.25B, so 4,213 in millions passes and 4,300 does not.

Exit code: 0 if clean, 1 if any mismatch was found, 2 on bad input.
"""
from __future__ import annotations

import argparse
import csv
import re
import sys
import webbrowser
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal
from html import escape
from itertools import combinations
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import openpyxl
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER

# ---------------------------------------------------------------------------
# Number and period parsing
# ---------------------------------------------------------------------------

PERIOD_RE = re.compile(
    r"""
    (?<![\w$€£.,])
    (?:
        (?P<q>[QH][1-4])(?:\s*(?:FY|CY)?\s*(?P<qy>(?:19|20)\d{2}|['’]\d{2}))?
      | (?P<q2>[1-4][QH])\s*['’]?(?P<q2y>(?:19|20)\d{2}|\d{2})
      | (?:FY|CY|Fiscal\s+(?:Year\s+)?)\s*['’]?(?P<fy>(?:19|20)\d{2}|\d{2})[AEPFB]?
      | (?P<bare>(?:19|20)\d{2})(?P<bare_sfx>[AEPFB])?
        (?!,\d|\.\d|\s?%|\s?(?:x|mm|bn|million|billion|thousand)\b)
      | (?P<rel>LTM|TTM|NTM)
    )
    (?!\w)
    """,
    re.X | re.I,
)

# Digits that are part of a name or date, not a financial figure.
NOISE_RE = re.compile(
    r"\b\d+-[KQ]\b|S&P\s*\d+|COVID-19|401\(k\)"
    r"|\b\d{1,2}/\d{1,2}(?:/\d{2,4})?\b"
    r"|\b(?:slide|page|note|appendix|top)\s+\d+\b"
    r"|\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s+\d{1,2}(?:,\s*\d{4})?\b",
    re.I,
)

NUMBER_RE = re.compile(
    r"""
    (?<![\w.,])
    (?P<minus>[-−](?=\(?[$€£\d.]))?
    (?P<lp>\(\s*)?
    (?P<cur>US\$|USD\s?|[$€£])?
    (?P<lp2>\(\s*)?
    (?P<minus2>[-−])?
    (?P<num>\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?|\.\d+)
    (?:(?P<sp>\s?)(?P<sfx>%|percent\b|bps\b|basis\s+points\b
        |trillion\b|tn\b|billion\b|bn\b|million\b|mm\b|mn\b|thousand\b
        |x(?![a-z0-9])|[tbmk](?![a-z0-9])))?
    (?P<rp>\s*\))?
    (?!\w)
    """,
    re.X | re.I,
)

SCALES = {
    "t": 1e12, "tn": 1e12, "trillion": 1e12,
    "b": 1e9, "bn": 1e9, "billion": 1e9, "billions": 1e9,
    "m": 1e6, "mm": 1e6, "mn": 1e6, "million": 1e6, "millions": 1e6,
    "k": 1e3, "thousand": 1e3, "thousands": 1e3, "000s": 1e3, "000's": 1e3,
}

SCALE_NOTE_RE = re.compile(
    r"(?:\bin\s+(?:\$\s*|us\$\s*|usd\s*)?|\$\s*)(thousands|millions|billions|mm|mn|bn|000'?s)\b(?!\s+of\b)",
    re.I,
)
LABEL_SCALE_RE = re.compile(
    r"\(\s*(?:\$|us\$|usd)?\s*(?:in\s+)?(?:\$\s*)?(mm|mn|millions?|bn|billions?|thousands?|000'?s)\s*\)",
    re.I,
)

MONEY, PERCENT, MULTIPLE, PER_SHARE, COUNT, PLAIN = (
    "money", "percent", "multiple", "per_share", "count", "number",
)
COMPATIBLE = {
    PERCENT: {PERCENT},
    MULTIPLE: {MULTIPLE},
    MONEY: {MONEY, PER_SHARE, COUNT},
    PLAIN: {MONEY, PER_SHARE, COUNT, PERCENT, MULTIPLE},
}


@dataclass
class Parsed:
    value: float
    decimals: int
    kind: str
    scale: Optional[float]
    raw: str
    start: int = 0
    end: int = 0
    tol: float = 0.0


def _year(text: str) -> str:
    text = text.strip("'’ ")
    return text if len(text) == 4 else "20" + text


def period_of(match: re.Match) -> str:
    g = match.groupdict()
    if g["rel"]:
        return "NTM" if g["rel"].upper() == "NTM" else "LTM"
    if g["q"]:
        return f"{_year(g['qy'])}{g['q'].upper()}" if g["qy"] else g["q"].upper()
    if g["q2"]:
        return f"{_year(g['q2y'])}{g['q2'][1].upper()}{g['q2'][0]}"
    return _year(g["fy"] or g["bare"])


def full_period(text: str) -> Optional[str]:
    """The period if the whole string is one, e.g. 'FY2024A'."""
    m = PERIOD_RE.fullmatch(text.strip())
    return period_of(m) if m else None


def period_in(text: str) -> Optional[str]:
    found = {period_of(m) for m in PERIOD_RE.finditer(text)}
    return found.pop() if len(found) == 1 else None


def show_period(period: Optional[str]) -> str:
    if not period:
        return "no period"
    return f"FY{period}" if period.isdigit() else period


def mask(text: str, keep_bare_years: bool = False) -> str:
    """Blank out periods and names so digits in them are not read as numbers."""

    def blank(m: re.Match) -> str:
        return " " * len(m.group(0))

    def blank_period(m: re.Match) -> str:
        if keep_bare_years and m.group("bare") and not m.group("bare_sfx"):
            return m.group(0)
        return blank(m)

    return PERIOD_RE.sub(blank_period, NOISE_RE.sub(blank, text))


def parse_numbers(text: str, keep_bare_years: bool = False) -> List[Parsed]:
    masked = mask(text, keep_bare_years)
    out: List[Parsed] = []
    for m in NUMBER_RE.finditer(masked):
        digits = m.group("num").replace(",", "")
        value = float(digits)
        decimals = len(digits.split(".")[1]) if "." in digits else 0
        sfx = (m.group("sfx") or "").lower()
        cur = m.group("cur")
        end = m.end()
        # "3 m" with no currency sign is more likely metres than millions.
        if sfx and (sfx == "x" or len(sfx) == 1 and sfx != "%") and m.group("sp") and not cur:
            sfx = ""
            end = m.end("num")

        kind, scale = (MONEY if cur else PLAIN), None
        if sfx in ("%", "percent"):
            kind = PERCENT
        elif sfx in ("bps", "basis points") or sfx.startswith("basis"):
            kind, value, decimals = PERCENT, value / 100, decimals + 2
        elif sfx == "x":
            kind = MULTIPLE
        elif sfx:
            kind, scale = MONEY, SCALES[sfx]

        whole = masked.strip() == m.group(0).strip()
        parens = (m.group("lp") or m.group("lp2")) and m.group("rp")
        # Accounting negatives: "($45.2)" anywhere, "(45.2)" only as a whole cell.
        accounting = parens and (whole or cur is not None)
        if m.group("minus") or m.group("minus2") or accounting:
            value = -value

        out.append(Parsed(value, decimals, kind, scale, text[m.start():end].strip(),
                          m.start(), end, 0.5 * 10 ** -decimals))

    # "$3.6-4.2B" and "10-12x": the first number takes the second one's unit.
    for a, b in zip(out, out[1:]):
        needs_unit = a.kind == PLAIN and b.kind != PLAIN or a.kind == MONEY and a.scale is None and b.scale
        if needs_unit and re.fullmatch("\\s?[-\u2013\u2014]\\s?|\\s+to\\s+", masked[a.end:b.start]):
            a.kind, a.scale = b.kind, b.scale
    return out


# ---------------------------------------------------------------------------
# Labels
# ---------------------------------------------------------------------------

ALIASES = {
    "revenue": ["revenue", "revenues", "total revenue", "net revenue", "net revenues", "sales",
                "net sales", "total sales", "turnover"],
    "revenue growth": ["revenue growth", "sales growth", "growth in revenue"],
    "cogs": ["cogs", "cost of goods sold", "cost of sales", "cost of revenue"],
    "gross profit": ["gross profit"],
    "gross margin": ["gross margin", "gross profit margin"],
    "operating expenses": ["operating expenses", "opex"],
    "ebitda": ["ebitda"],
    "adjusted ebitda": ["adjusted ebitda", "adj ebitda"],
    "ebitda margin": ["ebitda margin"],
    "adjusted ebitda margin": ["adjusted ebitda margin", "adj ebitda margin"],
    "d&a": ["d a", "depreciation and amortization", "depreciation amortization"],
    "ebit": ["ebit", "operating income", "operating profit"],
    "ebit margin": ["ebit margin", "operating margin"],
    "interest expense": ["interest expense", "net interest expense"],
    "net income": ["net income", "net profit", "net earnings"],
    "net margin": ["net margin", "net income margin"],
    "eps": ["eps", "earnings per share", "diluted eps"],
    "diluted shares": ["diluted shares", "diluted shares outstanding", "shares outstanding"],
    "capex": ["capex", "capital expenditures", "capital expenditure"],
    "fcf": ["fcf", "free cash flow", "levered free cash flow"],
    "operating cash flow": ["operating cash flow", "cash flow from operations"],
    "cash": ["cash", "cash and equivalents", "cash and cash equivalents"],
    "total debt": ["total debt", "gross debt"],
    "net debt": ["net debt"],
    "enterprise value": ["enterprise value", "ev", "tev", "total enterprise value",
                         "implied enterprise value"],
    "equity value": ["equity value", "implied equity value", "market cap", "market capitalization",
                     "equity purchase price"],
    "share price": ["share price", "stock price", "current share price"],
    "offer price": ["offer price", "offer price per share"],
    "premium": ["premium", "offer premium", "implied premium"],
    "ev/ebitda": ["ev ebitda", "tev ebitda", "enterprise value ebitda"],
    "ev/revenue": ["ev revenue", "tev revenue", "ev sales", "enterprise value revenue"],
    "p/e": ["p e", "pe ratio", "p e ratio", "price earnings"],
    "net leverage": ["net leverage", "net debt ebitda", "net debt to ebitda", "leverage"],
    "irr": ["irr"],
    "moic": ["moic"],
    "wacc": ["wacc"],
    "employees": ["employees", "headcount", "ftes", "full time employees"],
    "net revenue retention": ["net revenue retention", "net retention", "revenue retention", "nrr"],
    "revenue cagr": ["revenue cagr", "sales cagr"],
    "ebitda cagr": ["ebitda cagr"],
}
ALIAS_LOOKUP = {term: key for key, terms in ALIASES.items() for term in terms}
# Excel labels too generic to search for inside slide sentences.
VOCAB_STOPLIST = {"other", "total", "value", "amount", "net", "date", "name", "notes", "company", "median", "mean"}

PERCENT_LABEL_RE = re.compile(r"%|margin|growth|yield|premium|\birr\b|tax rate|discount rate|\bwacc\b", re.I)
MULTIPLE_LABEL_RE = re.compile(r"multiple|/\s*(?:ltm\s+|ntm\s+)?(?:ebitda|ebit|revenue|sales)\b|\bp\s*/\s*e\b|\bmoic\b|leverage|coverage", re.I)
PER_SHARE_LABEL_RE = re.compile(r"per share|\beps\b|(?:share|offer|stock) price", re.I)
COUNT_LABEL_RE = re.compile(r"shares|employees|customers|stores|headcount", re.I)
EXPENSE_RE = re.compile(r"capex|expense|\bcost|cogs|d&a|\bd a\b|depreciation|amortization|interest|tax", re.I)
DECLINE_RE = re.compile(r"\b(?:declin|decreas|down\b|fell|falling|loss|lower|reduc|negative|contract)", re.I)
GROWTH_RE = re.compile(r"\b(?:grow|grew|increas|up\b|rose|rising|expand|declin|decreas|down\b|fell)|\byoy\b", re.I)


def norm(text: str) -> str:
    text = PERIOD_RE.sub(" ", text)
    text = re.sub(r"\([^)]*\)", " ", text)
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def canon(label: str) -> str:
    """Reduce a label to a comparable key: 'Net Sales ($mm)' becomes 'revenue'."""
    n = norm(label)
    if n in ALIAS_LOOKUP:
        return ALIAS_LOOKUP[n]
    n = re.sub(r"^(?:total|implied|current|reported)\s+", "", n)
    if n in ALIAS_LOOKUP:
        return ALIAS_LOOKUP[n]
    if len(n) > 3 and n.endswith("s") and not n.endswith("ss"):
        n = n[:-1]
    return n


ACRONYMS = {w.lower(): w for w in ["EBITDA", "EBIT", "EPS", "EV", "TEV", "IRR", "MOIC", "WACC", "COGS", "NRR", "ARR",
                                   "CAGR", "ROE", "ROIC", "FCF", "LTM", "NTM", "YoY", "D&A", "P/E"]}
SMALL_WORDS = {"of", "and", "in", "to", "the", "a", "per", "for", "vs"}
DISPLAY_NAMES = {"ev/ebitda": "EV / EBITDA", "ev/revenue": "EV / Revenue", "fcf": "Free Cash Flow", "capex": "Capex"}


def _title_word(word: str, first: bool) -> str:
    if word.lower() in ACRONYMS:
        return ACRONYMS[word.lower()]
    if "/" in word and word != "/":
        return "/".join(_title_word(part, first) for part in word.split("/"))
    if word.lower() in SMALL_WORDS and not first:
        return word.lower()
    return word[:1].upper() + word[1:] if word.islower() else word


def title_case(label: str) -> str:
    text = re.sub(r"\([^)]*\)", " ", PERIOD_RE.sub(" ", label))
    return " ".join(_title_word(w, i == 0) for i, w in enumerate(text.split()))


def cross_key(a: str, b: str) -> str:
    return f"{canon(a)}|{canon(b)}"


# ---------------------------------------------------------------------------
# Facts
# ---------------------------------------------------------------------------


@dataclass
class Fact:
    location: str
    label: str
    metric: str
    period: Optional[str]
    kind: str
    value: float
    scale: Optional[float]
    tol: float
    raw: str
    slide: Optional[int] = None
    note_scale: Optional[float] = None
    strong: bool = False  # came from a table, chart or KPI card, so the label is reliable
    sign_free: bool = False
    context: str = ""
    best: Optional["Fact"] = None  # closest Excel cell, set during comparison
    quote: Tuple[str, str, str] = ("", "", "")  # text before, the number, text after
    status: Optional[str] = None  # ok, wrong or review, set during comparison
    expected: str = ""  # model value written the way the slide writes numbers
    note: str = ""


@dataclass
class Issue:
    type: str
    item: str
    period: str
    slide: int
    location: str
    slide_value: str
    excel_cell: str = ""
    excel_value: str = ""
    other_slide: Optional[int] = None
    other_location: str = ""
    other_value: str = ""
    note: str = ""


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------


def _note_scale(text: str) -> Optional[float]:
    m = SCALE_NOTE_RE.search(text)
    return SCALES[m.group(1).lower()] if m else None


def _label_scale(text: str) -> Optional[float]:
    m = LABEL_SCALE_RE.search(text)
    return SCALES.get(m.group(1).lower()) if m else None


def _is_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _classify(label: str, fmt: str, value: float) -> Tuple[str, float]:
    """Return the kind and the value in display terms (0.227 becomes 22.7 for percents)."""
    fmt = (fmt or "").lower()
    if "%" in fmt:
        return PERCENT, value * 100
    if re.search(r'x"?\s*(?:;|$)', fmt):
        return MULTIPLE, value
    if MULTIPLE_LABEL_RE.search(label):
        return MULTIPLE, value
    if PERCENT_LABEL_RE.search(label):
        return PERCENT, value * 100 if abs(value) <= 1.5 else value
    if PER_SHARE_LABEL_RE.search(label):
        return PER_SHARE, value
    if COUNT_LABEL_RE.search(label):
        return COUNT, value
    return MONEY, value


def _header_rows(cells: Dict[Tuple[int, int], object]) -> Dict[int, Dict[int, str]]:
    """Find rows of period headers, as text, dates, or consecutive integer years."""
    by_row: Dict[int, Dict[int, str]] = defaultdict(dict)
    int_years: Dict[int, Dict[int, str]] = defaultdict(dict)
    for (r, c), v in cells.items():
        if isinstance(v, str) and full_period(v):
            by_row[r][c] = full_period(v)
        elif isinstance(v, (datetime, date)):
            int_years[r][c] = str(v.year)
        elif _is_number(v) and float(v).is_integer() and 1900 <= v <= 2100:
            int_years[r][c] = str(int(v))
    for r, cols in int_years.items():
        years = [int(cols[c]) for c in sorted(cols)]
        if len(years) >= 2 and all(b - a == 1 for a, b in zip(years, years[1:])):
            by_row[r].update(cols)
    return {r: cols for r, cols in by_row.items()}


def _show_excel(kind: str, value: float, scale: float) -> str:
    if kind == PERCENT:
        return f"{value:,.2f}%"
    if kind == MULTIPLE:
        return f"{value:,.2f}x"
    if kind == PER_SHARE:
        return f"${value:,.2f}"
    if kind == COUNT and scale == 1:
        return f"{value:,.0f}" if float(value).is_integer() else f"{value:,.1f}"
    unit = {1e3: "k", 1e6: "mm", 1e9: "bn"}.get(scale, "")
    return f"{value:,.1f}{unit}"


def read_excel(path: Path, default_scale: Optional[float] = None) -> Tuple[List[Fact], List[str]]:
    values = openpyxl.load_workbook(path, data_only=True)
    formulas = openpyxl.load_workbook(path, data_only=False)
    facts: List[Fact] = []
    warnings: List[str] = []
    uncached = 0

    for ws in values.worksheets:
        cells = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    cells[(cell.row, cell.column)] = cell.value
                elif isinstance(formulas[ws.title][cell.coordinate].value, str) and str(
                    formulas[ws.title][cell.coordinate].value
                ).startswith("="):
                    uncached += 1
        headers = _header_rows(cells)

        sheet_scale = None
        for (r, _), v in sorted(cells.items()):
            if r <= 10 and isinstance(v, str) and _note_scale(v):
                sheet_scale = _note_scale(v)
                break
        first_new = len(facts)

        for (r, c), v in sorted(cells.items()):
            if not _is_number(v) or (r in headers and c in headers[r]):
                continue

            label = ""
            for cc in range(c - 1, 0, -1):
                left = cells.get((r, cc))
                if isinstance(left, str) and left.strip() and not full_period(left):
                    label = left.strip()
                    break

            col_text = ""
            for rr in range(r - 1, 0, -1):
                above = cells.get((rr, c))
                if above is None:
                    break
                if isinstance(above, str) and not full_period(above):
                    col_text = above.strip()
                    break
            if _note_scale(col_text):
                col_text = ""
            if not label:
                label, col_text = col_text, ""

            period = None
            for rr in sorted((h for h in headers if h < r), reverse=True):
                if c in headers[rr]:
                    period = headers[rr][c]
                    break
            period = period or period_in(label) or period_in(col_text)

            fmt = ws.cell(row=r, column=c).number_format
            kind, shown = _classify(f"{label} {col_text}", fmt, float(v))
            if kind in (PERCENT, MULTIPLE, PER_SHARE):
                scale = 1.0
            elif kind == COUNT and not re.search("shares", label, re.I):
                scale = _label_scale(label) or 1.0
            else:
                scale = _label_scale(label) or _label_scale(col_text) or sheet_scale or default_scale or 1.0

            full_label = f"{label} / {col_text}" if col_text else label
            facts.append(Fact(
                location=f"{ws.title}!{ws.cell(row=r, column=c).coordinate}",
                label=full_label,
                metric=canon(label) if label else "",
                period=period,
                kind=kind,
                value=round(shown, 10),
                scale=scale,
                tol=abs(shown) * 1e-9,
                raw=_show_excel(kind, shown, scale),
                context=col_text,
            ))

        if sheet_scale is None and default_scale is None and any(f.kind == MONEY for f in facts[first_new:]):
            warnings.append(
                f'Sheet "{ws.title}" has no units note such as "($ in millions)"; '
                "treating its values as whole units. Set the model units to override."
            )

    if uncached:
        warnings.append(
            f"{uncached} formula cell(s) have no saved value and were skipped. "
            "Open the model in Excel and save it so calculated values are stored."
        )
    return facts, warnings


class Model:
    def __init__(self, facts: List[Fact]):
        self.facts = facts
        self.index: Dict[str, List[Fact]] = defaultdict(list)
        terms = dict(ALIAS_LOOKUP)
        for f in facts:
            if not f.metric:
                continue
            self.index[f.metric].append(f)
            if f.context:
                self.index[cross_key(f.label.split(" / ")[0], f.context)].append(f)
                self.index[cross_key(f.context, f.label.split(" / ")[0])].append(f)
            term = norm(f.label.split(" / ")[0])
            if term and term not in VOCAB_STOPLIST and len(term) >= 2 and len(term.split()) <= 6:
                terms.setdefault(term, f.metric)
        self.terms = terms
        ordered = sorted(terms, key=len, reverse=True)
        pattern = "|".join(r"[\s/\-.&]+".join(map(re.escape, t.split())) for t in ordered)
        self.term_re = re.compile(rf"(?<!\w)(?:{pattern})(?!\w)", re.I)
        self.known = set(ALIASES) | set(self.index)
        self.scales = [f.scale for f in facts if f.kind == MONEY]

    def find_terms(self, masked: str) -> List[Tuple[int, int, str, str]]:
        hits = []
        for m in self.term_re.finditer(masked):
            term = re.sub(r"[^a-z0-9]+", " ", m.group(0).lower()).strip()
            if len(term.replace(" ", "")) <= 3 and m.group(0).islower():
                continue
            if term in self.terms:
                hits.append((m.start(), m.end(), self.terms[term], m.group(0)))
        return hits

    def metric_from_label(self, label: str) -> str:
        key = canon(label)
        if key in self.known:
            return key
        hits = self.find_terms(mask(label))
        if hits:
            return max(hits, key=lambda h: h[1] - h[0])[2]
        return key

    def grid_metric(self, row: str, col: str) -> str:
        """Pick the key for a table or chart value from its row and column labels."""
        if row and col:
            for key in (cross_key(row, col), cross_key(col, row)):
                if key in self.index:
                    return key
        for label in (row, col):
            if label:
                key = self.metric_from_label(label)
                if key in self.known:
                    return key
        return canon(row or col)

    def refine(self, base: str, kind: str, context: str) -> str:
        """'EBITDA ... 22.7% margin' is EBITDA margin; '12.4x EBITDA' is EV/EBITDA."""
        if kind == PERCENT:
            if re.search("margin", context, re.I) and f"{base} margin" in self.known:
                return f"{base} margin"
            if GROWTH_RE.search(context) and f"{base} growth" in self.known:
                return f"{base} growth"
        if kind == MULTIPLE and base in ("ebitda", "revenue"):
            if base == "ebitda" and re.search("leverage|net debt", context, re.I):
                return "net leverage"
            return f"ev/{base}"
        return base

    @property
    def default_scale(self) -> float:
        return max(set(self.scales), key=self.scales.count) if self.scales else 1.0


# ---------------------------------------------------------------------------
# PowerPoint
# ---------------------------------------------------------------------------

LEFT_GAP_RE = re.compile(r"[\s:=|\u2013-]*(?:of|was|is|at|to|reached|totaled|totaling|approximately|about)?[\s:=~\u2013-]*", re.I)
RIGHT_GAP_RE = re.compile(r"\s*(?:of|in)?\s*", re.I)
PERIOD_GAP_RE = re.compile(r"\s*(?:in|for|during|of)?\s*\(?\s*", re.I)
LABEL_STOP = {"across", "in", "for", "of", "and", "with", "to", "from", "at", "by", "on", "the", "a", "an",
              "was", "is", "as", "vs", "versus", "per", "up", "down", "over"}
SKIPPED_PLACEHOLDERS = (PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER)
INCH = 914400


def _clean(text: str) -> str:
    return text.replace("\x0b", " ").strip()


def _flatten(shapes, container=0) -> Iterable[Tuple[object, int]]:
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _flatten(shape.shapes, shape.shape_id)
        else:
            yield shape, container


def _pure_number(text: str) -> bool:
    nums = parse_numbers(text)
    if len(nums) != 1:
        return False
    masked = mask(text)
    return not re.search("[A-Za-z]", masked[: nums[0].start] + masked[nums[0].end :])


def _decimals_from_format(fmt: str, value: float) -> int:
    if fmt and fmt.lower() != "general":
        section = re.sub(r'"[^"]*"', "", fmt.split(";")[0])
        if re.search("[0#]", section):
            return len(re.findall("[0#]", section.split(".")[1])) if "." in section else 0
    text = repr(round(value, 6))
    return len(text.split(".")[1].rstrip("0")) if "." in text else 0


def _grid_label(row: str, col: str, metric: str) -> str:
    parts = (row, col) if "|" in metric or not row else (row,)
    return " / ".join(re.sub(r"\s*\([^)]*\)", "", t).strip() for t in parts if t)


class SlideReader:
    def __init__(self, number: int, slide, model: Model):
        self.number = number
        self.model = model
        self.facts: List[Fact] = []
        self.shapes = []
        for shape, container in _flatten(slide.shapes):
            if shape.is_placeholder and shape.placeholder_format.type in SKIPPED_PLACEHOLDERS:
                continue
            self.shapes.append((shape, container))
        title = slide.shapes.title
        self.title_id = title.shape_id if title is not None else None
        self.title_period = period_in(title.text_frame.text) if title is not None else None
        self.note_scale = self._find_note_scale()

    def _find_note_scale(self) -> Optional[float]:
        for shape, _ in self.shapes:
            texts = []
            if shape.has_text_frame:
                texts.append(shape.text_frame.text)
            if getattr(shape, "has_table", False) and shape.has_table:
                texts.extend(c.text for r in shape.table.rows for c in r.cells)
            if getattr(shape, "has_chart", False) and shape.has_chart and shape.chart.has_title:
                texts.append(shape.chart.chart_title.text_frame.text)
            for text in texts:
                if _note_scale(text):
                    return _note_scale(text)
        return None

    def read(self) -> List[Fact]:
        for shape, container in self.shapes:
            if getattr(shape, "has_table", False) and shape.has_table:
                self._read_table(shape)
            elif getattr(shape, "has_chart", False) and shape.has_chart:
                self._read_chart(shape)
            elif shape.has_text_frame and shape.text_frame.text.strip():
                self._read_text(shape, container)
        return self.facts

    def _add(self, p: Parsed, label, metric, period, location, quote, strong=False, sign_free=False, context=""):
        self.facts.append(Fact(
            location=location, label=label, metric=metric, period=period, kind=p.kind, value=p.value,
            scale=p.scale, tol=p.tol, raw=p.raw, slide=self.number, note_scale=self.note_scale,
            strong=strong, sign_free=sign_free, context=context, quote=quote,
        ))

    # Tables -----------------------------------------------------------------

    def _read_table(self, shape):
        rows = [[_clean(c.text) for c in r.cells] for r in shape.table.rows]
        header = rows[0]
        has_header = any(re.search("[A-Za-z]", t) or full_period(t) for t in header[1:])
        for r, row in enumerate(rows):
            if r == 0 and has_header:
                continue
            for c, text in enumerate(row):
                if c == 0 or not text:
                    continue
                row_label = next((row[i] for i in range(c) if re.search("[A-Za-z]", row[i])), "")
                col_label = header[c] if has_header else ""
                period = (full_period(col_label) or full_period(row_label) or period_in(row_label)
                          or period_in(col_label) or self.title_period)
                row_text = "" if full_period(row_label) else row_label
                col_text = "" if full_period(col_label) else col_label
                metric = self.model.grid_metric(row_text, col_text)
                label = _grid_label(row_text, col_text, metric)
                location = f'table "{shape.name}" row {r + 1}, col {c + 1}'
                for p in parse_numbers(text, keep_bare_years=True):
                    shown_col = col_label if full_period(col_label) or "|" in metric else ""
                    where = ", ".join(t for t in (row_label, shown_col) if t)
                    quote = (f"{where}: {text[:p.start]}" if where else text[:p.start], text[p.start:p.end], text[p.end:])
                    self._add(p, label, metric, period, location, quote, strong=True, context=text)

    # Charts -----------------------------------------------------------------

    def _read_chart(self, shape):
        chart = shape.chart
        title = chart.chart_title.text_frame.text if chart.has_title and chart.chart_title.has_text_frame else ""
        for plot in chart.plots:
            categories = [str(c) for c in plot.categories]
            for series in plot.series:
                name = series.name or title
                codes = series._element.xpath("./c:val//c:formatCode/text()")
                fmt = str(codes[0]) if codes else ""
                for category, v in zip(categories, series.values):
                    if v is None:
                        continue
                    period = full_period(category) or self.title_period
                    row, col = (name, "") if full_period(category) else (category, name)
                    metric = self.model.grid_metric(row, col)
                    label = _grid_label(row, col, metric)
                    parsed = self._chart_value(float(v), fmt, f"{name} {title}")
                    scale = _label_scale(name) or _note_scale(name) or _note_scale(title)
                    if parsed.kind == PLAIN and scale:
                        parsed.scale = None
                    location = f'chart "{shape.name}", series "{name}", {category}'
                    self._add(parsed, label, metric, period, location, (f"{title or name}, {category}: ", parsed.raw, ""),
                              strong=True)
                    if scale:
                        self.facts[-1].note_scale = scale

    @staticmethod
    def _chart_value(v: float, fmt: str, label: str) -> Parsed:
        decimals = _decimals_from_format(fmt, v)
        if "%" in fmt:
            kind, v, decimals = PERCENT, v * 100, max(decimals, 0)
        elif re.search(r'x"?\s*(?:;|$)', fmt.lower()):
            kind = MULTIPLE
        elif PERCENT_LABEL_RE.search(label) and abs(v) <= 1.5:
            kind, v, decimals = PERCENT, v * 100, max(decimals - 2, 0)
        else:
            kind = PLAIN
        sign = {PERCENT: "%", MULTIPLE: "x"}.get(kind, "")
        return Parsed(round(v, 10), decimals, kind, None, f"{v:,.{decimals}f}{sign}", tol=0.5 * 10 ** -decimals)

    # Text -------------------------------------------------------------------

    def _read_text(self, shape, container):
        paragraphs = [_clean(p.text) for p in shape.text_frame.paragraphs if _clean(p.text)]
        # A period in one bullet says nothing about its siblings; a heading line does.
        heading = paragraphs[0] if paragraphs and not parse_numbers(paragraphs[0]) else ""
        shape_period = period_in(heading) or self.title_period
        location = f'text "{shape.name}"'
        for i, para in enumerate(paragraphs):
            if _pure_number(para):
                self._read_kpi(para, i, paragraphs, shape, container, location, shape_period)
            else:
                self._read_sentences(para, location, shape_period)

    def _read_kpi(self, para, i, paragraphs, shape, container, location, shape_period):
        label = ""
        for j in (i + 1, i - 1):
            if 0 <= j < len(paragraphs) and not _pure_number(paragraphs[j]):
                label = paragraphs[j]
                break
        label = label or self._nearest_label(shape, container)
        p = parse_numbers(para)[0]
        if not label:
            self._add(p, "", "", period_in(para) or shape_period, location, ("", para, ""), context=para)
            return
        metric = self.model.metric_from_label(label)
        metric = self.model.refine(metric, p.kind, label)
        period = period_in(para) or period_in(label) or shape_period
        self._add(p, label, metric, period, location, (f"{label}: ", para, ""), strong=True, context=f"{para} / {label}")

    def _nearest_label(self, shape, container) -> str:
        best, best_score = "", float("inf")
        for other, other_container in self.shapes:
            if other is shape or other_container != container or other.shape_id == self.title_id:
                continue
            if not other.has_text_frame:
                continue
            text = _clean(other.text_frame.text)
            if not text or not re.search("[A-Za-z]", mask(text)) or len(text.split()) > 8:
                continue
            dx = max(other.left - (shape.left + shape.width), shape.left - (other.left + other.width), 0)
            dy = max(other.top - (shape.top + shape.height), shape.top - (other.top + other.height), 0)
            if dx > INCH or dy > INCH or (dx and dy):
                continue
            score = dx + dy + (0 if other.top >= shape.top else INCH / 100)
            if score < best_score:
                best, best_score = text, score
        return best

    def _read_sentences(self, para, location, shape_period):
        for sentence in re.split(r"(?<=[;!?])\s+|(?<=[a-z%)])\.\s+", para):
            if sentence.strip():
                self._read_sentence(sentence, location, shape_period)

    def _read_sentence(self, sentence, location, shape_period):
        masked = mask(sentence)
        nums = parse_numbers(sentence)
        terms = self.model.find_terms(masked)
        periods = [(m.start(), m.end(), period_of(m)) for m in PERIOD_RE.finditer(sentence)]
        sentence_period = period_in(sentence)
        inherited = None
        used_periods = set()
        for i, n in enumerate(nums):
            lo = nums[i - 1].end if i else 0
            hi = nums[i + 1].start if i + 1 < len(nums) else len(sentence)
            left = [t for t in terms if lo <= t[0] and t[1] <= n.start]
            right = [t for t in terms if n.end <= t[0] and t[1] <= hi]

            term = None
            if left and LEFT_GAP_RE.fullmatch(masked[left[-1][1]:n.start]):
                term = left[-1]
            elif right and RIGHT_GAP_RE.fullmatch(masked[n.end:right[0][0]]):
                term = right[0]
            elif left:
                term = left[-1]
            elif not inherited and right:
                term = right[0]

            period = None
            after = [p for p in periods if n.end <= p[0] and p[1] <= hi]
            before = [p for p in periods if lo <= p[0] and p[1] <= n.start and p[0] not in used_periods]
            if after and PERIOD_GAP_RE.fullmatch(sentence[n.end:after[0][0]]):
                period = after[0][2]
                used_periods.add(after[0][0])
            elif before:
                period = before[-1][2]
            elif after:
                period = after[0][2]
            period = period or sentence_period or shape_period

            context = masked[lo:hi]
            base = term[2] if term else inherited
            if base:
                inherited = base
                metric = self.model.refine(base, n.kind, context)
                label = term[3] if term and metric == base else metric
            else:
                label = self._words_near(masked, lo, n, hi)
                metric = canon(label) if label else ""
            sign_free = bool(DECLINE_RE.search(context))
            quote = (sentence[:n.start], sentence[n.start:n.end], sentence[n.end:])
            self._add(n, label, metric, period, location, quote, sign_free=sign_free, context=sentence)

    @staticmethod
    def _words_near(masked: str, lo: int, n: Parsed, hi: int) -> str:
        words = re.findall(r"[A-Za-z][\w&'-]*", masked[n.end:hi])
        if re.match(r"\s*[A-Za-z]", masked[n.end:hi]):
            label = []
            for word in words[:3]:
                if word.lower() in LABEL_STOP:
                    break
                label.append(word)
            if label:
                return " ".join(label)
        left = re.findall(r"[A-Za-z][\w&'-]*", masked[lo:n.start])
        while left and left[-1].lower() in LABEL_STOP:
            left.pop()
        return " ".join(left[-3:])


def read_pptx(path: Path, model: Model) -> List[Fact]:
    facts: List[Fact] = []
    for number, slide in enumerate(Presentation(str(path)).slides, 1):
        facts.extend(SlideReader(number, slide, model).read())
    return facts


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------


def _scale_for(f: Fact, c: Fact) -> float:
    if f.scale:
        return f.scale
    if c.kind in (PERCENT, MULTIPLE, PER_SHARE):
        return 1.0
    if c.kind == COUNT:
        return c.scale  # a "$ in millions" note says nothing about headcount or share counts
    return f.note_scale or c.scale


def _sign_free(f: Fact, c: Fact) -> bool:
    return f.sign_free or bool(EXPENSE_RE.search(f.metric) or EXPENSE_RE.search(c.label))


def _gap(f: Fact, c: Fact) -> Tuple[float, float, float, float]:
    scale = _scale_for(f, c)
    a, b = f.value * scale, c.value * c.scale
    allowed = f.tol * scale + abs(b) * 1e-9 + 1e-9
    return a, b, allowed, abs(abs(a) - abs(b)) if _sign_free(f, c) else abs(a - b)


def values_match(f: Fact, c: Fact) -> bool:
    _, _, allowed, diff = _gap(f, c)
    return diff <= allowed


def format_like(f: Fact, c: Fact) -> str:
    """Write the model value the way the slide writes this number: $214mm, $1.2 billion, 2,400."""
    number = re.search(r"\d[\d,]*(?:\.\d+)?|\.\d+", f.raw)
    lead, trail = f.raw[: number.start()], f.raw[number.end() :]
    decimals = len(number.group(0).split(".")[1]) if "." in number.group(0) else 0
    shown = c.value * c.scale / _scale_for(f, c)
    if _sign_free(f, c):
        shown = abs(shown) if f.value >= 0 else -abs(shown)
    if re.search("bps|basis", trail, re.I):
        shown *= 100
    rounded = Decimal(str(round(abs(shown), 10))).quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)
    currency = re.search(r"US\$|USD\s?|[$€£]", lead)
    prefix = currency.group(0) if currency and c.kind != COUNT else ""
    body = f"{prefix}{rounded:,}{trail.replace(')', '').rstrip()}"
    if shown < 0 and rounded:
        return f"({body})" if "(" in lead else f"-{body}"
    return body


def display_name(f: Fact, model: "Model") -> str:
    if f.metric in DISPLAY_NAMES:
        return DISPLAY_NAMES[f.metric]
    if f.metric in ALIASES:
        return title_case(f.metric)
    return title_case(f.label or f.metric) or "Unlabeled Number"


def _diagnose(f: Fact, c: Fact, candidates: List[Fact]) -> str:
    a, b, allowed, _ = _gap(f, c)
    notes = []
    if abs(abs(a) - abs(b)) <= allowed:
        notes.append("sign differs from model")
    elif a and b:
        for factor in (10, 100, 1000, 1e6):
            if abs(abs(a) - abs(b) * factor) <= allowed or abs(abs(a) * factor - abs(b)) <= allowed * factor:
                notes.append(f"off by a factor of {factor:,.0f}; check units")
                break
        else:
            if c.kind in (PERCENT, MULTIPLE):
                unit = " pts" if c.kind == PERCENT else "x"
                notes.append(f"slide differs from model by {a - b:+.2f}{unit}")
            else:
                notes.append(f"slide is {abs(a - b) / abs(b):.1%} {'above' if a > b else 'below'} model")
    for other in candidates:
        if other is not c and other.period != f.period and other.period and values_match(f, other):
            notes.append(f"value matches {show_period(other.period)} in the model; wrong period?")
            break
    if not f.period and len({x.period for x in candidates}) > 1:
        notes.append("no period on slide; compared with closest model value")
    return "; ".join(notes)


def _candidates(f: Fact, model: Model) -> List[Fact]:
    found = model.index.get(f.metric, [])
    if not found and f.metric.startswith("adjusted "):
        found = model.index.get(f.metric[len("adjusted "):], [])
    return [c for c in found if c.kind in COMPATIBLE[f.kind]]


def compare_to_model(slides: List[Fact], model: Model) -> Tuple[List[Issue], List[Fact], int]:
    issues: List[Issue] = []
    review: List[Fact] = []
    matched = 0
    for f in slides:
        if f.kind == PLAIN and not f.strong:
            if not f.metric or f.metric not in model.index:
                continue
        candidates = _candidates(f, model) if f.metric else []
        if not candidates:
            f.status = "review"
            review.append(f)
            continue

        if f.period:
            scope = [c for c in candidates if c.period == f.period or c.period is None]
            if not scope:
                f.status = "review"
                review.append(f)
                continue
        else:
            scope = candidates

        hit = next((c for c in scope if values_match(f, c)), None)
        closest = hit or min(scope, key=lambda c: _gap(f, c)[3])
        f.best = closest
        if hit:
            f.status = "ok"
            matched += 1
            continue
        f.status, f.expected, f.note = "wrong", format_like(f, closest), _diagnose(f, closest, candidates)
        issues.append(Issue(
            type="slide_vs_model", item=f.label or f.metric, period=show_period(f.period), slide=f.slide,
            location=f.location, slide_value=f.raw, excel_cell=closest.location, excel_value=closest.raw,
            note=f.note,
        ))
    return issues, review, matched


def _bucket(f: Fact) -> str:
    kind = f.best.kind if f.best and f.kind == PLAIN else f.kind
    return kind if kind in (PERCENT, MULTIPLE) else "amount"


def _resolved(f: Fact, model: Model) -> float:
    if f.scale:
        scale = f.scale
    elif f.kind in (PERCENT, MULTIPLE) or (f.best and f.best.kind in (PERCENT, MULTIPLE, PER_SHARE)):
        scale = 1.0
    elif f.best and f.best.kind == COUNT:
        scale = f.best.scale
    else:
        scale = f.note_scale or (f.best.scale if f.best else None) or model.default_scale
    return scale


def _conflicting(f: Fact, g: Fact, model: Model) -> bool:
    if _bucket(f) != _bucket(g):
        return False
    sf, sg = _resolved(f, model), _resolved(g, model)
    a, b = f.value * sf, g.value * sg
    if f.sign_free or g.sign_free or EXPENSE_RE.search(f.metric):
        a, b = abs(a), abs(b)
    return abs(a - b) > f.tol * sf + g.tol * sg + 1e-9


def compare_across_slides(slides: List[Fact], model: Model) -> List[Issue]:
    groups: Dict[Tuple[str, Optional[str]], List[Fact]] = defaultdict(list)
    for f in slides:
        if f.metric:
            groups[(f.metric, f.period)].append(f)

    issues: List[Issue] = []
    seen = set()
    for (metric, period), facts in groups.items():
        # Without a period, only compare items the model holds at a single point in time.
        if period is None and len({c.period for c in model.index.get(metric, [])}) > 1:
            continue
        for f, g in combinations(sorted(facts, key=lambda x: x.slide), 2):
            if f.slide == g.slide or not _conflicting(f, g, model):
                continue
            key = (metric, period, f.slide, f.raw, g.slide, g.raw)
            if key in seen:
                continue
            seen.add(key)
            best = f.best or g.best
            issues.append(Issue(
                type="slide_vs_slide", item=f.label or g.label or metric, period=show_period(period),
                slide=f.slide, location=f.location, slide_value=f.raw,
                other_slide=g.slide, other_location=g.location, other_value=g.raw,
                excel_cell=best.location if best else "not in model", excel_value=best.raw if best else "",
            ))
    return sorted(issues, key=lambda i: (i.slide, i.other_slide))


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------


def render_report(deck, workbook, slides, excel, warnings, model_issues, cross_issues, review, matched) -> str:
    lines = [
        "PITCHBOOK CHECK",
        f"Deck:  {deck}",
        f"Model: {workbook}",
        "",
        f"Numbers on slides:         {len(slides)}",
        f"Numbers in model:          {len(excel)}",
        f"Tied to model:             {matched}",
        f"Slide vs model mismatches: {len(model_issues)}",
        f"Slide vs slide conflicts:  {len(cross_issues)}",
        f"Could not be checked:      {len(review)}",
    ]
    if warnings:
        lines += ["", "WARNINGS"] + [f"  ! {w}" for w in warnings]

    lines += ["", "=" * 72, f"SLIDE VS MODEL MISMATCHES ({len(model_issues)})", "=" * 72]
    for n, i in enumerate(sorted(model_issues, key=lambda x: x.slide), 1):
        lines += [
            f"[{n}] {i.item} ({i.period})",
            f"    Slide {i.slide}: {i.slide_value}  [{i.location}]",
            f"    Excel {i.excel_cell}: {i.excel_value}",
            f"    {i.note}",
            "",
        ]
    if not model_issues:
        lines += ["None.", ""]

    lines += ["=" * 72, f"SAME ITEM, DIFFERENT VALUES ACROSS SLIDES ({len(cross_issues)})", "=" * 72]
    for n, i in enumerate(cross_issues, 1):
        excel_line = f"{i.excel_cell}: {i.excel_value}" if i.excel_value else i.excel_cell
        lines += [
            f"[{n}] {i.item} ({i.period})",
            f"    Slide {i.slide}: {i.slide_value}  [{i.location}]",
            f"    Slide {i.other_slide}: {i.other_value}  [{i.other_location}]",
            f"    Excel {excel_line}",
            "",
        ]
    if not cross_issues:
        lines += ["None.", ""]

    lines += ["=" * 72, f"COULD NOT BE CHECKED: NO MATCHING MODEL CELL ({len(review)})", "=" * 72]
    for f in sorted(review, key=lambda x: x.slide):
        lines.append(f"Slide {f.slide}: {f.raw}  {f.label or '(no label)'}, {show_period(f.period)}  [{f.location}]")
    if not review:
        lines.append("None.")
    return "\n".join(lines) + "\n"


FONT_LINK = (
    '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">'
)

REPORT_CSS = """
* { box-sizing: border-box; }
body { margin: 0; background: #000; color: #e8eaed;
       font: 16px/1.5 Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif; }
main { max-width: 960px; margin: 0 auto; padding: 48px 16px 72px; }
h1 { font-size: 28px; font-weight: 600; letter-spacing: -0.01em; margin: 0 0 6px; }
h2 { font-size: 18px; font-weight: 600; margin: 40px 0 16px; }
.files { color: #9aa3ae; font-size: 14px; margin-bottom: 28px; overflow-wrap: anywhere; }
.warn { background: #16181c; border: 1px solid #2a2d33; border-radius: 10px; padding: 12px 16px; margin-bottom: 16px; }
"""

CARD_CSS = """
.pc-card, .pc-summary, .pc-clear {
  font-family: Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
  color: #e8eaed; border-radius: 10px; margin-bottom: 16px; }
.pc-card, .pc-summary { background: #16181c; border: 1px solid #2a2d33;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.5), 0 6px 18px rgba(255, 255, 255, 0.03); }
.pc-card { padding: 22px 26px; border-left: 4px solid #e07a5f; }
.pc-card.pc-missing { border-left-color: #f2cc8f; }
.pc-anchor { display: block; height: 0; scroll-margin-top: 72px; }
.pc-head { display: flex; flex-wrap: wrap; align-items: center; gap: 8px 10px; margin-bottom: 18px; }
.pc-title { font-size: 16px; font-weight: 600; overflow-wrap: anywhere; }
.pc-period { font-size: 13px; color: #9aa3ae; }
.pc-tag { margin-left: auto; font-size: 12px; font-weight: 600; padding: 3px 10px; border-radius: 999px;
  background: rgba(224, 122, 95, 0.16); color: #f2a893; }
.pc-missing .pc-tag { background: rgba(242, 204, 143, 0.14); color: #f2cc8f; }
.pc-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 16px 24px; }
.pc-label { font-size: 12px; font-weight: 500; color: #9aa3ae; margin-bottom: 4px; }
.pc-value { font-size: 16px; font-weight: 500; font-variant-numeric: tabular-nums; overflow-wrap: anywhere; }
.pc-value.pc-strong { font-size: 19px; font-weight: 600; }
.pc-value.pc-empty { color: #858e9a; font-weight: 400; }
.pc-foot { margin-top: 18px; padding-top: 14px; border-top: 1px solid #2a2d33; font-size: 13px; color: #9aa3ae;
  overflow-wrap: anywhere; }
.pc-foot b { color: #c9cfd6; font-weight: 500; }
.pc-quotes { margin-top: 18px; padding-top: 6px; border-top: 1px solid #2a2d33; }
.pc-row { display: flex; flex-wrap: wrap; align-items: baseline; gap: 2px 14px; padding: 8px 0; }
.pc-row-slide { flex: 0 0 56px; font-size: 13px; font-weight: 500; color: #9aa3ae; }
.pc-quote { flex: 1 1 240px; min-width: 0; font-size: 14px; color: #d5dae0; overflow-wrap: anywhere; }
.pc-hit { background: rgba(224, 122, 95, 0.22); color: #f7b8a6; font-weight: 600; padding: 1px 5px; border-radius: 4px; }
.pc-hit-amber { background: rgba(242, 204, 143, 0.18); color: #f2cc8f; }
.pc-status { font-size: 12px; font-weight: 500; white-space: nowrap; }
.pc-status-ok { color: #81b29a; }
.pc-status-wrong { color: #f2a893; }
.pc-status-unverified { color: #f2cc8f; }
.pc-fix { font-size: 14px; color: #e8eaed; }
.pc-fix b { color: #e8eaed; font-weight: 600; }
.pc-note { margin-top: 6px; }
.pc-jumps { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; margin-top: 16px; }
.pc-jump-label { font-size: 13px; color: #9aa3ae; margin-right: 2px; }
.pc-jump { font-size: 13px; font-weight: 500; color: #e8eaed !important; text-decoration: none !important;
  border: 1px solid #3a3e46; border-radius: 999px; padding: 4px 12px; }
.pc-jump:hover { border-color: #e07a5f; }
.pc-summary { padding: 24px 26px; margin-bottom: 28px; }
.pc-headline { font-size: 22px; font-weight: 600; letter-spacing: -0.01em; margin-bottom: 12px; }
.pc-chips { display: flex; flex-wrap: wrap; gap: 8px 24px; font-size: 14px; color: #b8c0ca; }
.pc-chip { display: inline-flex; align-items: center; gap: 8px; }
.pc-dot { width: 9px; height: 9px; border-radius: 50%; background: #e07a5f; }
.pc-dot.pc-amber { background: #f2cc8f; }
.pc-dot.pc-green { background: #81b29a; }
.pc-clear { background: #81b29a; color: #12312a; padding: 24px 26px; margin-bottom: 28px; }
.pc-clear .pc-headline { margin-bottom: 4px; }
.pc-clear p { margin: 0; font-size: 15px; }
.pc-section { font-family: Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
  font-size: 13px; font-weight: 600; color: #b8c0ca; margin: 28px 0 14px; }
@media (max-width: 640px) {
  .pc-card, .pc-summary, .pc-clear { padding: 18px; }
  .pc-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
}
"""


@dataclass
class Occurrence:
    slide: int
    shown: str
    quote: Tuple[str, str, str]
    status: str  # ok, wrong or unverified


@dataclass
class Finding:
    style: str  # mismatch, conflict or missing
    item: str
    period: str
    occurrences: List[Occurrence]
    slide_value: str
    excel_cell: str
    model_value: str
    fix: str
    note: str = ""
    anchors: List[int] = field(default_factory=list)

    @property
    def affected(self) -> List[int]:
        return sorted({o.slide for o in self.occurrences if o.status != "ok"})


FINDING_TAGS = {"mismatch": "Value mismatch", "conflict": "Slide conflict", "missing": "Not in model"}
FINDING_SECTIONS = [
    ("mismatch", "Value mismatches against the model"),
    ("conflict", "Same item, different values across slides"),
    ("missing", "Numbers not found in the model"),
]
STATUS_TEXT = {"ok": "Ties to model", "wrong": "Does not match", "unverified": "Not verified"}


def _slide_list(numbers: Iterable[int]) -> str:
    items = sorted(set(numbers))
    if len(items) == 1:
        return f"slide {items[0]}"
    return f"slides {', '.join(map(str, items[:-1]))} and {items[-1]}"


def _fix_text(style: str, wrong: List[Fact], facts: List[Fact], ambiguous: bool, model_value: str) -> str:
    if style == "mismatch" and ambiguous:
        return (f"Confirm which period {_slide_list(f.slide for f in wrong)} refers to. "
                f"The closest model value is {model_value}.")
    if style == "mismatch":
        changes: Dict[Tuple[str, str], List[int]] = defaultdict(list)
        for f in wrong:
            changes[(f.raw, f.expected)].append(f.slide)
        return " ".join(f"Change {old} to {new} on {_slide_list(where)}." for (old, new), where in changes.items())
    if style == "conflict":
        return f"Confirm the correct figure, then use it on {_slide_list(f.slide for f in facts)}."
    values = ", ".join(dict.fromkeys(f.raw for f in facts))
    return f"Add this item to the model, or confirm the source of {values} on {_slide_list(f.slide for f in facts)}."


def build_findings(slides: List[Fact], model: Model) -> List[Finding]:
    """Merge model mismatches and slide conflicts into one finding per item and period."""
    groups: Dict[tuple, List[Fact]] = defaultdict(list)
    for n, f in enumerate(slides):
        if f.metric:
            several = f.period is None and len({c.period for c in model.index.get(f.metric, [])}) > 1
            groups[(f.metric, f.period, _bucket(f), f.best.location if several and f.best else "")].append(f)
        elif f.status == "review":
            groups[("", n)].append(f)

    findings: List[Finding] = []
    for key, members in groups.items():
        facts = sorted(members, key=lambda x: x.slide)
        ambiguous = facts[0].period is None and len({c.period for c in model.index.get(facts[0].metric, [])}) > 1
        wrong = [f for f in facts if f.status == "wrong"]
        conflict = not ambiguous and any(
            f.slide != g.slide and _conflicting(f, g, model) for f, g in combinations(facts, 2)
        )
        if wrong:
            style = "mismatch"
        elif conflict:
            style = "conflict"
        elif any(f.status == "review" for f in facts):
            style = "missing"
        else:
            continue

        best = next((f.best for f in wrong + facts if f.best), None)
        model_value = "n/a"
        if wrong:
            model_value = wrong[0].expected
        elif best:
            model_value = format_like(facts[0], best)

        occurrences: List[Occurrence] = []
        seen = set()
        for f in facts:
            mark = (f.slide, f.raw, f.quote)
            if mark not in seen:
                seen.add(mark)
                status = f.status if f.status in ("ok", "wrong") else "unverified"
                occurrences.append(Occurrence(f.slide, f.raw, f.quote, status))
        highlighted = wrong or facts
        findings.append(Finding(
            style=style, item=display_name(facts[0], model), period=show_period(facts[0].period),
            occurrences=occurrences, slide_value=", ".join(dict.fromkeys(f.raw for f in highlighted)),
            excel_cell=best.location if best else "", model_value=model_value,
            fix=_fix_text(style, wrong, facts, ambiguous, model_value),
            note=wrong[0].note if wrong else "",
        ))

    rank = {style: i for i, (style, _) in enumerate(FINDING_SECTIONS)}
    findings.sort(key=lambda x: (rank[x.style], x.affected[0] if x.affected else 0))
    placed = set()
    for finding in findings:
        finding.anchors = [s for s in finding.affected if s not in placed]
        placed.update(finding.anchors)
    return findings


def _field(label: str, value: str, strong: bool = False, empty: bool = False) -> str:
    classes = "pc-value" + (" pc-strong" if strong else "") + (" pc-empty" if empty else "")
    return f'<div><div class="pc-label">{escape(label)}</div><div class="{classes}">{escape(value)}</div></div>'


def _sentence(note: str) -> str:
    return escape(note[:1].upper() + note[1:])


def _occurrence_row(o: Occurrence, style: str) -> str:
    before, number, after = o.quote
    if o.status == "ok":
        quote = escape(before + number + after)
    else:
        kind = " pc-hit-amber" if style == "missing" else ""
        quote = f'{escape(before)}<mark class="pc-hit{kind}">{escape(number)}</mark>{escape(after)}'
    return (f'<div class="pc-row"><span class="pc-row-slide">Slide {o.slide}</span>'
            f'<span class="pc-quote">{quote}</span>'
            f'<span class="pc-status pc-status-{o.status}">{STATUS_TEXT[o.status]}</span></div>')


def finding_card(finding: Finding) -> str:
    slides = sorted({o.slide for o in finding.occurrences})
    has_cell = bool(finding.excel_cell)
    fields = [
        _field("Slide" if len(slides) == 1 else "Slides", ", ".join(map(str, slides))),
        _field("Excel cell", finding.excel_cell or "Not in model", empty=not has_cell),
        _field("Slide value", finding.slide_value, strong=True),
        _field("Model value", finding.model_value, strong=has_cell, empty=not has_cell),
    ]
    anchors = "".join(f'<span class="pc-anchor" id="slide-{s}"></span>' for s in finding.anchors)
    period = f'<span class="pc-period">{escape(finding.period)}</span>' if finding.period != "no period" else ""
    rows = "".join(_occurrence_row(o, finding.style) for o in finding.occurrences)
    note = f'<div class="pc-note">{_sentence(finding.note)}</div>' if finding.note else ""
    return (f'{anchors}<article class="pc-card pc-{finding.style}"><div class="pc-head">'
            f'<span class="pc-title">{escape(finding.item)}</span>{period}'
            f'<span class="pc-tag">{FINDING_TAGS[finding.style]}</span></div>'
            f'<div class="pc-grid">{"".join(fields)}</div><div class="pc-quotes">{rows}</div>'
            f'<div class="pc-foot"><div class="pc-fix"><b>Suggested fix:</b> {escape(finding.fix)}</div>{note}</div>'
            f"</article>")


def findings_html(findings: List[Finding]) -> str:
    parts = []
    for style, title in FINDING_SECTIONS:
        cards = [finding_card(f) for f in findings if f.style == style]
        if cards:
            parts.append(f'<div class="pc-section">{escape(title)} ({len(cards)})</div>')
            parts.extend(cards)
    return "".join(parts)


def _plural(count: int, noun: str) -> str:
    return f"{count} {noun}{'es' if count != 1 and noun.endswith('ch') else 's' if count != 1 else ''}"


def summary_html(findings: List[Finding], matched: int) -> str:
    if not findings and not matched:
        return '<section class="pc-summary"><div class="pc-headline">No numbers could be checked</div></section>'
    if not findings:
        return ('<section class="pc-clear"><div class="pc-headline">All clear</div>'
                f"<p>All {_plural(matched, 'number')} tie to the model and agree across slides.</p></section>")
    affected = sorted({s for f in findings for s in f.affected})
    count = lambda style: sum(1 for f in findings if f.style == style)
    chips = [
        ("", _plural(count("mismatch"), "value mismatch")),
        ("", _plural(count("conflict"), "slide conflict")),
        (" pc-amber", f"{count('missing')} not found in model"),
        (" pc-green", f"{matched} tied to model"),
    ]
    chip_html = "".join(f'<span class="pc-chip"><span class="pc-dot{c}"></span>{t}</span>' for c, t in chips)
    jumps = "".join(f'<a class="pc-jump" href="#slide-{s}">Slide {s}</a>' for s in affected)
    return (f'<section class="pc-summary"><div class="pc-headline">{_plural(len(findings), "issue")} found on '
            f'{_plural(len(affected), "slide")}</div><div class="pc-chips">{chip_html}</div>'
            f'<div class="pc-jumps"><span class="pc-jump-label">Go to</span>{jumps}</div></section>')


def render_html(deck, workbook, warnings, findings, matched) -> str:
    parts = [
        '<!doctype html><html lang="en"><head><meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1">',
        FONT_LINK,
        f"<title>Pitchbook Check</title><style>{REPORT_CSS}{CARD_CSS}</style></head><body><main>",
        "<h1>Pitchbook Check</h1>",
        f'<div class="files">{escape(str(deck))} checked against {escape(str(workbook))}</div>',
        summary_html(findings, matched),
    ]
    parts.extend(f'<div class="warn">{escape(w)}</div>' for w in warnings)
    parts.append(findings_html(findings))
    parts.append("</main></body></html>")
    return "\n".join(parts)


def write_csv(path: Path, header: List[str], rows: Iterable[List]) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        writer.writerows(rows)


@dataclass
class Result:
    slides: List[Fact]
    excel: List[Fact]
    warnings: List[str]
    model_issues: List[Issue]
    cross_issues: List[Issue]
    review: List[Fact]
    matched: int = 0
    report: str = field(default="", repr=False)
    html: str = field(default="", repr=False)
    findings: List[Finding] = field(default_factory=list)


def check(deck: Path, workbook: Path, default_scale: Optional[float] = None) -> Result:
    excel, warnings = read_excel(workbook, default_scale)
    model = Model(excel)
    slides = read_pptx(deck, model)
    model_issues, review, matched = compare_to_model(slides, model)
    cross_issues = compare_across_slides(slides, model)
    report = render_report(deck, workbook, slides, excel, warnings, model_issues, cross_issues, review, matched)
    findings = build_findings(slides, model)
    return Result(slides, excel, warnings, model_issues, cross_issues, review, matched, report,
                  render_html(deck, workbook, warnings, findings, matched), findings)


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Check a PowerPoint pitchbook against an Excel model.")
    parser.add_argument("deck", type=Path, help="pitchbook .pptx")
    parser.add_argument("model", type=Path, help="financial model .xlsx")
    parser.add_argument("--out", type=Path, default=Path("report"), help="output folder (default: report)")
    parser.add_argument("--model-units", choices=["ones", "thousands", "millions", "billions"],
                        help="units for sheets with no note such as ($ in millions)")
    parser.add_argument("--no-open", action="store_true", help="do not open the HTML report in a browser")
    args = parser.parse_args(argv)

    for path in (args.deck, args.model):
        if not path.is_file():
            print(f"File not found: {path}", file=sys.stderr)
            return 2

    default = {"ones": 1.0, "thousands": 1e3, "millions": 1e6, "billions": 1e9}.get(args.model_units)
    result = check(args.deck, args.model, default)

    args.out.mkdir(parents=True, exist_ok=True)
    (args.out / "report.txt").write_text(result.report, encoding="utf-8")
    html_path = (args.out / "report.html").resolve()
    html_path.write_text(result.html, encoding="utf-8")
    columns = ["type", "item", "period", "slide", "location", "slide_value", "excel_cell", "excel_value",
               "other_slide", "other_location", "other_value", "note"]
    write_csv(args.out / "mismatches.csv", columns,
              ([getattr(i, c) for c in columns] for i in result.model_issues + result.cross_issues))
    write_csv(args.out / "slide_numbers.csv", ["slide", "location", "label", "period", "value_as_shown", "kind", "context"],
              ([f.slide, f.location, f.label, f.period or "", f.raw, f.kind, f.context] for f in result.slides))
    write_csv(args.out / "excel_numbers.csv", ["cell", "label", "period", "value", "kind"],
              ([f.location, f.label, f.period or "", f.raw, f.kind] for f in result.excel))

    print(result.report)
    print(f"Files written to {args.out}/: report.html, report.txt, mismatches.csv, slide_numbers.csv, excel_numbers.csv")
    if not args.no_open:
        webbrowser.open(html_path.as_uri())
    return 1 if result.model_issues or result.cross_issues else 0


if __name__ == "__main__":
    sys.exit(main())
