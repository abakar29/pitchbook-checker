# Pitchbook Checker

Ties every number in a PowerPoint pitchbook to an Excel financial model and flags anything that does not match.

## Setup

    python3 -m venv .venv
    .venv/bin/pip install -r requirements.txt

## Use

    .venv/bin/python pitchbook_checker.py deck.pptx model.xlsx --out report

If a sheet has no units note such as `($ in millions)`, add `--model-units millions`.

| File in `report/` | Contents |
|---|---|
| `report.html` | Browser report with a red card per mismatch; opens automatically unless you pass `--no-open` |
| `report.txt` | Readable list of mismatches with slide, Excel cell and both values |
| `mismatches.csv` | The same mismatches as a table |
| `slide_numbers.csv` | Every number found on the slides, with label and location |
| `excel_numbers.csv` | Every number found in the model, with label and cell |

The exit code is 0 if clean, 1 if mismatches were found, 2 on bad input.

## Web app

    .venv/bin/streamlit run app.py

Upload a deck and a model, then click **Check**, or click **Try with sample files**. On Streamlit Community
Cloud, use `app.py` as the entrypoint. A public repo makes a public app, so do not upload confidential files to it.

## What it checks

- **Slide vs model:** finds the Excel cell for the same item and period, then compares values.
- **Slide vs slide:** the same item and period shown with different values on different slides.
- **Could not be checked:** numbers with no matching model cell, listed for manual review.

Text boxes, grouped shapes, tables, native charts and KPI cards are all read. Speaker notes and images are not.

## How units and rounding work

A displayed number stands for a range set by its precision. `$4.2B` means 4.15B to 4.25B, so 4,213 in a
`($ in millions)` sheet passes and 4,300 fails. Percent cells (0.227) compare with `22.7%`, `bps` converts to
percent, and accounting negatives like `($45.2)` are understood. Costs such as capex may differ in sign,
because models usually hold them as negatives.

## Good to know

- The model must have been saved by Excel. Formula cells with no stored result are skipped with a warning.
- Labels are matched through the `ALIASES` table in `pitchbook_checker.py` (Sales = Revenue, TEV = Enterprise
  Value). Add firm-specific terms there.
- Currency is not converted, so `EUR` and `$` figures are compared as plain amounts.

## Sample and tests

    .venv/bin/python make_samples.py
    .venv/bin/python -m pytest -q

`make_samples.py` writes a fictional 5-slide deck and model to `samples/`. Its `PLANTED` table lists each
deliberate error, and the tests require exactly those to be reported. They also require a corrected deck to
report none.

## Vercel version

`vercel/` is a separate deployable: a static upload page (`public/index.html`) plus a Python function
(`api/check.py`) that runs the same checker. In Vercel, import the repo and set **Root Directory** to `vercel`,
or run `vercel --prod` from inside that folder. Uploads are limited to about 4 MB combined by Vercel.
After editing `pitchbook_checker.py`, run `vercel/sync.sh` to refresh the copy the function uses
(a test fails if they drift apart).
