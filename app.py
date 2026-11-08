"""Streamlit front end for pitchbook_checker. Start with: streamlit run app.py"""
from __future__ import annotations

import logging
import tempfile
from pathlib import Path
from typing import Optional

import streamlit as st

from pitchbook_checker import CARD_CSS, check, findings_html, summary_html

SAMPLES = Path(__file__).resolve().parent / "samples"
SAMPLE_DECK = SAMPLES / "project_atlas_pitchbook.pptx"
SAMPLE_MODEL = SAMPLES / "project_atlas_model.xlsx"
UNITS = {"Detect from the file": None, "Ones": 1.0, "Thousands": 1e3, "Millions": 1e6, "Billions": 1e9}


def run_check(deck: Path, model: Path, names: str, scale: Optional[float]) -> None:
    """Store the result, or a readable error, for the results section to show."""
    st.session_state.pop("result", None)
    st.session_state.pop("error", None)
    try:
        with st.spinner("Checking every number..."):
            st.session_state["result"] = check(deck, model, scale)
        st.session_state["names"] = names
    except Exception:
        logging.exception("Check failed")
        st.session_state["error"] = (
            "These files could not be read. Make sure they are valid .pptx and .xlsx files "
            "and are not password protected."
        )


def show_results() -> None:
    if "error" in st.session_state:
        st.error(st.session_state["error"])
    if "result" not in st.session_state:
        return
    result = st.session_state["result"]

    st.divider()
    st.caption(st.session_state["names"])
    st.html(summary_html(result.findings, result.matched))
    for warning in result.warnings:
        st.warning(warning)
    if result.findings:
        st.html(findings_html(result.findings))


st.set_page_config(page_title="Pitchbook Checker", page_icon="📊", layout="centered")
st.markdown(f"<style>{CARD_CSS}</style>", unsafe_allow_html=True)

st.title("Pitchbook Checker")
st.write("Upload a pitchbook and its financial model to find numbers that do not match.")

left, right = st.columns(2)
deck_file = left.file_uploader("PowerPoint pitchbook", type=["pptx"], key="deck")
model_file = right.file_uploader("Excel model", type=["xlsx", "xlsm"], key="model")

with st.expander("Options"):
    units = st.selectbox("Model units, for sheets with no note like ($ in millions)", list(UNITS))

check_column, sample_column = st.columns(2)
check_clicked = check_column.button(
    "Check", key="check", type="primary", use_container_width=True, disabled=not (deck_file and model_file)
)
sample_clicked = sample_column.button("Try with sample files", key="sample", use_container_width=True)

if check_clicked:
    with tempfile.TemporaryDirectory() as folder:
        deck_path = Path(folder) / "deck.pptx"
        model_path = Path(folder) / f"model{Path(model_file.name).suffix}"
        deck_path.write_bytes(deck_file.getvalue())
        model_path.write_bytes(model_file.getvalue())
        run_check(deck_path, model_path, f"{deck_file.name} checked against {model_file.name}", UNITS[units])
elif sample_clicked:
    run_check(SAMPLE_DECK, SAMPLE_MODEL, "Sample files: a fictional deal with planted errors", None)

show_results()
