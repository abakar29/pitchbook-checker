"""Drives the Streamlit app headlessly. Run with: python -m pytest -q"""
from streamlit.testing.v1 import AppTest


def fresh_app():
    return AppTest.from_file("app.py", default_timeout=30).run()


def test_check_is_disabled_until_both_files_are_uploaded():
    at = fresh_app()
    assert not at.exception
    assert at.button(key="check").disabled


def test_sample_button_shows_every_planted_issue():
    at = fresh_app().button(key="sample").click().run()
    assert not at.exception
    assert len(at.session_state["result"].cross_issues) == 9
    assert len(at.session_state["result"].model_issues) == 9


def test_vercel_copies_are_in_sync():
    """vercel/sync.sh copies the checker and samples; fail if they have drifted."""
    from pathlib import Path
    root = Path(__file__).parent
    for name in ["pitchbook_checker.py", "samples/project_atlas_pitchbook.pptx", "samples/project_atlas_model.xlsx"]:
        assert (root / name).read_bytes() == (root / "vercel" / name).read_bytes(), "run vercel/sync.sh"
