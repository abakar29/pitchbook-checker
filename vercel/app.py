"""Vercel service. GET / serves the upload page; POST (deck, model, units, sample as multipart) returns JSON."""
from __future__ import annotations

import json
import logging
import sys
import tempfile
from email import policy
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from pitchbook_checker import CARD_CSS, check, findings_html, summary_html  # noqa: E402

SAMPLES = ROOT / "samples"
PAGE = ROOT / "public" / "index.html"
UNITS = {"detect": None, "ones": 1.0, "thousands": 1e3, "millions": 1e6, "billions": 1e9}
MAX_BYTES = 4_400_000  # Vercel rejects request bodies over about 4.5 MB
READ_ERROR = ("These files could not be read. Make sure they are valid .pptx and .xlsx files "
              "and are not password protected.")


def parse_form(body: bytes, content_type: str) -> dict:
    """Return {field name: (filename, bytes)} from a multipart/form-data body."""
    message = BytesParser(policy=policy.HTTP).parsebytes(
        b"Content-Type: " + content_type.encode() + b"\r\n\r\n" + body)
    fields = {}
    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if name:
            fields[name] = (part.get_filename(), part.get_payload(decode=True) or b"")
    return fields


def run(fields: dict) -> dict:
    units = fields.get("units", (None, b"detect"))[1].decode(errors="ignore")
    scale = UNITS.get(units)
    if fields.get("sample", (None, b""))[1] == b"1":
        deck, model = SAMPLES / "project_atlas_pitchbook.pptx", SAMPLES / "project_atlas_model.xlsx"
        label = "Sample files: a fictional deal with planted errors"
        return result_json(deck, model, scale, label)
    if "deck" not in fields or "model" not in fields:
        raise ValueError("Upload both a pitchbook and a model.")
    deck_name, deck_bytes = fields["deck"]
    model_name, model_bytes = fields["model"]
    suffix = Path(model_name or "model.xlsx").suffix.lower()
    if suffix not in (".xlsx", ".xlsm") or not (deck_name or "").lower().endswith(".pptx"):
        raise ValueError("The pitchbook must be a .pptx and the model an .xlsx or .xlsm file.")
    with tempfile.TemporaryDirectory() as folder:
        deck, model = Path(folder) / "deck.pptx", Path(folder) / f"model{suffix}"
        deck.write_bytes(deck_bytes)
        model.write_bytes(model_bytes)
        return result_json(deck, model, scale, f"{deck_name} checked against {model_name}")


def result_json(deck: Path, model: Path, scale, label: str) -> dict:
    try:
        result = check(deck, model, scale)
    except Exception:
        logging.exception("Check failed")
        raise ValueError(READ_ERROR)
    return {
        "label": label,
        "css": CARD_CSS,
        "summary": summary_html(result.findings, result.matched),
        "findings": findings_html(result.findings) if result.findings else "",
        "warnings": list(result.warnings),
    }


class handler(BaseHTTPRequestHandler):
    def _send(self, status: int, payload: dict) -> None:
        data = json.dumps(payload).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self) -> None:
        if self.path.split("?")[0] not in ("/", "/index.html"):
            return self._send(404, {"error": "Not found."})
        data = PAGE.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self) -> None:
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_BYTES:
            return self._send(413, {"error": "Files are too large. The limit is 4 MB combined."})
        try:
            fields = parse_form(self.rfile.read(length), self.headers.get("Content-Type", ""))
            self._send(200, run(fields))
        except ValueError as error:
            self._send(400, {"error": str(error)})
        except Exception:
            logging.exception("Request failed")
            self._send(500, {"error": "Something went wrong. Please try again."})
