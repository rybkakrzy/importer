"""
Generatory danych testowych i stałe.
"""

from __future__ import annotations

import random
import string

# ── Stałe ──────────────────────────────────────

SAMPLE_HTML_CONTENT = "<p>Przykładowy dokument testowy</p>"
SAMPLE_BARCODE_CONTENT = "TEST-12345"
SAMPLE_QR_URL = "https://example.com/test"
SAMPLE_EAN13 = "5901234123457"

VALID_BARCODE_TYPES = ["QRCode", "Code128", "EAN13", "EAN8", "Code39"]

# ── Generatory ─────────────────────────────────


def random_text(length: int = 20) -> str:
    """Generuje losowy tekst o podanej długości."""
    return "".join(random.choices(string.ascii_letters + " ", k=length))


def random_filename(extension: str = "docx") -> str:
    """Generuje losową nazwę pliku."""
    name = "".join(random.choices(string.ascii_lowercase, k=8))
    return f"{name}.{extension}"


def sample_document_payload(html: str | None = None) -> dict:
    """Zwraca przykładowy payload do POST /document/save."""
    return {
        "html": html or SAMPLE_HTML_CONTENT,
        "originalFileName": random_filename(),
    }


def sample_barcode_payload(
    content: str | None = None,
    barcode_type: str = "QRCode",
    width: int = 300,
    height: int = 300,
) -> dict:
    """Zwraca przykładowy payload do POST /barcode/generate."""
    return {
        "content": content or SAMPLE_BARCODE_CONTENT,
        "barcodeType": barcode_type,
        "width": width,
        "height": height,
    }
