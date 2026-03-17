"""
Generatory danych testowych i stałe.
"""

from __future__ import annotations

import base64
import random
import string

# ── Stałe — kody kreskowe ──────────────────────

SAMPLE_HTML_CONTENT = "<p>Przykładowy dokument testowy</p>"
SAMPLE_BARCODE_CONTENT = "TEST-12345"
SAMPLE_QR_URL = "https://example.com/test"
SAMPLE_EAN13 = "5901234123457"
SAMPLE_EAN8 = "12345678"
SAMPLE_CODE39 = "TEST-CODE"

# Wszystkie 14 typów obsługiwanych przez BarcodeService
VALID_BARCODE_TYPES = [
    "QRCode",
    "Code128",
    "EAN13",
    "UPC-A",
    "EAN8",
    "I2of5",
    "Code39",
    "Aztec",
    "DataMatrix",
    "PDF417",
    "MicroPDF",
    "Codabar",
    "Code93",
    "Maxicode",
]

# Typy obsługiwane przez dialog UI (podzbiór)
UI_BARCODE_TYPES = ["QRCode", "Code128", "EAN13", "EAN8", "Code39"]

# ── Stałe — podpisy cyfrowe ────────────────────

SAMPLE_SIGNER_NAME = "Jan Kowalski"
SAMPLE_SIGNER_TITLE = "Dyrektor"
SAMPLE_SIGNER_EMAIL = "jan.kowalski@example.com"
SAMPLE_SIGNATURE_REASON = "Zatwierdzenie dokumentu"

# ── Stałe — document storage ──────────────────

SAMPLE_TEXT_CONTENT_B64 = base64.b64encode(b"Przykladowy dokument testowy").decode()
SAMPLE_MIME_TYPES = [
    "text/plain",
    "text/markdown",
    "application/pdf",
    "application/octet-stream",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
]

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


def sample_upload_payload(
    name: str | None = None,
    mime_type: str = "text/plain",
    content_bytes: bytes | None = None,
    created_by: str = "TestUser",
) -> dict:
    """Zwraca przykładowy payload do POST /documentstorage/upload."""
    raw = content_bytes or b"Przykladowy dokument testowy"
    return {
        "name": name or random_filename("txt"),
        "mimeType": mime_type,
        "content": base64.b64encode(raw).decode(),
        "createdBy": created_by,
    }


def sample_save_version_payload(
    content_text: str = "Zaktualizowana treść dokumentu",
    created_by: str = "TestEditor",
) -> dict:
    """Zwraca przykładowy payload do POST /documentstorage/{id}/save."""
    return {
        "content": base64.b64encode(content_text.encode()).decode(),
        "createdBy": created_by,
    }


def sample_sign_payload(
    html: str | None = None,
    cert_base64: str = "",
    cert_password: str = "",
    signer_name: str | None = None,
) -> dict:
    """Zwraca minimalny payload do POST /document/sign."""
    return {
        "html": html or SAMPLE_HTML_CONTENT,
        "originalFileName": random_filename(),
        "certificateBase64": cert_base64,
        "certificatePassword": cert_password,
        "signerName": signer_name or SAMPLE_SIGNER_NAME,
        "signerTitle": SAMPLE_SIGNER_TITLE,
        "signerEmail": SAMPLE_SIGNER_EMAIL,
        "signatureReason": SAMPLE_SIGNATURE_REASON,
    }

