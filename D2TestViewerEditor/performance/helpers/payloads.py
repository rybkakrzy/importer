"""Generatory danych testowych (payloads) dla poszczególnych endpointów."""

import io
import random
import struct
import zipfile
import zlib
from typing import Any

from faker import Faker

fake = Faker("pl_PL")

# ── Barcode ───────────────────────────────────────

BARCODE_TYPES = [
    "QRCode",
    "Code128",
    "Code39",
    "EAN13",
    "EAN8",
    "UPC_A",
    "DataMatrix",
]


def barcode_generate_payload() -> dict[str, Any]:
    """Losowy payload dla POST /api/Barcode/generate."""
    barcode_type = random.choice(BARCODE_TYPES)
    content = (
        fake.ean13() if barcode_type in ("EAN13",) else
        fake.ean8() if barcode_type in ("EAN8",) else
        fake.bothify("??##??##")
    )
    return {
        "content": content,
        "barcodeType": barcode_type,
        "width": random.choice([200, 300, 400]),
        "height": random.choice([200, 300, 400]),
        "showText": random.choice([True, False]),
    }


# ── Document – save ──────────────────────────────

def document_save_payload() -> dict[str, Any]:
    """Payload JSON dla POST /api/Document/save."""
    paragraphs = "\n".join(f"<p>{fake.paragraph(nb_sentences=3)}</p>" for _ in range(5))
    return {
        "html": f"<h1>{fake.sentence()}</h1>{paragraphs}",
        "originalFileName": f"{fake.file_name(extension='docx')}",
        "metadata": {
            "title": fake.sentence(nb_words=4),
            "author": fake.name(),
        },
    }


# ── Document – open (multipart DOCX upload) ──────

def minimal_docx_bytes() -> bytes:
    """Zwraca minimalny poprawny plik DOCX."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            "</Types>",
        )
        zf.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="word/document.xml"/>'
            "</Relationships>",
        )
        zf.writestr(
            "word/_rels/document.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            "</Relationships>",
        )
        zf.writestr(
            "word/document.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            "<w:body><w:p><w:r><w:t>Performance test document</w:t></w:r></w:p></w:body>"
            "</w:document>",
        )
    buf.seek(0)
    return buf.read()


# Cache — generujemy DOCX raz, reużywamy w każdym uzyciu
_CACHED_DOCX: bytes | None = None


def docx_file_tuple() -> tuple[str, bytes, str]:
    """Zwraca (filename, bytes, content_type) do multipart uploadu DOCX."""
    global _CACHED_DOCX
    if _CACHED_DOCX is None:
        _CACHED_DOCX = minimal_docx_bytes()
    return (
        f"perf-test-{fake.uuid4()[:8]}.docx",
        _CACHED_DOCX,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ── Document – upload-image (multipart PNG upload) ──

def _minimal_png(width: int = 4, height: int = 4) -> bytes:
    """Generuje minimalny prawidłowy obraz PNG."""

    def _chunk(chunk_type: bytes, data: bytes) -> bytes:
        c = chunk_type + data
        crc = struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)
        return struct.pack(">I", len(data)) + c + crc

    header = b"\x89PNG\r\n\x1a\n"
    ihdr_data = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    ihdr = _chunk(b"IHDR", ihdr_data)
    # niebieski piksel — raw scanlines
    raw_rows = b""
    for _ in range(height):
        raw_rows += b"\x00"  # filter byte
        raw_rows += b"\x00\x00\xff" * width  # blue RGB
    idat = _chunk(b"IDAT", zlib.compress(raw_rows))
    iend = _chunk(b"IEND", b"")
    return header + ihdr + idat + iend


_CACHED_PNG: bytes | None = None


def png_file_tuple() -> tuple[str, bytes, str]:
    """Zwraca (filename, bytes, content_type) do multipart uploadu obrazu."""
    global _CACHED_PNG
    if _CACHED_PNG is None:
        _CACHED_PNG = _minimal_png()
    return (
        f"perf-test-{fake.uuid4()[:8]}.png",
        _CACHED_PNG,
        "image/png",
    )


# ── FileUpload – ZIP ─────────────────────────────

def zip_file_bytes(num_files: int = 3, file_size: int = 1024) -> tuple[bytes, str]:
    """Tworzy minimalny plik ZIP w pamięci. Zwraca (bytes, filename)."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i in range(num_files):
            zf.writestr(f"file_{i}.txt", fake.text(max_nb_chars=file_size))
    buf.seek(0)
    filename = f"test-upload-{fake.uuid4()[:8]}.zip"
    return buf.read(), filename
