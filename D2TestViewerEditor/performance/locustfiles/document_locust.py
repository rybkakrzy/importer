"""Test wydajnościowy — Document API.

Scenariusze:
  1. GET  /api/Document/new              — nowy pusty dokument
  2. POST /api/Document/save             — zapis HTML → DOCX
  3. POST /api/Document/open             — otwarcie DOCX → HTML
  4. GET  /api/Document/templates        — lista szablonów
  5. POST /api/Document/upload-image     — upload obrazu → Base64

Uruchomienie:
  locust -f locustfiles/document_locust.py --host http://localhost:5190
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from locust import HttpUser, between, tag, task

from helpers.payloads import docx_file_tuple, document_save_payload, png_file_tuple


class DocumentUser(HttpUser):
    """Użytkownik korzystający z endpointów Document."""

    wait_time = between(1, 4)

    # ── GET new document ─────────────────────────
    @tag("read", "document")
    @task(3)
    def new_document(self) -> None:
        with self.client.get("/api/Document/new", catch_response=True) as resp:
            if resp.status_code != 200:
                resp.failure(f"HTTP {resp.status_code}")

    # ── POST save (HTML → DOCX) ─────────────────
    @tag("write", "document")
    @task(5)
    def save_document(self) -> None:
        payload = document_save_payload()
        with self.client.post(
            "/api/Document/save", json=payload, catch_response=True,
        ) as resp:
            if resp.status_code == 200:
                ct = resp.headers.get("Content-Type", "")
                if "officedocument" not in ct and "octet-stream" not in ct:
                    resp.failure(f"Unexpected Content-Type: {ct}")
            else:
                resp.failure(f"HTTP {resp.status_code}")

    # ── POST open (DOCX → HTML) ─────────────────
    @tag("write", "document")
    @task(3)
    def open_document(self) -> None:
        fname, fbytes, fmime = docx_file_tuple()
        with self.client.post(
            "/api/Document/open",
            files={"file": (fname, fbytes, fmime)},
            catch_response=True,
        ) as resp:
            if resp.status_code == 200:
                body = resp.json()
                if not body.get("html"):
                    resp.failure("Brak 'html' w odpowiedzi")
            else:
                resp.failure(f"HTTP {resp.status_code}")

    # ── GET templates ────────────────────────────
    @tag("read", "document")
    @task(2)
    def get_templates(self) -> None:
        with self.client.get("/api/Document/templates", catch_response=True) as resp:
            if resp.status_code != 200:
                resp.failure(f"HTTP {resp.status_code}")

    # ── POST upload-image ────────────────────────
    @tag("write", "document")
    @task(2)
    def upload_image(self) -> None:
        fname, fbytes, fmime = png_file_tuple()
        with self.client.post(
            "/api/Document/upload-image",
            files={"file": (fname, fbytes, fmime)},
            catch_response=True,
        ) as resp:
            if resp.status_code != 200:
                resp.failure(f"HTTP {resp.status_code}")
