"""Test wydajnościowy — WSZYSTKIE endpointy łącznie.

Plik łączy scenariusze z poszczególnych locustfiles
i definiuje wagi proporcjonalne do przewidywanego ruchu produkcyjnego.

Uruchomienie:
  locust -f locustfiles/all_endpoints_locust.py --host http://localhost:5190
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from locust import HttpUser, between, tag, task

from helpers.payloads import (
    barcode_generate_payload,
    docx_file_tuple,
    document_save_payload,
    png_file_tuple,
    zip_file_bytes,
)


class MixedTrafficUser(HttpUser):
    """Symuluje typowego użytkownika aplikacji D2Tools.

    Proporcje tasków odzwierciedlają realistyczny mix ruchu:
      - Health check         — sporadycznie (monitoring)
      - Barcode (read/write) — umiarkowanie
      - Document (CRUD)      — najczęściej
      - FileUpload           — rzadziej
    """

    wait_time = between(1, 4)

    # ────────────────────────────────────────────
    # Health
    # ────────────────────────────────────────────
    @tag("health")
    @task(1)
    def health_check(self) -> None:
        self.client.get("/api/Health")

    # ────────────────────────────────────────────
    # Barcode
    # ────────────────────────────────────────────
    @tag("barcode", "read")
    @task(2)
    def barcode_types(self) -> None:
        self.client.get("/api/Barcode/types")

    @tag("barcode", "write")
    @task(4)
    def barcode_generate(self) -> None:
        self.client.post("/api/Barcode/generate", json=barcode_generate_payload())

    @tag("barcode", "write")
    @task(2)
    def barcode_generate_image(self) -> None:
        self.client.post("/api/Barcode/generate-image", json=barcode_generate_payload())

    # ────────────────────────────────────────────
    # Document
    # ────────────────────────────────────────────
    @tag("document", "read")
    @task(3)
    def document_new(self) -> None:
        self.client.get("/api/Document/new")

    @tag("document", "write")
    @task(5)
    def document_save(self) -> None:
        self.client.post("/api/Document/save", json=document_save_payload())

    @tag("document", "write")
    @task(3)
    def document_open(self) -> None:
        fname, fbytes, fmime = docx_file_tuple()
        self.client.post(
            "/api/Document/open",
            files={"file": (fname, fbytes, fmime)},
        )

    @tag("document", "read")
    @task(2)
    def document_templates(self) -> None:
        self.client.get("/api/Document/templates")

    @tag("document", "write")
    @task(2)
    def document_upload_image(self) -> None:
        fname, fbytes, fmime = png_file_tuple()
        self.client.post(
            "/api/Document/upload-image",
            files={"file": (fname, fbytes, fmime)},
        )

    # ────────────────────────────────────────────
    # FileUpload
    # ────────────────────────────────────────────
    @tag("fileupload", "write")
    @task(2)
    def upload_zip(self) -> None:
        data, filename = zip_file_bytes(num_files=3, file_size=1024)
        self.client.post(
            "/api/FileUpload/upload",
            data=data,
            headers={
                "Content-Type": "application/octet-stream",
                "X-File-Name": filename,
            },
            name="/api/FileUpload/upload",
        )
