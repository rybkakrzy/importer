"""Test wydajnościowy — Barcode API.

Scenariusze:
  1. GET  /api/Barcode/types             — pobranie listy obsługiwanych typów
  2. POST /api/Barcode/generate          — generowanie kodu (base64)
  3. POST /api/Barcode/generate-image    — generowanie kodu (PNG)

Uruchomienie:
  locust -f locustfiles/barcode_locust.py --host http://localhost:5190
"""

import sys
from pathlib import Path

# Upewniamy się, że root projektu jest w sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from locust import HttpUser, between, tag, task

from helpers.payloads import barcode_generate_payload


class BarcodeUser(HttpUser):
    """Użytkownik korzystający z endpointów Barcode."""

    wait_time = between(1, 3)

    # ── GET supported types ──────────────────────
    @tag("read", "barcode")
    @task(2)
    def get_supported_types(self) -> None:
        with self.client.get("/api/Barcode/types", catch_response=True) as resp:
            if resp.status_code != 200:
                resp.failure(f"HTTP {resp.status_code}")

    # ── POST generate (base64) ───────────────────
    @tag("write", "barcode")
    @task(5)
    def generate_barcode_base64(self) -> None:
        payload = barcode_generate_payload()
        with self.client.post(
            "/api/Barcode/generate", json=payload, catch_response=True,
        ) as resp:
            if resp.status_code == 200:
                body = resp.json()
                if not body.get("base64Image"):
                    resp.failure("Brak base64Image w odpowiedzi")
            else:
                resp.failure(f"HTTP {resp.status_code}")

    # ── POST generate-image (PNG bytes) ──────────
    @tag("write", "barcode")
    @task(3)
    def generate_barcode_image(self) -> None:
        payload = barcode_generate_payload()
        with self.client.post(
            "/api/Barcode/generate-image", json=payload, catch_response=True,
        ) as resp:
            if resp.status_code == 200:
                content_type = resp.headers.get("Content-Type", "")
                if "image/png" not in content_type:
                    resp.failure(f"Oczekiwano image/png, otrzymano {content_type}")
            else:
                resp.failure(f"HTTP {resp.status_code}")
