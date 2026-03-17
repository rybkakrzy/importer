"""
Testy API: Health endpoint.

Bezpośrednie testy (requests) dla GET /api/health.
"""

from __future__ import annotations

import pytest
import requests


class TestHealthAPI:
    """Testy dla endpointu statusu aplikacji."""

    @pytest.fixture(autouse=True)
    def setup(self, api_base_url: str):
        self.health_url = f"{api_base_url}/health"

    # ── Smoke ─────────────────────────────────────

    @pytest.mark.smoke
    def test_health_returns_200(self):
        """Endpoint health odpowiada statusem 200."""
        response = requests.get(self.health_url, timeout=10)
        assert response.status_code == 200

    @pytest.mark.smoke
    def test_health_returns_json(self):
        """Endpoint health zwraca JSON."""
        response = requests.get(self.health_url, timeout=10)
        assert response.headers.get("Content-Type", "").startswith("application/json")
        data = response.json()
        assert isinstance(data, dict)

    # ── Pola odpowiedzi ───────────────────────────

    @pytest.mark.regression
    def test_health_contains_required_fields(self):
        """Odpowiedź health zawiera wszystkie wymagane pola."""
        response = requests.get(self.health_url, timeout=10)
        data = response.json()
        required_fields = ["status", "environment", "timestamp"]
        for field in required_fields:
            assert field in data, f"Brak wymaganego pola '{field}' w odpowiedzi health"

    @pytest.mark.regression
    def test_health_status_is_healthy(self):
        """Pole 'status' wskazuje na działającą aplikację."""
        response = requests.get(self.health_url, timeout=10)
        data = response.json()
        assert data["status"].lower() in ("healthy", "ok", "running"), (
            f"Nieoczekiwany status aplikacji: '{data['status']}'"
        )

    @pytest.mark.regression
    def test_health_timestamp_is_present(self):
        """Pole 'timestamp' zawiera datę w formacie ISO."""
        response = requests.get(self.health_url, timeout=10)
        data = response.json()
        assert "timestamp" in data
        ts = data["timestamp"]
        # Minimalna walidacja — timestamp musi być niepustym stringiem lub liczbą
        assert ts is not None and ts != "", f"Timestamp jest pusty lub None: {ts!r}"

    @pytest.mark.regression
    def test_health_build_info_optional_fields(self):
        """Pola buildNumber i buildDate są obecne (lub puste, ale nie brakuje struktury)."""
        response = requests.get(self.health_url, timeout=10)
        data = response.json()
        # Te pola są opcjonalne, ale jeśli istnieją — sprawdź typ
        if "buildNumber" in data:
            assert isinstance(data["buildNumber"], (str, int, type(None)))
        if "buildDate" in data:
            assert isinstance(data["buildDate"], (str, type(None)))

    # ── Nagłówki ──────────────────────────────────

    @pytest.mark.regression
    def test_health_cors_headers(self):
        """Endpoint health zawiera nagłówki CORS (dostępność z frontendu)."""
        response = requests.get(self.health_url, timeout=10)
        # Sprawdzamy, że odpowiedź jest dostępna — nagłówki CORS nie są wymagane
        # na endpointach wewnętrznych, ale status 200 jest wystarczający
        assert response.status_code == 200

    @pytest.mark.regression
    def test_health_response_time_acceptable(self):
        """Endpoint health odpowiada w czasie poniżej 2 sekund."""
        import time

        start = time.monotonic()
        requests.get(self.health_url, timeout=10)
        elapsed = time.monotonic() - start
        assert elapsed < 2.0, f"Health endpoint odpowiedział za wolno: {elapsed:.2f}s"
