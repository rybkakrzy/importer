"""
Fixtures specyficzne dla testów API.

Dostarcza skonfigurowanego klienta HTTP (requests.Session)
z bazowym URL-em API.
"""

import pytest
import requests

from config.settings import get_settings


class ApiClient:
    """Lekki wrapper wokół requests.Session ułatwiający testy REST API."""

    def __init__(self, base_url: str, timeout: int = 30):
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Accept": "application/json",
                "Content-Type": "application/json",
            }
        )
        self.last_response: requests.Response | None = None

    # ── HTTP verbs ────────────────────────────

    def get(self, path: str, **kwargs) -> requests.Response:
        url = f"{self.base_url}{path}"
        self.last_response = self.session.get(url, timeout=self.timeout, **kwargs)
        return self.last_response

    def post(self, path: str, json: dict | None = None, **kwargs) -> requests.Response:
        url = f"{self.base_url}{path}"
        self.last_response = self.session.post(url, json=json, timeout=self.timeout, **kwargs)
        return self.last_response

    def put(self, path: str, json: dict | None = None, **kwargs) -> requests.Response:
        url = f"{self.base_url}{path}"
        self.last_response = self.session.put(url, json=json, timeout=self.timeout, **kwargs)
        return self.last_response

    def delete(self, path: str, **kwargs) -> requests.Response:
        url = f"{self.base_url}{path}"
        self.last_response = self.session.delete(url, timeout=self.timeout, **kwargs)
        return self.last_response

    def post_file(self, path: str, file_path: str, field_name: str = "file") -> requests.Response:
        """Wysyła plik przez multipart/form-data."""
        url = f"{self.base_url}{path}"
        with open(file_path, "rb") as f:
            files = {field_name: f}
            # Usuwamy Content-Type, żeby requests sam ustawił multipart boundary
            headers = {k: v for k, v in self.session.headers.items() if k.lower() != "content-type"}
            self.last_response = self.session.post(
                url, files=files, headers=headers, timeout=self.timeout
            )
        return self.last_response

    # ── helpers ────────────────────────────────

    def close(self):
        self.session.close()


@pytest.fixture(scope="session")
def api_client() -> ApiClient:
    """Sesyjny klient API — jeden na cały przebieg testów."""
    settings = get_settings()
    client = ApiClient(base_url=settings.api_base_url)
    yield client
    client.close()


@pytest.fixture()
def api(api_client: ApiClient) -> ApiClient:
    """Per-testowy alias z wyczyszczonym last_response."""
    api_client.last_response = None
    return api_client
