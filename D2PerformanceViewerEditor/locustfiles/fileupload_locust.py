"""Test wydajnościowy — FileUpload (ZIP) API.

Scenariusz:
  POST /api/FileUpload/upload  — przesłanie pliku ZIP (octet-stream)

Uruchomienie:
  locust -f locustfiles/fileupload_locust.py --host http://localhost:5190
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from locust import HttpUser, between, tag, task

from helpers.payloads import zip_file_bytes


class FileUploadUser(HttpUser):
    """Użytkownik przesyłający pliki ZIP."""

    wait_time = between(2, 5)

    @tag("write", "fileupload")
    @task
    def upload_zip_small(self) -> None:
        """Upload małego pliku ZIP (~3 pliki po 1 KB)."""
        data, filename = zip_file_bytes(num_files=3, file_size=1024)
        headers = {
            "Content-Type": "application/octet-stream",
            "X-File-Name": filename,
        }
        with self.client.post(
            "/api/FileUpload/upload",
            data=data,
            headers=headers,
            catch_response=True,
            name="/api/FileUpload/upload [small]",
        ) as resp:
            if resp.status_code == 200:
                body = resp.json()
                if not body.get("success"):
                    resp.failure(f"success=false: {body.get('message')}")
            else:
                resp.failure(f"HTTP {resp.status_code}")

    @tag("write", "fileupload", "large")
    @task
    def upload_zip_medium(self) -> None:
        """Upload średniego pliku ZIP (~10 plików po 10 KB)."""
        data, filename = zip_file_bytes(num_files=10, file_size=10_240)
        headers = {
            "Content-Type": "application/octet-stream",
            "X-File-Name": filename,
        }
        with self.client.post(
            "/api/FileUpload/upload",
            data=data,
            headers=headers,
            catch_response=True,
            name="/api/FileUpload/upload [medium]",
        ) as resp:
            if resp.status_code == 200:
                body = resp.json()
                if not body.get("success"):
                    resp.failure(f"success=false: {body.get('message')}")
            else:
                resp.failure(f"HTTP {resp.status_code}")
