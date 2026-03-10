"""Test wydajnościowy — Health Check endpoint.

Scenariusz smoke / baseline:
  - GET /api/Health
  - Waliduje status 200 + pole „status" = „healthy"

Uruchomienie:
  locust -f locustfiles/health_locust.py --host http://localhost:5190
"""

from locust import HttpUser, between, task


class HealthUser(HttpUser):
    """Użytkownik odpytujący wyłącznie health check."""

    wait_time = between(0.5, 2)

    @task
    def health_check(self) -> None:
        with self.client.get("/api/Health", catch_response=True) as resp:
            if resp.status_code == 200:
                body = resp.json()
                if body.get("status") != "healthy":
                    resp.failure(f"Unexpected status: {body.get('status')}")
            else:
                resp.failure(f"HTTP {resp.status_code}")
