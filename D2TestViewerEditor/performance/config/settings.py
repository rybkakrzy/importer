"""Centralna konfiguracja testów wydajnościowych."""

import os
from pathlib import Path

from dotenv import load_dotenv

# Załaduj .env z katalogu głównego D2PerformenceTools
_env_path = Path(__file__).resolve().parent.parent / ".env"
load_dotenv(_env_path)

# ── URL-e ─────────────────────────────────────────
API_BASE_URL: str = os.getenv("API_BASE_URL", "http://localhost:5190/api")

# ── Parametry Locust (defaults – można nadpisać przez CLI) ──
LOCUST_USERS: int = int(os.getenv("LOCUST_USERS", "10"))
LOCUST_SPAWN_RATE: int = int(os.getenv("LOCUST_SPAWN_RATE", "2"))
LOCUST_RUN_TIME: str = os.getenv("LOCUST_RUN_TIME", "60s")
