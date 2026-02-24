"""Centralna konfiguracja — ładuje zmienne z .env i wybiera środowisko."""

import os
from pathlib import Path

from dotenv import load_dotenv

from config.environments import ENVIRONMENTS, Environment, LOCAL

# Załaduj .env z katalogu głównego projektu
_env_path = Path(__file__).resolve().parent.parent / ".env"
load_dotenv(_env_path)


def get_environment() -> Environment:
    """Zwraca środowisko na podstawie zmiennej ``TEST_ENV`` lub .env."""
    env_name = os.getenv("TEST_ENV", "local")
    return ENVIRONMENTS.get(env_name, LOCAL)


# Singleton konfiguracji
settings = get_environment()

# Nadpisy z .env (opcjonalne)
BASE_URL: str = os.getenv("BASE_URL", settings.base_url)
API_BASE_URL: str = os.getenv("API_BASE_URL", settings.api_base_url)
HEADLESS: bool = os.getenv("HEADLESS", str(settings.headless)).lower() == "true"
SLOW_MO: int = int(os.getenv("SLOW_MO", str(settings.slow_mo)))
BROWSER: str = os.getenv("BROWSER", "chromium")
DEFAULT_TIMEOUT: int = int(os.getenv("DEFAULT_TIMEOUT", str(settings.default_timeout)))
NAVIGATION_TIMEOUT: int = int(os.getenv("NAVIGATION_TIMEOUT", str(settings.navigation_timeout)))
