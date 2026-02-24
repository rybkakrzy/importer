"""Konfiguracja środowisk testowych."""

from dataclasses import dataclass, field


@dataclass(frozen=True)
class Environment:
    """Definicja środowiska testowego."""

    name: str
    base_url: str
    api_base_url: str
    headless: bool = True
    slow_mo: int = 0
    default_timeout: int = 30_000
    navigation_timeout: int = 30_000


# ──────────────────────────────────────────────
# Predefiniowane środowiska
# ──────────────────────────────────────────────

LOCAL = Environment(
    name="local",
    base_url="http://localhost:4200",
    api_base_url="http://localhost:5190/api",
    headless=False,
    slow_mo=50,
)

DEV = Environment(
    name="dev",
    base_url="http://localhost:4200",
    api_base_url="http://localhost:5190/api",
    headless=True,
)

CI = Environment(
    name="ci",
    base_url="http://localhost:4200",
    api_base_url="http://localhost:5190/api",
    headless=True,
    default_timeout=60_000,
    navigation_timeout=60_000,
)

ENVIRONMENTS: dict[str, Environment] = {
    "local": LOCAL,
    "dev": DEV,
    "ci": CI,
}
