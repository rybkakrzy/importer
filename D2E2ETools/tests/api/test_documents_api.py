"""
Testy BDD: Documents & Barcode API.

Scenariusze z features/api/documents_api.feature
"""

from __future__ import annotations

import pytest
from pytest_bdd import given, parsers, scenarios, then, when

from tests.api.conftest import ApiClient

scenarios("../../features/api/documents_api.feature")


# ──────────────────────────────────────────────
# Kontekst współdzielony między krokami
# ──────────────────────────────────────────────


@pytest.fixture()
def api_ctx() -> dict:
    """Mały kontener na dane przekazywane między krokami."""
    return {}


# ──────────────────────────────────────────────
# WHEN — GET
# ──────────────────────────────────────────────


@when(parsers.parse('wysyłam GET na "{path}"'))
def send_get(api: ApiClient, path: str, api_ctx: dict):
    resp = api.get(path)
    api_ctx["response"] = resp


# ──────────────────────────────────────────────
# WHEN — POST z tabelą danych
# ──────────────────────────────────────────────


@when(parsers.parse('wysyłam POST na "{path}" z danymi:'))
def send_post_with_table(api: ApiClient, path: str, api_ctx: dict, datatable):
    """
    Konwertuje tabelę z feature-a (nagłówki + wiersz) na dict JSON.
    pytest-bdd przekazuje datatable jako listę wierszy.
    """
    # datatable to obiekt pytest-bdd; interpretujemy jako dict z jednego wiersza
    if hasattr(datatable, "rows"):
        headers = datatable.rows[0]
        values = datatable.rows[1]
        payload = dict(zip(headers, values))
    else:
        # Fallback — datatable jako lista list
        headers = datatable[0]
        values = datatable[1]
        payload = dict(zip(headers, values))

    # Automatyczna konwersja typów
    for key, val in payload.items():
        if val.isdigit():
            payload[key] = int(val)

    resp = api.post(path, json=payload)
    api_ctx["response"] = resp


# ──────────────────────────────────────────────
# THEN — status, body
# ──────────────────────────────────────────────


@then(parsers.parse("status odpowiedzi to {code:d}"))
def response_status(api_ctx: dict, code: int):
    resp = api_ctx["response"]
    assert resp.status_code == code, (
        f"Oczekiwany status {code}, otrzymano {resp.status_code}.\n"
        f"Body: {resp.text[:500]}"
    )


@then(parsers.parse('odpowiedź zawiera pole "{field}"'))
def response_has_field(api_ctx: dict, field: str):
    data = api_ctx["response"].json()
    assert field in data, f"Brak pola '{field}' w odpowiedzi: {list(data.keys())}"


@then("odpowiedź jest listą")
def response_is_list(api_ctx: dict):
    data = api_ctx["response"].json()
    assert isinstance(data, list), f"Oczekiwano listy, otrzymano {type(data).__name__}"


@then(parsers.parse('odpowiedź ma content-type "{content_type}"'))
def response_content_type(api_ctx: dict, content_type: str):
    actual_ct = api_ctx["response"].headers.get("Content-Type", "")
    assert content_type in actual_ct, (
        f"Oczekiwany content-type zawierający '{content_type}', "
        f"otrzymano: '{actual_ct}'"
    )
