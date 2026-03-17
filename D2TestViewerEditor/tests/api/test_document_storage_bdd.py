"""
Testy BDD: Document Storage API — wersjonowanie dokumentów.

Scenariusze z features/api/document_storage.feature
"""

from __future__ import annotations

import base64

import pytest
from pytest_bdd import given, parsers, scenarios, then, when

from tests.api.conftest import ApiClient

scenarios("../../features/api/document_storage.feature")

_SAMPLE_CONTENT_B64 = base64.b64encode(b"Tresc testowego dokumentu").decode()


# ──────────────────────────────────────────────
# GIVEN — przygotowanie danych (reużywa api_ctx z conftest)
# ──────────────────────────────────────────────


@given(
    parsers.parse('przesłano dokument o nazwie "{name}" i typie "{mime}"'),
    target_fixture="api_ctx",
)
def given_upload_document(api: ApiClient, name: str, mime: str, api_ctx: dict) -> dict:
    payload = {
        "name": name,
        "mimeType": mime,
        "content": _SAMPLE_CONTENT_B64,
        "createdBy": "BDDTest",
    }
    resp = api.post("/documentstorage/upload", json=payload)
    assert resp.status_code == 200, f"Upload nie powiódł się: {resp.text}"
    data = resp.json()
    api_ctx["masterId"] = data["masterId"]
    api_ctx["versionId"] = data["versionId"]
    api_ctx["response"] = resp
    return api_ctx


@given(
    parsers.parse("zapisano {count:d} dodatkowe wersje dokumentu"),
    target_fixture="api_ctx",
)
def given_save_extra_versions(api: ApiClient, count: int, api_ctx: dict) -> dict:
    master_id = api_ctx["masterId"]
    for i in range(count):
        content = base64.b64encode(f"Wersja dodatkowa {i + 2}".encode()).decode()
        resp = api.post(
            f"/documentstorage/{master_id}/save",
            json={"content": content, "createdBy": "BDDTest"},
        )
        assert resp.status_code == 200
    return api_ctx


# ──────────────────────────────────────────────
# WHEN
# ──────────────────────────────────────────────


@when(parsers.parse('wysyłam uploadDocument z nazwą "{name}" i typem "{mime}"'))
def when_upload_document(api: ApiClient, name: str, mime: str, api_ctx: dict):
    payload = {
        "name": name,
        "mimeType": mime,
        "content": _SAMPLE_CONTENT_B64,
        "createdBy": "BDDTest",
    }
    resp = api.post("/documentstorage/upload", json=payload)
    api_ctx["response"] = resp
    if resp.status_code == 200:
        data = resp.json()
        api_ctx["masterId"] = data.get("masterId")
        api_ctx["versionId"] = data.get("versionId")


@when("wysyłam uploadDocument z pustą zawartością")
def when_upload_empty_content(api: ApiClient, api_ctx: dict):
    payload = {
        "name": "empty.txt",
        "mimeType": "text/plain",
        "content": "",
        "createdBy": "BDDTest",
    }
    resp = api.post("/documentstorage/upload", json=payload)
    api_ctx["response"] = resp


@when("pobieram dokument po masterId")
def when_get_document(api: ApiClient, api_ctx: dict):
    master_id = api_ctx["masterId"]
    resp = api.get(f"/documentstorage/{master_id}")
    api_ctx["response"] = resp


@when("zapisuję nową wersję dokumentu")
def when_save_version(api: ApiClient, api_ctx: dict):
    master_id = api_ctx["masterId"]
    new_content = base64.b64encode(b"Zaktualizowana tresc wersji 2").decode()
    resp = api.post(
        f"/documentstorage/{master_id}/save",
        json={"content": new_content, "createdBy": "BDDEditor"},
    )
    api_ctx["response"] = resp


@when("pobieram listę wersji dokumentu")
def when_get_versions(api: ApiClient, api_ctx: dict):
    master_id = api_ctx["masterId"]
    resp = api.get(f"/documentstorage/{master_id}/versions")
    api_ctx["response"] = resp


@when("przywracam pierwszą wersję dokumentu")
def when_restore_first_version(api: ApiClient, api_ctx: dict):
    master_id = api_ctx["masterId"]
    version_id = api_ctx["versionId"]  # ID wersji 1 (z uploadu)
    resp = api.post(f"/documentstorage/{master_id}/restore/{version_id}")
    api_ctx["response"] = resp


# ──────────────────────────────────────────────
# THEN — asercje specyficzne dla storage
# ──────────────────────────────────────────────


@then(parsers.parse("wersja dokumentu wynosi {num:d}"))
def then_version_is(api_ctx: dict, num: int):
    data = api_ctx["response"].json()
    assert data.get("versionNumber") == num, (
        f"Oczekiwano versionNumber={num}, otrzymano {data.get('versionNumber')}"
    )


@then(parsers.parse("numer wersji wynosi {num:d}"))
def then_version_number_is(api_ctx: dict, num: int):
    data = api_ctx["response"].json()
    assert data.get("versionNumber") == num, (
        f"Oczekiwano versionNumber={num}, otrzymano {data.get('versionNumber')}"
    )


@then("wersje są posortowane malejąco")
def then_versions_sorted(api_ctx: dict):
    versions = api_ctx["response"].json()
    nums = [v["versionNumber"] for v in versions]
    assert nums == sorted(nums, reverse=True), (
        f"Wersje nie są posortowane malejąco: {nums}"
    )


@then("tylko najnowsza wersja jest aktywna")
def then_only_latest_is_active(api_ctx: dict):
    versions = api_ctx["response"].json()
    active_versions = [v for v in versions if v.get("isActive") is True]
    assert len(active_versions) == 1, (
        f"Oczekiwano dokładnie 1 aktywnej wersji, znaleziono {len(active_versions)}"
    )
    assert active_versions[0]["versionNumber"] == versions[0]["versionNumber"], (
        "Aktywna wersja nie jest najnowszą wersją"
    )

