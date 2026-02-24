"""
Helpery do oczekiwania na warunki w testach UI.
"""

from __future__ import annotations

from playwright.sync_api import Page, expect


def wait_for_network_idle(page: Page, timeout: int = 10_000) -> None:
    """Czeka aż sieć będzie bezczynna (0 połączeń) przez chwilę."""
    page.wait_for_load_state("networkidle", timeout=timeout)


def wait_for_element_text(
    page: Page,
    selector: str,
    expected_text: str,
    timeout: int = 5_000,
) -> None:
    """Czeka aż element o podanym selektorze zawiera oczekiwany tekst."""
    locator = page.locator(selector)
    expect(locator).to_contain_text(expected_text, timeout=timeout)


def wait_for_toast(page: Page, text: str | None = None, timeout: int = 5_000) -> None:
    """Czeka na pojawienie się powiadomienia toast."""
    toast = page.locator(".toast, .notification, [role='alert']").first
    toast.wait_for(state="visible", timeout=timeout)
    if text:
        expect(toast).to_contain_text(text)
