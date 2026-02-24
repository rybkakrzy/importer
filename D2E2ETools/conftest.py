"""
Globalne fixture'y pytest / Playwright.

Konfiguruje przeglądarkę, kontekst i stronę na podstawie ustawień z config/.
"""

import pytest
from playwright.sync_api import Browser, BrowserContext, Page, Playwright

from config.settings import (
    BASE_URL,
    BROWSER,
    DEFAULT_TIMEOUT,
    HEADLESS,
    NAVIGATION_TIMEOUT,
    SLOW_MO,
)


# ──────────────────────────────────────────────
# Fixtures: przeglądarka
# ──────────────────────────────────────────────


@pytest.fixture(scope="session")
def browser_instance(playwright: Playwright) -> Browser:
    """Uruchamia przeglądarkę raz na całą sesję testową."""
    launcher = getattr(playwright, BROWSER)
    browser = launcher.launch(headless=HEADLESS, slow_mo=SLOW_MO)
    yield browser
    browser.close()


@pytest.fixture()
def context(browser_instance: Browser) -> BrowserContext:
    """Tworzy izolowany kontekst przeglądarki dla każdego testu."""
    ctx = browser_instance.new_context(
        viewport={"width": 1920, "height": 1080},
        locale="pl-PL",
        timezone_id="Europe/Warsaw",
    )
    ctx.set_default_timeout(DEFAULT_TIMEOUT)
    ctx.set_default_navigation_timeout(NAVIGATION_TIMEOUT)
    yield ctx
    ctx.close()


@pytest.fixture()
def page(context: BrowserContext) -> Page:
    """Tworzy nową stronę (kartę) w kontekście."""
    pg = context.new_page()
    yield pg
    pg.close()


# ──────────────────────────────────────────────
# Fixtures: nawigacja
# ──────────────────────────────────────────────


@pytest.fixture()
def app_page(page: Page) -> Page:
    """Otwiera aplikację D2 i czeka na załadowanie edytora."""
    page.goto(BASE_URL)
    page.wait_for_selector(".document-editor-container", state="visible")
    return page


# ──────────────────────────────────────────────
# Hooks: screenshoty przy failures
# ──────────────────────────────────────────────


@pytest.hookimpl(tryfirst=True, hookwrapper=True)
def pytest_runtest_makereport(item, call):  # noqa: ANN001, ANN201
    """Dołącza screenshot do raportu gdy test upadnie."""
    import allure

    outcome = yield
    report = outcome.get_result()

    if report.when == "call" and report.failed:
        pg: Page | None = item.funcargs.get("page") or item.funcargs.get("app_page")
        if pg and not pg.is_closed():
            screenshot = pg.screenshot(full_page=True)
            allure.attach(screenshot, name="screenshot", attachment_type=allure.attachment_type.PNG)
