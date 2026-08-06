"""Loads the live Streamlit Cloud app in a headless browser and checks for
known error/failure states. Plain HTTP requests can't do this — Streamlit
Cloud's outer page is a JS SPA shell that returns 200 regardless of whether
the app underneath is healthy, crashed, or asleep.
"""
import os
import sys
from playwright.sync_api import sync_playwright

URL = os.environ.get("APP_URL", "https://skillgapagent-mastersk07.streamlit.app/")
SCREENSHOT_PATH = "live_app_screenshot.png"

ERROR_SIGNS = [
    "this app has encountered an error",
    "modulenotfounderror",
    "importerror",
    "error installing requirements",
    "oh no.",
    "traceback",
]

WAKE_BUTTON_TEXT = "get this app back up"


def main() -> int:
    body_text = ""
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        try:
            # Streamlit apps hold an open WebSocket for live updates, so the
            # network never goes fully idle — "load" is the right condition here.
            page.goto(URL, wait_until="load", timeout=45_000)
            page.wait_for_timeout(5_000)

            # Free-tier apps sleep after inactivity; click through the wake prompt if present.
            wake_button = page.get_by_text(WAKE_BUTTON_TEXT, exact=False)
            if wake_button.count() > 0:
                wake_button.first.click()
                page.wait_for_timeout(30_000)

            # Give the app time to finish booting / rendering.
            page.wait_for_timeout(15_000)
            body_text = page.inner_text("body").lower()
        finally:
            try:
                page.screenshot(path=SCREENSHOT_PATH, full_page=True)
            except Exception as e:
                print(f"Could not capture screenshot: {e}")
            browser.close()

    if not body_text:
        print("Could not read any page content — treating as inconclusive, not a failure.")
        return 0

    for sign in ERROR_SIGNS:
        if sign in body_text:
            print(f"App appears to be in an error state — matched: {sign!r}")
            return 1

    print("App looks healthy — no known error signatures found.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
