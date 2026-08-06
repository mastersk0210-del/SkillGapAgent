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
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.goto(URL, wait_until="networkidle", timeout=60_000)

        # Free-tier apps sleep after inactivity; click through the wake prompt if present.
        try:
            wake_button = page.get_by_text(WAKE_BUTTON_TEXT, exact=False)
            if wake_button.count() > 0:
                wake_button.first.click()
                page.wait_for_timeout(30_000)
        except Exception:
            pass

        # Give the app time to finish booting.
        page.wait_for_timeout(15_000)

        body_text = page.inner_text("body").lower()
        page.screenshot(path=SCREENSHOT_PATH, full_page=True)
        browser.close()

    for sign in ERROR_SIGNS:
        if sign in body_text:
            print(f"App appears to be in an error state — matched: {sign!r}")
            return 1

    print("App looks healthy — no known error signatures found.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
