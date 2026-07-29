"""
One-time interactive LinkedIn login for the posts scraper.
------------------------------------------------------------
Opens a real, visible Chromium window using a persistent profile folder.
Log in to LinkedIn yourself in that window (password typed by you only —
this script never sees or stores your password). Once you land on your
feed, the session cookie is saved into the profile folder, and
linkedin_posts_scraper.py can reuse it headless from then on.

Usage:
    python linkedin_login.py
"""

from pathlib import Path

from playwright.sync_api import sync_playwright

PROFILE_DIR = Path(__file__).resolve().parent / ".linkedin_profile"


def main() -> None:
    PROFILE_DIR.mkdir(exist_ok=True)
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            str(PROFILE_DIR), headless=False
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.goto("https://www.linkedin.com/login")
        print("A browser window has opened. Log in to LinkedIn there.")
        print("Waiting for you to finish logging in (up to 10 minutes)...")
        print("(Leave the window open until this script says 'Login detected'.)")

        deadline = 10 * 60
        waited = 0
        detected = False
        while waited < deadline:
            page.wait_for_timeout(3000)
            waited += 3
            url = page.url
            if "login" not in url and "checkpoint" not in url and "linkedin.com" in url:
                detected = True
                break

        if detected:
            print("Login detected at", page.url)
            print("Session saved to", PROFILE_DIR)
        else:
            print(
                "Didn't detect a logged-in page in time (last URL: "
                f"{page.url}). If you did finish logging in, the session "
                "may still be saved — check by running the posts scraper. "
                "Otherwise, rerun this script."
            )
        context.close()


if __name__ == "__main__":
    main()
