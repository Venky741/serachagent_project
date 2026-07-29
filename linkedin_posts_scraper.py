"""
LinkedIn "Posts" search scraper
--------------------------------
Separate from the Jobs-board scraper: this targets recruiters' organic
feed posts (the "Posts" tab you get when searching LinkedIn directly),
which the Jobs API never returns.

Requires a logged-in session created once via linkedin_login.py. Reuses
that session headless on every subsequent run.

LinkedIn renders search results with obfuscated, auto-generated CSS class
names that change on every deploy, so this locates each result via the
accessibility tree instead (role="listitem" containers, one per post) —
a more stable hook than styling classes, but still not a public API, so
a layout change degrades to "found nothing" rather than crashing the run.
Scraping LinkedIn like this is against their Terms of Service; this is
for personal, low-volume, rate-limited use on your own account, not for
scaling up or redistributing.

There's no direct permalink exposed in the search-results view, so the
"link" captured per post is the author's profile URL (the person to
contact / the source of the post), not a link to the post itself.
"""

import re
from pathlib import Path
from urllib.parse import quote

from playwright.sync_api import sync_playwright

PROFILE_DIR = Path(__file__).resolve().parent / ".linkedin_profile"

TRAILING_JUNK = re.compile(r"^\s*\d+\s*$")
# LinkedIn shows each post's age as e.g. "4h", "20m", "1d", "2w" right after
# the author's headline. There's no hour-level date filter in the search
# URL itself, so freshness is enforced here instead, after scraping.
AGE_PATTERN = re.compile(r"^(\d+)\s*(m|min|mins|h|hr|hrs|d|w|mo)\b", re.IGNORECASE)
_AGE_MULTIPLIER = {"m": 1, "min": 1, "mins": 1, "h": 60, "hr": 60, "hrs": 60,
                    "d": 1440, "w": 10080, "mo": 43200}


def _age_minutes(lines: list[str]) -> int | None:
    for ln in lines:
        m = AGE_PATTERN.match(ln.replace("•", "").strip())
        if m:
            return int(m.group(1)) * _AGE_MULTIPLIER[m.group(2).lower()]
    return None


def _search_url(term: str) -> str:
    return (
        "https://www.linkedin.com/search/results/content/"
        f"?keywords={quote(term)}&datePosted=%22past-24h%22&sortBy=%22date_posted%22"
    )


def _expand_truncated_posts(page) -> None:
    for _ in range(15):
        buttons = page.get_by_text("… more", exact=False).all()
        if not buttons:
            break
        clicked_any = False
        for b in buttons:
            try:
                b.click(timeout=1000)
                clicked_any = True
            except Exception:
                continue
        if not clicked_any:
            break
        page.wait_for_timeout(400)


def _parse_card_text(text: str) -> dict | None:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip() and ln.strip() != "​"]
    if lines and lines[0] == "Feed post":
        lines = lines[1:]
    if len(lines) < 3:
        return None

    author = lines[0]
    body_start = None
    for i, ln in enumerate(lines):
        if ln in ("Follow", "Join", "Following", "Pending"):
            body_start = i + 1
            break
    if body_start is None:
        body_start = min(3, len(lines))

    age_minutes = _age_minutes(lines[:body_start])

    body_lines = lines[body_start:]
    while body_lines and TRAILING_JUNK.match(body_lines[-1]):
        body_lines.pop()
    if body_lines and body_lines[-1] == "… more":
        body_lines.pop()

    description = "\n".join(body_lines).strip()
    if len(description) < 40:
        return None

    return {"author": author, "description": description, "age_minutes": age_minutes}


def scrape_posts(search_terms: list[str], max_scrolls: int = 6, max_age_minutes: int = 60) -> list[dict]:
    if not PROFILE_DIR.exists():
        print("  (skipping posts scrape: no LinkedIn session - run linkedin_login.py first)")
        return []

    results = []
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(str(PROFILE_DIR), headless=True)
        page = context.pages[0] if context.pages else context.new_page()

        for term in search_terms:
            try:
                page.goto(_search_url(term), wait_until="domcontentloaded", timeout=30000)
                page.wait_for_timeout(2500)

                if "/login" in page.url or "checkpoint" in page.url:
                    print(f"  !! LinkedIn session expired/blocked for {term!r} - rerun linkedin_login.py")
                    break

                for _ in range(max_scrolls):
                    page.mouse.wheel(0, 2500)
                    page.wait_for_timeout(1000)

                _expand_truncated_posts(page)

                cards = page.locator('[role="listitem"]').all()
                term_count = 0
                skipped_stale = 0
                for card in cards:
                    try:
                        text = card.inner_text(timeout=2000)
                    except Exception:
                        continue
                    if "Feed post" not in text:
                        continue

                    parsed = _parse_card_text(text)
                    if not parsed:
                        continue
                    if parsed["age_minutes"] is None or parsed["age_minutes"] > max_age_minutes:
                        skipped_stale += 1
                        continue

                    link = ""
                    try:
                        link = card.locator('a[href*="/in/"]').first.get_attribute("href", timeout=1000) or ""
                        link = link.split("?")[0]
                    except Exception:
                        pass

                    results.append({
                        "search_term": term,
                        "title": parsed["description"].splitlines()[0][:120],
                        "company": parsed["author"],
                        "company_url": link,
                        "location": "",
                        "date_posted": "",
                        "job_type": "contract",  # unknown from posts; treated as a contract candidate
                        "site": "linkedin_posts",
                        "job_url": link,
                        "description": parsed["description"],
                    })
                    term_count += 1
                print(f"  posts search {term!r} -> {term_count} candidate posts "
                      f"(skipped {skipped_stale} older than {max_age_minutes}m)")
            except Exception as exc:
                print(f"  !! posts scrape failed for {term!r}: {exc}")

        context.close()

    # de-dupe by description text (same post can appear across scrolls/terms)
    seen = set()
    deduped = []
    for r in results:
        key = r["description"][:200]
        if key in seen:
            continue
        seen.add(key)
        deduped.append(r)
    return deduped


if __name__ == "__main__":
    found = scrape_posts(["Senior Java Developer", "Java Developer"])
    print(f"Found {len(found)} posts")
    for f in found[:5]:
        print("-", f["title"], "->", f["job_url"])
