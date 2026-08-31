"""
Java Contract Job Search Agent
------------------------------
Searches LinkedIn (and optionally other boards) for Java Developer /
Senior Java Developer contract roles in the United States, then extracts:

  - Job details (title, location, posted date, description, URL)
  - Vendor details (company name, company LinkedIn/profile URL)
  - Recruiter contact info (emails / phone numbers found in the posting)
  - A C2C flag (postings that mention Corp-to-Corp / C2C / 1099)

Only contract postings that mention C2C are kept in the final output.

Output:
  - A timestamped CSV you can open in Excel
  - job_viewer_data.js (used by job_viewer.html to auto-load the latest run)
  - An email to GMAIL_TO with a summary + the CSV attached

Usage:
    pip install python-jobspy pandas python-dotenv
    Create a .env file (see .env.example) with GMAIL_USER / GMAIL_APP_PASSWORD / GMAIL_TO
    python java_contract_job_agent.py
"""

import ctypes
import json
import os
import re
import smtplib
import sys
import traceback
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from jobspy import scrape_jobs

from linkedin_posts_scraper import scrape_posts

load_dotenv()

# When output is redirected to a file (as the scheduled task does), Python
# fully buffers stdout instead of flushing per line. If the process is then
# killed (e.g. the machine goes to sleep mid-run), whatever was still in the
# buffer is lost — which is why some failed runs left no trace in the log at
# all. Reopen stdout/stderr unbuffered so every print lands on disk immediately.
sys.stdout.reconfigure(line_buffering=True)
sys.stderr.reconfigure(line_buffering=True)

# ------------------------- KEEP-AWAKE (Windows) ------------------------------
# A running task doesn't by itself stop Windows from sleeping. If the machine
# suspends mid-run, the process gets killed outright (surfaces as Windows
# status STATUS_CONTROL_C_EXIT) instead of pausing and resuming. Tell Windows
# to stay awake for the duration of this script, then release that hold when
# done so normal sleep behavior resumes afterward.
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001


def prevent_sleep() -> None:
    if sys.platform == "win32":
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)


def allow_sleep() -> None:
    if sys.platform == "win32":
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

SCRIPT_DIR = Path(__file__).resolve().parent
VIEWER_DATA_FILE = SCRIPT_DIR / "job_viewer_data.js"

GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
GMAIL_TO = os.environ.get("GMAIL_TO", GMAIL_USER)

# ----------------------------- CONFIG ---------------------------------------

SEARCH_TERMS = [
    "Senior Java Developer",
    "Java Developer",
    "Java Full Stack Developer",
    "Java Backend Developer",
]

# Boards to search. LinkedIn is the main one you asked for.
# You can add: "indeed", "zip_recruiter", "glassdoor"
SITES = ["linkedin"]

# Also scrape LinkedIn's organic "Posts" search (recruiters posting in their
# feed), which the Jobs board API never surfaces. Requires a one-time login
# via linkedin_login.py.
ENABLE_POSTS_SCRAPE = True

LOCATION = "United States"
JOB_TYPE = "contract"        # contract roles only
HOURS_OLD = 1                # only postings from the last hour
RESULTS_PER_TERM = 50        # per search term, per site

POSTS_MAX_AGE_MINUTES = 60

C2C_PATTERN = re.compile(
    r"\b(c2c|corp[\s\-]?to[\s\-]?corp|1099|w2\s*/\s*c2c|c2c\s*/\s*w2)\b",
    re.IGNORECASE,
)
# Postings that explicitly reject C2C (or say W2-only) should be excluded even
# though they contain the word "c2c" and would otherwise trip C2C_PATTERN.
NO_C2C_PATTERN = re.compile(
    r"\b(no\s*c2c|not\s*open\s*to\s*c2c|"
    r"unable\s*to\s*(?:do|sponsor)\s*c2c|cannot\s*do\s*c2c|"
    r"w-?2\s*only|only\s*w-?2|"
    r"w-?2\s*(?:candidates?|contractors?|basis|employment)\s*only|"
    r"only\s*w-?2\s*(?:candidates?|contractors?)|"
    r"w-?2\s*salaried\s*employee|"
    r"c2c\s*not\s*(?:accepted|available|allowed))\b",
    re.IGNORECASE,
)
# Recruiters marketing available bench consultants, not an actual open role.
HOTLIST_PATTERN = re.compile(
    r"\b(hot\s*list|hotlist|bench\s*(?:consultants?|sales)|"
    r"consultants?\s*(?:available|on\s*bench|ready)|"
    r"available\s*(?:on\s*)?c2c\s*consultants?|"
    r"c2c\s*(?:it\s*)?consultants?\s*ready)\b",
    re.IGNORECASE,
)
# A candidate marketing themselves ("I'm open to work"), not a recruiter
# posting an opening at all.
CANDIDATE_SELF_POST_PATTERN = re.compile(
    r"\b(open\s*to\s*work|#opentowork|not_a_hiring_post|"
    r"i'?m\s*(?:currently\s*)?(?:open\s*to|exploring|looking\s*for)\s*(?:new\s*)?"
    r"(?:c2c\s*)?opportunit|"
    r"i\s*am\s*(?:currently\s*)?(?:open\s*to|exploring|looking\s*for))\b",
    re.IGNORECASE,
)
# Posts search matches "java" anywhere in a post (even a stray hashtag), so a
# lot of non-Java roles (PEGA, .NET, mobile/Android-iOS, etc.) slip through.
# Require Java to actually be a subject of the role, not just mentioned in
# passing. \bjava\b won't match "javascript" (no word boundary after "java").
JAVA_ROLE_PATTERN = re.compile(r"\bjava\b(?!\s*script)", re.IGNORECASE)

EMAIL_PATTERN = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
PHONE_PATTERN = re.compile(
    r"(?:\+?1[\s.-]?)?\(?\d{3}\)?[\s.-]\d{3}[\s.-]\d{4}"
)

# ----------------------------- HELPERS --------------------------------------


def extract_contacts(description: str) -> tuple[str, str]:
    """Pull recruiter emails / phone numbers out of the job description."""
    if not isinstance(description, str):
        return "", ""
    emails = sorted(set(EMAIL_PATTERN.findall(description)))
    phones = sorted(set(p.strip() for p in PHONE_PATTERN.findall(description)))
    return "; ".join(emails), "; ".join(phones)


def _snippet(text: str, match: "re.Match") -> str:
    start = max(0, match.start() - 40)
    end = min(len(text), match.end() + 40)
    return re.sub(r"\s+", " ", text[start:end]).strip()


def is_c2c(description: str, title: str) -> bool:
    text = f"{title or ''} {description or ''}"
    return bool(C2C_PATTERN.search(text))


def c2c_evidence(description: str, title: str) -> str:
    text = f"{title or ''} {description or ''}"
    match = C2C_PATTERN.search(text)
    return _snippet(text, match) if match else ""


def rejects_c2c(description: str, title: str) -> bool:
    text = f"{title or ''} {description or ''}"
    return bool(NO_C2C_PATTERN.search(text))


def rejects_c2c_evidence(description: str, title: str) -> str:
    text = f"{title or ''} {description or ''}"
    match = NO_C2C_PATTERN.search(text)
    return _snippet(text, match) if match else ""


def is_hotlist(description: str, title: str) -> bool:
    text = f"{title or ''} {description or ''}"
    return bool(HOTLIST_PATTERN.search(text))


def is_candidate_self_post(description: str, title: str) -> bool:
    text = f"{title or ''} {description or ''}"
    return bool(CANDIDATE_SELF_POST_PATTERN.search(text))


def is_java_role(description: str, title: str) -> bool:
    # Require Java in the title, or early in the description (role/skills
    # section) rather than anywhere at all — keeps out posts where "java"
    # only shows up buried in a trailing hashtag list.
    if JAVA_ROLE_PATTERN.search(title or ""):
        return True
    return bool(JAVA_ROLE_PATTERN.search((description or "")[:300]))


def write_viewer_data(jobs: pd.DataFrame, outfile: str) -> None:
    """Dump the latest run as job_viewer_data.js so job_viewer.html can auto-load it."""
    records = jobs.where(pd.notnull(jobs), "").to_dict(orient="records")
    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_file": outfile,
        "jobs": records,
    }
    VIEWER_DATA_FILE.write_text(
        "window.__JOB_VIEWER_DATA__ = " + json.dumps(payload, default=str) + ";\n",
        encoding="utf-8",
    )


def _html_escape(text: str) -> str:
    return (
        (text or "")
        .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def send_email_notification(jobs: pd.DataFrame, outfile: str) -> None:
    if not (GMAIL_USER and GMAIL_APP_PASSWORD and GMAIL_TO):
        print("  (skipping email: GMAIL_USER / GMAIL_APP_PASSWORD / GMAIL_TO not set in .env)")
        return

    count = len(jobs)
    with_email = int((jobs["vendor_emails"] != "").sum())
    shown = jobs.head(40)

    plain_lines = []
    html_rows = []
    for r in shown.itertuples():
        source = "Job board" if r.site != "linkedin_posts" else "Recruiter post"
        raw_link = r.job_url or ""
        link = raw_link if raw_link.startswith(("http://", "https://")) else ""
        plain_lines.append(
            f"{r.title} — {r.company} ({r.location or 'n/a'}) [{source}]\n"
            f"  Contact: {r.vendor_emails or 'n/a'}\n"
            f"  Link: {link or 'n/a'}"
        )
        link_cell = (
            f'<a href="{_html_escape(link)}" style="color:#3b5bdb;">Open</a>'
            if link else "&mdash;"
        )
        html_rows.append(f"""
          <tr>
            <td style="padding:8px 10px;border-bottom:1px solid #e0e0e5;">
              <div style="font-weight:600;">{_html_escape(r.title)}</div>
              <div style="color:#6b6b76;font-size:12px;">{_html_escape(source)}</div>
            </td>
            <td style="padding:8px 10px;border-bottom:1px solid #e0e0e5;">{_html_escape(r.company)}</td>
            <td style="padding:8px 10px;border-bottom:1px solid #e0e0e5;">{_html_escape(r.location or '')}</td>
            <td style="padding:8px 10px;border-bottom:1px solid #e0e0e5;">{_html_escape(r.vendor_emails or '')}</td>
            <td style="padding:8px 10px;border-bottom:1px solid #e0e0e5;">{link_cell}</td>
          </tr>""")

    html = f"""\
    <div style="font-family:Arial,Helvetica,sans-serif;color:#1c1c1f;">
      <p><strong>{count}</strong> contract/C2C Java job(s) found this run
         &middot; <strong>{with_email}</strong> include a vendor email.</p>
      <table style="border-collapse:collapse;width:100%;font-size:13px;">
        <thead>
          <tr style="background:#f0f1f7;text-align:left;">
            <th style="padding:8px 10px;">Title / Source</th>
            <th style="padding:8px 10px;">Company / Posted by</th>
            <th style="padding:8px 10px;">Location</th>
            <th style="padding:8px 10px;">Contact</th>
            <th style="padding:8px 10px;">Link</th>
          </tr>
        </thead>
        <tbody>{''.join(html_rows)}</tbody>
      </table>
      <p style="color:#6b6b76;font-size:12px;">Full list saved locally to {_html_escape(outfile)}.</p>
    </div>
    """

    msg = EmailMessage()
    msg["Subject"] = f"Java Contract C2C Jobs — {count} found ({datetime.now():%Y-%m-%d %H:%M})"
    msg["From"] = GMAIL_USER
    msg["To"] = GMAIL_TO
    msg.set_content(
        f"{count} contract/C2C Java job(s) found this run.\n"
        f"{with_email} include a vendor email in the posting.\n\n" +
        "\n\n".join(plain_lines) +
        f"\n\nFull list saved locally to {outfile}."
    )
    msg.add_alternative(html, subtype="html")

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        smtp.send_message(msg)
    print(f"  Emailed summary to {GMAIL_TO}")


# ----------------------------- MAIN -----------------------------------------


def _scrape_job_board() -> list[pd.DataFrame]:
    frames = []
    for term in SEARCH_TERMS:
        print(f"Searching {', '.join(SITES)} for: {term!r} ...")
        try:
            df = scrape_jobs(
                site_name=SITES,
                search_term=term,
                location=LOCATION,
                job_type=JOB_TYPE,
                hours_old=HOURS_OLD,
                results_wanted=RESULTS_PER_TERM,
                country_indeed="USA",
                linkedin_fetch_description=True,  # full description + company URL
                verbose=1,
            )
            df["search_term"] = term
            frames.append(df)
            print(f"  -> {len(df)} results")
        except Exception as exc:  # rate limits, network hiccups, etc.
            print(f"  !! search failed for {term!r}: {exc}")
    return frames


def _scrape_posts_source() -> list[pd.DataFrame]:
    if not ENABLE_POSTS_SCRAPE:
        return []
    print("Searching linkedin_posts for recruiter feed posts ...")
    posts = scrape_posts(SEARCH_TERMS, max_age_minutes=POSTS_MAX_AGE_MINUTES)
    if posts:
        print(f"  -> {len(posts)} candidate posts")
        return [pd.DataFrame(posts)]
    return []


def run() -> pd.DataFrame:
    start = datetime.now()
    print(f"=== RUN START {start:%Y-%m-%d %H:%M:%S} (pid={os.getpid()}) ===")
    prevent_sleep()
    try:
        result = _run()
        print(f"=== RUN END {datetime.now():%Y-%m-%d %H:%M:%S} "
              f"(status=ok, duration={datetime.now() - start}) ===")
        return result
    except Exception:
        print(f"=== RUN END {datetime.now():%Y-%m-%d %H:%M:%S} "
              f"(status=FAILED, duration={datetime.now() - start}) ===")
        traceback.print_exc()
        raise
    finally:
        allow_sleep()


def _run() -> pd.DataFrame:
    # The job board (plain HTTP requests via jobspy) and the posts scraper
    # (a full browser scrolling/expanding through search results) don't share
    # any state, so run them concurrently instead of back-to-back — the posts
    # scraper is the long pole either way, so this cuts total wall time from
    # their sum down to roughly whichever one is slower.
    with ThreadPoolExecutor(max_workers=2) as executor:
        job_board_future = executor.submit(_scrape_job_board)
        posts_future = executor.submit(_scrape_posts_source)
        frames = job_board_future.result() + posts_future.result()

    if not frames:
        print("No results retrieved. Try again later or reduce request volume.")
        return pd.DataFrame()

    jobs = pd.concat(frames, ignore_index=True)

    # De-duplicate (same job often matches both search terms). Posts without
    # an extractable author link have no job_url, so dedup on a separate key
    # instead of overwriting job_url with a fake, non-clickable placeholder.
    missing_url = jobs["job_url"].isna() | (jobs["job_url"] == "")
    dedup_key = jobs["job_url"].where(~missing_url, "post:" + jobs["description"].str[:200])
    jobs = jobs[~dedup_key.duplicated(keep="first")]

    # Vendor contact extraction + C2C flag
    contacts = jobs["description"].apply(extract_contacts)
    jobs["vendor_emails"] = contacts.apply(lambda t: t[0])
    jobs["vendor_phones"] = contacts.apply(lambda t: t[1])
    jobs["c2c_mentioned"] = [
        is_c2c(d, t) for d, t in zip(jobs["description"], jobs["title"])
    ]
    jobs["c2c_evidence"] = [
        c2c_evidence(d, t) for d, t in zip(jobs["description"], jobs["title"])
    ]
    jobs["rejects_c2c"] = [
        rejects_c2c(d, t) for d, t in zip(jobs["description"], jobs["title"])
    ]
    jobs["rejects_c2c_evidence"] = [
        rejects_c2c_evidence(d, t) for d, t in zip(jobs["description"], jobs["title"])
    ]
    jobs["is_hotlist"] = [
        is_hotlist(d, t) for d, t in zip(jobs["description"], jobs["title"])
    ]
    jobs["is_java_role"] = [
        is_java_role(d, t) for d, t in zip(jobs["description"], jobs["title"])
    ]
    jobs["is_candidate_self_post"] = [
        is_candidate_self_post(d, t) for d, t in zip(jobs["description"], jobs["title"])
    ]

    # Posts (unlike Jobs-board listings) have no "Apply" button, so a genuine
    # recruiter post almost always includes a way to reach them; lacking any
    # contact info is a strong signal it's commentary/opinion, not a posting.
    has_contact = (jobs["vendor_emails"].fillna("") != "") | (jobs["vendor_phones"].fillna("") != "")
    is_post = jobs["site"] == "linkedin_posts"

    # Only Java contract postings that mention C2C, don't explicitly reject
    # it, aren't a recruiter marketing bench consultants, and aren't a
    # candidate advertising themselves rather than an open role
    jobs = jobs[
        (jobs["job_type"] == "contract")
        & (jobs["c2c_mentioned"])
        & (~jobs["rejects_c2c"])
        & (~jobs["is_hotlist"])
        & (~jobs["is_candidate_self_post"])
        & (jobs["is_java_role"])
        & (~is_post | has_contact)
    ]

    if jobs.empty:
        print("No contract/C2C postings matched this run.")
        return jobs

    # Columns jobspy only populates for other boards (Indeed/ZipRecruiter/
    # Glassdoor), never for LinkedIn — drop them to keep the CSV readable.
    ALWAYS_EMPTY_FOR_LINKEDIN = [
        "job_url_direct", "salary_source", "interval", "min_amount",
        "max_amount", "currency", "listing_type", "company_url_direct",
        "company_addresses", "company_num_employees", "company_revenue",
        "company_description", "skills", "experience_range",
        "company_rating", "company_reviews_count", "vacancy_count",
        "work_from_home_type",
    ]
    jobs = jobs.drop(
        columns=[c for c in ALWAYS_EMPTY_FOR_LINKEDIN if c in jobs.columns]
    )
    # These are only informative when they're True/non-empty; once filtered,
    # every surviving row has rejects_c2c == False, so drop the noise.
    jobs = jobs.drop(columns=[
        "rejects_c2c", "rejects_c2c_evidence", "is_hotlist",
        "is_java_role", "is_candidate_self_post",
    ])

    # Tidy column order: vendor + job details first, description last
    preferred = [
        "search_term", "title", "company", "company_url",
        "vendor_emails", "vendor_phones", "c2c_mentioned", "c2c_evidence",
        "location", "date_posted", "job_type", "site", "job_url",
    ]
    cols = (
        [c for c in preferred if c in jobs.columns]
        + [c for c in jobs.columns if c not in preferred and c != "description"]
        + ["description"]
    )
    jobs = jobs[cols].sort_values("date_posted", ascending=False)

    # Save
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    outfile = f"java_contract_jobs_{stamp}.csv"
    jobs.to_csv(outfile, index=False)

    print(f"\nSaved {len(jobs)} contract/C2C jobs -> {outfile}")
    print(f"  Postings with a vendor email: "
          f"{int((jobs['vendor_emails'] != '').sum())}")

    write_viewer_data(jobs, outfile)
    send_email_notification(jobs, outfile)

    return jobs


if __name__ == "__main__":
    try:
        run()
    except Exception:
        sys.exit(1)