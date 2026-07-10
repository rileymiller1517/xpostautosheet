import json
import os
import random
import re
import sys
import tempfile
import time
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials as SACredentials
from google.oauth2.credentials import Credentials as UserCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from playwright.sync_api import sync_playwright

# ── Top-level config: only things that don't change post-to-post live here ───
GDRIVE_CREDS_JSON   = os.environ.get("GDRIVE_CREDENTIALS_JSON", "")
SHEET_CREDS_JSON    = os.environ.get("SHEET_CREDENTIALS_JSON", "") or GDRIVE_CREDS_JSON

ACCOUNTS_SPREADSHEET_ID = os.environ.get("ACCOUNTS_SPREADSHEET_ID", "")
ACCOUNTS_SHEET_TAB      = os.environ.get("ACCOUNTS_SHEET_TAB", "Accounts")
MASTER_SHEET_TAB        = os.environ.get("MASTER_SHEET_TAB", "Settings")

REPO    = os.environ.get("GITHUB_REPOSITORY", "unknown-repo")
RUN_ID  = os.environ.get("GITHUB_RUN_ID", "unknown-run")
ACTOR   = os.environ.get("GITHUB_ACTOR", "")
LOCK_IDENTIFIER = f"{REPO}#{RUN_ID}" + (f" ({ACTOR})" if ACTOR else "")

RUN_BUDGET_MINUTES  = float(os.environ.get("RUN_BUDGET_MINUTES", "355"))
MAX_CONSECUTIVE_SESSION_FAILURES = 2

IMAGE_MIMES = {"image/jpeg", "image/png", "image/gif", "image/webp"}
VIDEO_MIMES = {"video/mp4", "video/quicktime", "video/x-msvideo", "video/mpeg"}

SCREENSHOT_DIR = Path("debug_screenshots")
SCREENSHOT_DIR.mkdir(exist_ok=True)
step_counter = [0]

SESSION_ERROR_KEYWORDS = ("login", "session", "restriction", "graduated-access")

HASHTAG_RE = re.compile(r"#\w+")

# Data and Threads are single shared tabs, in the SAME spreadsheet as Accounts —
# every account's rows live together and are told apart by an ACCOUNT_NAME column.
# Tab names can still be renamed globally via the Settings tab if you want.
DATA_SHEET_TAB_DEFAULT = "Data"
THREAD_SHEET_TAB_DEFAULT = "Threads"

# Defaults used only if a setting is blank in BOTH the account row and the
# master Settings tab.
CONFIG_DEFAULTS = {
    "DATA_SHEET_TAB": DATA_SHEET_TAB_DEFAULT,
    "THREAD_SHEET_TAB": THREAD_SHEET_TAB_DEFAULT,
    "SOURCE_FOLDER_ID": "",
    "CLAIMED_FOLDER_ID": "",
    "POST_COUNT": "10",
    "IMAGE_PERCENT": "40",
    "VIDEO_PERCENT": "30",
    "THREAD_PERCENT": "30",
    "INTERVAL_MINUTES": "10",
    "SHUFFLE": "true",
    "THREAD_DELIMITER": "---",
    "LINK_URL": "",
    "LINK_ENABLE_IMAGE": "false",
    "LINK_PERCENT_IMAGE": "0",
    "LINK_ENABLE_VIDEO": "false",
    "LINK_PERCENT_VIDEO": "0",
}


def dbg(msg):
    ts = time.strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def screenshot(page, label):
    step_counter[0] += 1
    name = f"{step_counter[0]:03d}_{label}.png"
    path = SCREENSHOT_DIR / name
    try:
        page.screenshot(path=str(path), full_page=False)
        dbg(f"  📸 Screenshot saved: {path}")
    except Exception as e:
        dbg(f"  📸 Screenshot failed ({label}): {e}")


# ── Credential helper (shared by Drive + Sheets) ──────────────────────────────

def build_creds(creds_json_str, scopes):
    if not creds_json_str.strip():
        sys.exit("Required credentials secret is empty (check GDRIVE_CREDENTIALS_JSON / SHEET_CREDENTIALS_JSON in repo Settings > Secrets).")
    data = json.loads(creds_json_str)
    if data.get("type") == "service_account":
        return SACredentials.from_service_account_info(data, scopes=scopes)
    return UserCredentials(
        token=data.get("token"),
        refresh_token=data.get("refresh_token"),
        token_uri=data.get("token_uri", "https://oauth2.googleapis.com/token"),
        client_id=data.get("client_id"),
        client_secret=data.get("client_secret"),
        scopes=data.get("scopes", scopes),
    )


def build_drive_service():
    dbg("Building Google Drive service…")
    creds = build_creds(GDRIVE_CREDS_JSON, ["https://www.googleapis.com/auth/drive"])
    return build("drive", "v3", credentials=creds)


def build_sheet_client():
    dbg("Authenticating to Google Sheets…")
    creds = build_creds(SHEET_CREDS_JSON, [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ])
    return gspread.authorize(creds)


def get_tab(sh, tab_name, required_for=""):
    try:
        return sh.worksheet(tab_name)
    except gspread.exceptions.WorksheetNotFound:
        available = [w.title for w in sh.worksheets()]
        sys.exit(
            f"Tab '{tab_name}' not found{(' (' + required_for + ')') if required_for else ''}. "
            f"Available tabs in that spreadsheet: {available}."
        )


def update_cell(ws, row_number, header_name, value):
    """Write a value into the column matching `header_name` in row 1, for the given row_number."""
    headers = ws.row_values(1)
    try:
        col = headers.index(header_name) + 1
    except ValueError:
        dbg(f"  ⚠ Column '{header_name}' not found in sheet '{ws.title}' (tab headers: {headers}) — cannot write '{value}'.")
        return
    ws.update_cell(row_number, col, value)


def now_iso():
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


# ── Master Accounts sheet: settings, locking, per-account config merge ───────

def load_master_settings(sh):
    """Settings tab format: column A = SETTING name, column B = VALUE."""
    try:
        ws = sh.worksheet(MASTER_SHEET_TAB)
    except gspread.exceptions.WorksheetNotFound:
        dbg(f"No '{MASTER_SHEET_TAB}' tab found — proceeding with no master defaults (per-account values / hardcoded defaults only).")
        return {}
    rows = ws.get_all_values()
    settings = {}
    for row in rows[1:]:
        if len(row) >= 2 and row[0].strip():
            settings[row[0].strip().upper()] = row[1].strip()
    dbg(f"Loaded {len(settings)} master default(s) from '{MASTER_SHEET_TAB}' tab.")
    return settings


def claim_account(ws, records):
    """
    records: output of ws.get_all_records() — one dict per data row, in sheet order.
    Row 2 of the sheet == records[0].

    Rule: if THIS repo already has a row assigned to it, reuse that row forever.
    Otherwise, claim the first row from the top that has no LOCKED_BY yet.
    """
    for i, rec in enumerate(records):
        row_number = i + 2
        if str(rec.get("ASSIGNED_REPO", "")).strip() == REPO and str(rec.get("LOCKED_BY", "")).strip():
            dbg(f"Row {row_number} ('{rec.get('ACCOUNT_NAME','')}') is already assigned to this repo — reusing it.")
            return row_number, rec

    for i, rec in enumerate(records):
        row_number = i + 2
        if not str(rec.get("LOCKED_BY", "")).strip():
            update_cell(ws, row_number, "LOCKED_BY", LOCK_IDENTIFIER)
            update_cell(ws, row_number, "LOCKED_AT", now_iso())
            update_cell(ws, row_number, "ASSIGNED_REPO", REPO)
            update_cell(ws, row_number, "ASSIGNED_STATUS", "ACTIVE")
            dbg(f"Claimed row {row_number} ('{rec.get('ACCOUNT_NAME','')}') for {REPO}.")
            rec["LOCKED_BY"] = LOCK_IDENTIFIER
            rec["ASSIGNED_REPO"] = REPO
            rec["ASSIGNED_STATUS"] = "ACTIVE"
            return row_number, rec

    sys.exit(
        "No available account row found: every row in the Accounts sheet is already LOCKED_BY someone else. "
        "Add a new account row, or clear LOCKED_BY on a row you want to free up."
    )


def get_cfg(account_rec, master, key):
    val = str(account_rec.get(key, "")).strip()
    if val:
        return val
    val = str(master.get(key, "")).strip()
    if val:
        return val
    return CONFIG_DEFAULTS[key]


def build_effective_config(account_rec, master):
    cfg = {k: get_cfg(account_rec, master, k) for k in CONFIG_DEFAULTS}
    # type coercion
    cfg["POST_COUNT"] = int(float(cfg["POST_COUNT"]))
    cfg["IMAGE_PERCENT"] = float(cfg["IMAGE_PERCENT"])
    cfg["VIDEO_PERCENT"] = float(cfg["VIDEO_PERCENT"])
    cfg["THREAD_PERCENT"] = float(cfg["THREAD_PERCENT"])
    cfg["INTERVAL_MINUTES"] = float(cfg["INTERVAL_MINUTES"])
    cfg["SHUFFLE"] = cfg["SHUFFLE"].strip().lower() == "true"
    cfg["LINK_ENABLE_IMAGE"] = cfg["LINK_ENABLE_IMAGE"].strip().lower() == "true"
    cfg["LINK_ENABLE_VIDEO"] = cfg["LINK_ENABLE_VIDEO"].strip().lower() == "true"
    cfg["LINK_PERCENT_IMAGE"] = float(cfg["LINK_PERCENT_IMAGE"])
    cfg["LINK_PERCENT_VIDEO"] = float(cfg["LINK_PERCENT_VIDEO"])
    return cfg


# ── Google Drive helpers ──────────────────────────────────────────────────────

def list_media_in_folder(service, folder_id, mimes):
    if not folder_id or not mimes:
        return []
    dbg(f"Listing media in Drive folder: {folder_id} (mimes={mimes})")
    query = (
        f"'{folder_id}' in parents and trashed=false and ("
        + " or ".join(f"mimeType='{m}'" for m in mimes)
        + ")"
    )
    results, page_token = [], None
    while True:
        resp = service.files().list(
            q=query, fields="nextPageToken, files(id, name, mimeType)",
            pageSize=200, pageToken=page_token,
        ).execute()
        results.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    dbg(f"  Found {len(results)} file(s).")
    return results


def index_drive_files(files):
    idx = {}
    for f in files:
        idx[f["name"].strip().lower()] = f
        stem = Path(f["name"]).stem.strip().lower()
        idx.setdefault(stem, f)
    return idx


def match_file_for_row(data_row, file_index):
    name = data_row["file_name"].strip().lower()
    if name in file_index:
        return file_index[name]
    return file_index.get(Path(name).stem)


def download_file(service, file_id, file_name, dest_path, max_attempts=4):
    for attempt in range(1, max_attempts + 1):
        dbg(f"  Downloading '{file_name}' -> {dest_path} (attempt {attempt}/{max_attempts})")
        try:
            request = service.files().get_media(fileId=file_id)
            with open(dest_path, "wb") as fh:
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    status, done = downloader.next_chunk()
                    if status:
                        dbg(f"    Download progress: {int(status.progress()*100)}%")
            dbg("  Download complete.")
            return
        except Exception as e:
            dbg(f"  ⚠ Download attempt {attempt} failed: {e}")
            try:
                Path(dest_path).unlink(missing_ok=True)
            except Exception:
                pass
            if attempt < max_attempts:
                wait_s = 5 * attempt
                dbg(f"  Retrying download in {wait_s}s…")
                time.sleep(wait_s)
            else:
                raise


def move_to_claimed(service, claimed_folder_id, file_id, file_name):
    if not claimed_folder_id:
        return
    try:
        meta = service.files().get(fileId=file_id, fields="parents").execute()
        prev = ",".join(meta.get("parents", []))
        service.files().update(
            fileId=file_id, addParents=claimed_folder_id,
            removeParents=prev, fields="id, parents",
        ).execute()
        dbg(f"  ✓ Moved '{file_name}' -> claimed folder.")
    except Exception as e:
        dbg(f"  ⚠ WARNING: Failed to move '{file_name}': {e}")


# ── Content ("data") sheet: FILE_NAME | CAPTION | HASHTAGS | STATUS | POSTED_AT

def load_data_rows(sh, tab, account_name):
    """Data tab is shared by ALL accounts in the same spreadsheet as Accounts.
    Rows are filtered to this account's ACCOUNT_NAME and skip anything already Posted."""
    ws = get_tab(sh, tab, required_for="content/data sheet")
    records = ws.get_all_records()
    rows = []
    for i, rec in enumerate(records):
        row_number = i + 2
        if str(rec.get("ACCOUNT_NAME", "")).strip().lower() != account_name.strip().lower():
            continue
        status = str(rec.get("STATUS", "")).strip().lower()
        if status == "posted":
            continue
        file_name = str(rec.get("FILE_NAME", "")).strip()
        if not file_name:
            continue
        rows.append({
            "row_number": row_number,
            "file_name": file_name,
            "caption": str(rec.get("CAPTION", "")).strip(),
            "hashtags": str(rec.get("HASHTAGS", "")).strip(),
        })
    dbg(f"Loaded {len(rows)} unposted content row(s) for account '{account_name}' from tab '{tab}'.")
    return ws, rows


def plan_data_jobs(data_rows, file_index, want_images, want_videos):
    """Classify every unposted row that has a matching Drive file into an image
    or video job. No cap — ALL matching rows are returned; the run only stops
    when the sheets run dry or the time budget runs out."""
    image_jobs, video_jobs = [], []
    for row in data_rows:
        f = match_file_for_row(row, file_index)
        if not f:
            dbg(f"  ⚠ No matching Drive file for FILE_NAME '{row['file_name']}' (data row {row['row_number']}) — skipping this run.")
            continue
        mime = f.get("mimeType", "")
        if mime in IMAGE_MIMES and want_images:
            image_jobs.append({"kind": "image", "data_row": row, "drive_file": f})
        elif mime in VIDEO_MIMES and want_videos:
            video_jobs.append({"kind": "video", "data_row": row, "drive_file": f})
    return image_jobs, video_jobs


def dedupe_hashtags(caption, hashtags_raw):
    if not hashtags_raw.strip():
        return ""
    existing = {m.group(0).lower() for m in HASHTAG_RE.finditer(caption)}
    tags = re.split(r"[\s,]+", hashtags_raw.strip())
    out, seen = [], set()
    for t in tags:
        t = t.strip()
        if not t:
            continue
        if not t.startswith("#"):
            t = "#" + t
        low = t.lower()
        if low in existing or low in seen:
            continue
        seen.add(low)
        out.append(t)
    return " ".join(out)


def build_text_from_data_row(row, link):
    caption = row["caption"]
    hashtags = dedupe_hashtags(caption, row["hashtags"])
    parts = [caption]
    if hashtags:
        parts.append("")
        parts.append(hashtags)
    text = "\n".join(parts)
    if link:
        text += f"\n\n{link}"
    return text


def maybe_link(kind, cfg):
    if kind == "image":
        enabled, pct = cfg["LINK_ENABLE_IMAGE"], cfg["LINK_PERCENT_IMAGE"]
    else:
        enabled, pct = cfg["LINK_ENABLE_VIDEO"], cfg["LINK_PERCENT_VIDEO"]
    if enabled and cfg["LINK_URL"] and random.uniform(0, 100) < pct:
        return cfg["LINK_URL"]
    return ""


def mark_data_posted(ws, row_number):
    update_cell(ws, row_number, "STATUS", "Posted")
    update_cell(ws, row_number, "POSTED_AT", now_iso())


# ── Thread sheet: TEXT | STATUS | POSTED_AT ───────────────────────────────────

def load_thread_rows(sh, tab, account_name):
    """Threads tab is shared by ALL accounts in the same spreadsheet as Accounts.
    Rows are filtered to this account's ACCOUNT_NAME and skip anything already Posted."""
    ws = get_tab(sh, tab, required_for="thread sheet")
    records = ws.get_all_records()
    rows = []
    for i, rec in enumerate(records):
        row_number = i + 2
        if str(rec.get("ACCOUNT_NAME", "")).strip().lower() != account_name.strip().lower():
            continue
        status = str(rec.get("STATUS", "")).strip().lower()
        if status == "posted":
            continue
        text = str(rec.get("TEXT", "")).strip()
        if text:
            rows.append({"row_number": row_number, "text": text})
    dbg(f"Loaded {len(rows)} unposted thread row(s) for account '{account_name}' from tab '{tab}'.")
    return ws, rows


def mark_thread_posted(ws, row_number):
    update_cell(ws, row_number, "STATUS", "Posted")
    update_cell(ws, row_number, "POSTED_AT", now_iso())


def split_into_tweets(text, delimiter):
    parts = [p.strip() for p in text.split(delimiter)]
    parts = [p for p in parts if p]
    return parts if parts else [text.strip()]


# ── X / Playwright shared selectors ───────────────────────────────────────────

TEXTBOX_SELECTORS = [
    '[data-testid="tweetTextarea_0"]',
    '[data-testid="tweetTextarea_0EditorContainer"] div[contenteditable="true"]',
    'div[contenteditable="true"][data-testid]',
    'div[contenteditable="true"][aria-label]',
    'div[contenteditable="true"]',
    '[aria-label="Post text"]',
    '[aria-label="Tweet text"]',
    '[placeholder="What is happening?!"]',
    '[placeholder*="happening"]',
]

POST_BUTTON_SELECTORS = [
    '[data-testid="tweetButton"]',
    '[data-testid="tweetButtonInline"]',
    'button[data-testid*="tweet"]',
    'div[data-testid="tweetButton"]',
    'button:has-text("Post")',
    'button:has-text("Tweet")',
    '[aria-label="Post"]',
    '[aria-label="Tweet"]',
]

ADD_THREAD_TWEET_SELECTORS = [
    '[data-testid="addButton"]',
    'div[aria-label="Add post"]',
    'button[aria-label="Add post"]',
]

PREVIEW_SELECTORS = [
    '[data-testid="attachments"] video',
    '[data-testid="videoComponent"]',
    '[data-testid="tweetPhoto"]',
    '[data-testid="attachments"] img',
    '[data-testid="attachments"] [role="progressbar"]',
    '[data-testid="attachments"]',
    'img[src*="blob:"]',
    'video[src*="blob:"]',
]


def find_element_multi(page, selectors, label, timeout=15_000):
    for sel in selectors:
        try:
            el = page.locator(sel).first
            el.wait_for(state="visible", timeout=timeout // len(selectors))
            dbg(f"    ✓ Found {label} via: {sel}")
            return el
        except Exception:
            dbg(f"    ✗ Selector failed for {label}: {sel}")
    return None


def _textbox_visible(page):
    for sel in TEXTBOX_SELECTORS[:3]:
        try:
            if page.locator(sel).first.is_visible():
                return True
        except Exception:
            pass
    return False


def _is_genuinely_logged_out(page):
    try:
        page.goto("https://x.com/home", wait_until="domcontentloaded", timeout=20_000)
        page.wait_for_timeout(2_500)
        current = page.url
        if "login" in current or "signin" in current or "graduated-access" in current:
            return True
        return not _textbox_and_nav_healthy(page)
    except Exception as e:
        dbg(f"  [SESSION-CHECK] Recheck itself failed: {e} — treating as still logged out.")
        return True


def _textbox_and_nav_healthy(page):
    try:
        page.locator('[data-testid="AppTabBar_Home_Link"], [data-testid="SideNav_NewTweet_Button"]').first.wait_for(
            state="visible", timeout=8_000
        )
        return True
    except Exception:
        return False


def navigate_to_compose(page, post_index, attempt=1):
    for url in ["https://x.com/compose/post", "https://twitter.com/compose/tweet"]:
        dbg(f"  [NAV] Attempt {attempt}: navigating to {url}")
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=30_000)
            page.wait_for_timeout(3_000)
            current = page.url
            screenshot(page, f"p{post_index}_nav_attempt{attempt}")
            if "login" in current or "signin" in current or "graduated-access" in current:
                dbg("  [NAV] Landed on a login/restriction-looking URL — double-checking before aborting…")
                if _is_genuinely_logged_out(page):
                    raise RuntimeError("Session expired or account restricted (confirmed after recheck).")
                dbg("  [NAV] Recheck says session is actually fine — retrying compose navigation.")
                continue
            if "compose" in current or _textbox_visible(page):
                return True
        except RuntimeError:
            raise
        except Exception as e:
            dbg(f"  [NAV] Navigation to {url} failed: {e}")

    dbg("  [NAV] Trying home → compose button fallback…")
    try:
        page.goto("https://x.com/home", wait_until="domcontentloaded", timeout=30_000)
        page.wait_for_timeout(3_000)
        btn = page.locator(
            'a[href="/compose/post"], [data-testid="SideNav_NewTweet_Button"], '
            '[aria-label="Post"], a[aria-label*="compose"]'
        ).first
        btn.wait_for(state="visible", timeout=10_000)
        btn.click()
        page.wait_for_timeout(3_000)
        if _textbox_visible(page):
            return True
    except Exception as e:
        dbg(f"  [NAV] Home fallback failed: {e}")
    return False


def type_into_textbox(page, locator, text):
    locator.scroll_into_view_if_needed()
    locator.click()
    page.wait_for_timeout(400)
    page.keyboard.type(text, delay=15)
    page.wait_for_timeout(400)


def attach_media_robust(page, media_path, mime_type, post_index):
    is_video = mime_type in VIDEO_MIMES
    upload_timeout = 120_000 if is_video else 45_000
    file_inputs = page.locator('input[type="file"]')
    count = file_inputs.count()
    dbg(f"  [ATTACH] Found {count} file input(s).")

    if count > 0:
        for idx in range(count):
            inp = file_inputs.nth(idx)
            accept = inp.get_attribute("accept") or ""
            dbg(f"  [ATTACH] Input #{idx}: accept='{accept}'")
            if (is_video and ("video" in accept or accept == "")) or \
               (not is_video and ("image" in accept or accept == "")):
                try:
                    inp.set_input_files(media_path)
                    dbg(f"  [ATTACH] set_input_files on input #{idx} — success.")
                    if _wait_for_preview(page, upload_timeout, post_index, f"slot{idx}"):
                        return True
                except Exception as e:
                    dbg(f"  [ATTACH] input #{idx} failed: {e}")

        for idx in range(count):
            try:
                file_inputs.nth(idx).set_input_files(media_path)
                dbg(f"  [ATTACH] Fallback: set_input_files on input #{idx}.")
                if _wait_for_preview(page, upload_timeout, post_index, f"fallback{idx}"):
                    return True
            except Exception as e:
                dbg(f"  [ATTACH] Fallback input #{idx} failed: {e}")

    dbg("  [ATTACH] Trying media toolbar button…")
    for btn_sel in ['[data-testid="addMedia"]', '[aria-label*="edia"]', '[aria-label*="hoto"]']:
        try:
            btn = page.locator(btn_sel).first
            if btn.is_visible():
                btn.click()
                page.wait_for_timeout(1_000)
                fi2 = page.locator('input[type="file"]')
                if fi2.count() > 0:
                    fi2.first.set_input_files(media_path)
                    if _wait_for_preview(page, upload_timeout, post_index, "toolbar"):
                        return True
        except Exception as e:
            dbg(f"  [ATTACH] toolbar btn {btn_sel} failed: {e}")

    screenshot(page, f"p{post_index}_attach_failed")
    raise RuntimeError("All media attachment strategies failed.")


def _wait_for_preview(page, timeout, post_index, label):
    for sel in PREVIEW_SELECTORS:
        try:
            page.wait_for_selector(sel, timeout=timeout // len(PREVIEW_SELECTORS))
            try:
                page.wait_for_selector('[role="progressbar"]', state="detached", timeout=timeout)
            except Exception:
                pass
            page.wait_for_timeout(1_200)
            screenshot(page, f"p{post_index}_preview_{label}")
            return True
        except Exception:
            continue
    return False


def add_thread_tweet_box(page, post_index, tweet_index):
    dbg(f"  [THREAD] Adding tweet #{tweet_index + 1}…")
    btn = find_element_multi(page, ADD_THREAD_TWEET_SELECTORS, "add-thread-tweet button", timeout=10_000)
    if btn is None:
        screenshot(page, f"p{post_index}_no_add_button")
        raise RuntimeError("Could not find the 'Add post' (+) button to extend the thread.")
    btn.click()
    page.wait_for_timeout(800)
    locator = page.locator(f'[data-testid="tweetTextarea_{tweet_index}"]').first
    try:
        locator.wait_for(state="visible", timeout=10_000)
        return locator
    except Exception:
        all_ce = page.locator('div[contenteditable="true"][data-testid]')
        return all_ce.nth(all_ce.count() - 1)


def click_post_button(page, post_index):
    post_btn = find_element_multi(page, POST_BUTTON_SELECTORS, "post button", timeout=15_000)
    if post_btn is None:
        screenshot(page, f"p{post_index}_no_post_button")
        raise RuntimeError("Post button not found.")
    if post_btn.is_disabled():
        for _ in range(3):
            page.wait_for_timeout(5_000)
            if not post_btn.is_disabled():
                break
    try:
        post_btn.click()
    except Exception:
        page.evaluate(
            """(sel) => { for (const s of sel) { const el = document.querySelector(s); if (el) { el.click(); return s; } } return null; }""",
            POST_BUTTON_SELECTORS,
        )


def post_with_network_confirmation(page, post_index, click_timeout=25_000):
    result = {"sent": None, "tweet_id": None}

    def on_response(response):
        if "CreateTweet" not in response.url:
            return
        try:
            data = response.json()
        except Exception:
            return
        if isinstance(data, dict) and data.get("errors"):
            dbg(f"  [NET-CONFIRM] CreateTweet errors: {data['errors']}")
            result["sent"] = False
            return
        try:
            tweet_id = data["data"]["create_tweet"]["tweet_results"]["result"]["rest_id"]
            dbg(f"  [NET-CONFIRM] CreateTweet succeeded, tweet_id={tweet_id}")
            result["sent"] = True
            result["tweet_id"] = tweet_id
        except (KeyError, TypeError):
            result["sent"] = False

    page.on("response", on_response)
    try:
        click_post_button(page, post_index)
        waited = 0
        while waited < click_timeout:
            page.wait_for_timeout(500)
            waited += 500
            if result["sent"] is False:
                break
        page.wait_for_timeout(1_500)
    finally:
        page.remove_listener("response", on_response)

    screenshot(page, f"p{post_index}_after_click")
    if result["sent"]:
        return True, result["tweet_id"]
    return False, None


def post_one_job(page, job, post_index, max_attempts=3):
    session_error = False
    for attempt in range(1, max_attempts + 1):
        dbg(f"  ─── [{job['kind'].upper()}] attempt {attempt}/{max_attempts} ───")
        try:
            navigate_to_compose(page, post_index, attempt)

            first_box = find_element_multi(page, TEXTBOX_SELECTORS, "textbox", timeout=20_000)
            if first_box is None:
                raise RuntimeError("Could not find first tweet textbox.")
            type_into_textbox(page, first_box, job["tweets"][0])
            screenshot(page, f"p{post_index}_a{attempt}_tweet1_typed")

            if job["media_path"]:
                attach_media_robust(page, job["media_path"], job["mime_type"], post_index)

            for idx in range(1, len(job["tweets"])):
                box = add_thread_tweet_box(page, post_index, idx)
                type_into_textbox(page, box, job["tweets"][idx])
                screenshot(page, f"p{post_index}_a{attempt}_tweet{idx+1}_typed")

            sent, tweet_id = post_with_network_confirmation(page, post_index)
            if sent:
                dbg(f"  Post #{post_index} SUCCESS on attempt {attempt} (tweet_id={tweet_id}).")
                return True, False
            dbg(f"  Post #{post_index}: not confirmed — retrying.")
            if attempt < max_attempts:
                page.wait_for_timeout(5_000)

        except RuntimeError as e:
            dbg(f"  [ATTEMPT {attempt}] RuntimeError: {e}")
            if any(k in str(e).lower() for k in SESSION_ERROR_KEYWORDS):
                session_error = True
                if attempt < max_attempts:
                    dbg("  Session looked broken — waiting and retrying this job before treating it as fatal.")
                    page.wait_for_timeout(15_000)
                    continue
                else:
                    return False, True
            if attempt < max_attempts:
                page.wait_for_timeout(10_000)
            else:
                return False, False
        except Exception as e:
            dbg(f"  [ATTEMPT {attempt}] Unexpected error: {e}")
            try:
                screenshot(page, f"p{post_index}_a{attempt}_exception")
            except Exception:
                pass
            if attempt < max_attempts:
                page.wait_for_timeout(10_000)
            else:
                return False, False
    return False, session_error


# ── Job planning: turn percentages into a concrete list of posts ─────────────

def compute_type_counts(post_count, image_pct, video_pct, thread_pct):
    total_pct = image_pct + video_pct + thread_pct
    if total_pct <= 0:
        sys.exit("All three percentages (image/video/thread) are 0 — nothing to post.")
    img_frac = image_pct / total_pct
    vid_frac = video_pct / total_pct
    thr_frac = thread_pct / total_pct

    n_images = round(post_count * img_frac)
    n_videos = round(post_count * vid_frac)
    n_threads = post_count - n_images - n_videos
    if n_threads < 0:
        n_threads = 0

    if image_pct <= 0:
        n_images = 0
    if video_pct <= 0:
        n_videos = 0
    if thread_pct <= 0:
        n_threads = 0

    dbg(f"Normalized mix -> images:{img_frac:.0%} video:{vid_frac:.0%} thread:{thr_frac:.0%}")
    dbg(f"Planned counts -> images:{n_images} videos:{n_videos} threads:{n_threads} (target total {post_count})")
    return n_images, n_videos, n_threads


def main():
    run_start = time.time()
    run_deadline = run_start + RUN_BUDGET_MINUTES * 60

    dbg("=" * 60)
    dbg("Starting post_mix_to_x.py (sheet-driven config)")
    dbg(f"REPO: {REPO}  RUN_ID: {RUN_ID}")
    dbg(f"ACCOUNTS_SPREADSHEET_ID: {ACCOUNTS_SPREADSHEET_ID}")
    dbg("=" * 60)

    if not ACCOUNTS_SPREADSHEET_ID:
        sys.exit("ACCOUNTS_SPREADSHEET_ID is not set — this is the master accounts+settings spreadsheet ID.")

    gc = build_sheet_client()
    accounts_sh = gc.open_by_key(ACCOUNTS_SPREADSHEET_ID)
    accounts_ws = get_tab(accounts_sh, ACCOUNTS_SHEET_TAB, required_for="accounts sheet")
    master_settings = load_master_settings(accounts_sh)

    account_records = accounts_ws.get_all_records()
    if not account_records:
        sys.exit(f"No account rows found in '{ACCOUNTS_SHEET_TAB}' tab.")

    row_number, account_rec = claim_account(accounts_ws, account_records)
    cfg = build_effective_config(account_rec, master_settings)

    account_name = account_rec.get("ACCOUNT_NAME", f"row{row_number}")
    dbg(f"Using account: {account_name} (sheet row {row_number})")
    dbg(f"Effective config: POST_COUNT={cfg['POST_COUNT']} IMAGE%={cfg['IMAGE_PERCENT']} "
        f"VIDEO%={cfg['VIDEO_PERCENT']} THREAD%={cfg['THREAD_PERCENT']} "
        f"INTERVAL_MIN={cfg['INTERVAL_MINUTES']} SHUFFLE={cfg['SHUFFLE']}")
    dbg(f"Link config: URL={'set' if cfg['LINK_URL'] else 'none'} "
        f"IMAGE(enabled={cfg['LINK_ENABLE_IMAGE']}, pct={cfg['LINK_PERCENT_IMAGE']}) "
        f"VIDEO(enabled={cfg['LINK_ENABLE_VIDEO']}, pct={cfg['LINK_PERCENT_VIDEO']})")

    state_json = str(account_rec.get("X_STORAGE_STATE_JSON", "")).strip()
    if not state_json:
        sys.exit(f"Account row {row_number} ('{account_name}') has an empty X_STORAGE_STATE_JSON cell.")
    storage_state_path = "x_storage_state.json"
    with open(storage_state_path, "w", encoding="utf-8") as f:
        f.write(state_json)

    n_images, n_videos, n_threads = compute_type_counts(
        cfg["POST_COUNT"], cfg["IMAGE_PERCENT"], cfg["VIDEO_PERCENT"], cfg["THREAD_PERCENT"]
    )

    drive_service = None
    data_ws = None
    image_jobs, video_jobs = [], []
    if n_images or n_videos:
        drive_service = build_drive_service()
        data_ws, data_rows = load_data_rows(accounts_sh, cfg["DATA_SHEET_TAB"], account_name)
        drive_files = list_media_in_folder(drive_service, cfg["SOURCE_FOLDER_ID"], IMAGE_MIMES | VIDEO_MIMES)
        file_index = index_drive_files(drive_files)
        image_jobs, video_jobs = plan_data_jobs(data_rows, file_index, n_images, n_videos)
        if len(image_jobs) < n_images:
            dbg(f"⚠ Only found {len(image_jobs)} postable image row(s) — capping.")
        if len(video_jobs) < n_videos:
            dbg(f"⚠ Only found {len(video_jobs)} postable video row(s) — capping.")

    thread_ws = None
    thread_jobs = []
    if n_threads:
        thread_ws, thread_rows = load_thread_rows(accounts_sh, cfg["THREAD_SHEET_TAB"], account_name)
        if len(thread_rows) < n_threads:
            dbg(f"⚠ Only {len(thread_rows)} thread row(s) available — capping thread posts.")
            n_threads = len(thread_rows)
        if cfg["SHUFFLE"]:
            random.shuffle(thread_rows)
        thread_jobs = [{"kind": "thread", "sheet_row": r} for r in thread_rows[:n_threads]]

    jobs = image_jobs + video_jobs + thread_jobs
    if cfg["SHUFFLE"]:
        random.shuffle(jobs)

    if not jobs:
        dbg("No jobs to run after matching against sheet/Drive availability. Exiting.")
        update_cell(accounts_ws, row_number, "ASSIGNED_STATUS", "IDLE (no jobs this run)")
        return

    dbg(f"Final plan: {len(jobs)} post(s) -> images:{len(image_jobs)} videos:{len(video_jobs)} threads:{len(thread_jobs)}")

    results_summary = []
    consecutive_session_failures = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox", "--disable-setuid-sandbox",
                "--disable-blink-features=AutomationControlled",
                "--disable-web-security",
                "--disable-features=IsolateOrigins,site-per-process",
            ],
        )
        context = browser.new_context(
            storage_state=storage_state_path,
            viewport={"width": 1280, "height": 900},
            user_agent=(
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
            ),
            permissions=["notifications"],
        )
        context.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
            Object.defineProperty(navigator, 'plugins', { get: () => [1,2,3,4,5] });
        """)
        page = context.new_page()
        page.on("console", lambda msg: dbg(f"  [BROWSER {msg.type.upper()}] {msg.text[:200]}"))
        page.on("requestfailed", lambda req: dbg(f"  [NET FAIL] {req.url[:100]}"))

        for i, job in enumerate(jobs, start=1):
            if time.time() >= run_deadline:
                remaining = len(jobs) - i + 1
                dbg("")
                dbg(f"⏱ Reached RUN_BUDGET_MINUTES ({RUN_BUDGET_MINUTES} min) — stopping now.")
                dbg(f"   {remaining} job(s) left unrun this run; they'll be picked up next run.")
                break

            dbg("")
            dbg(f"{'='*50}")
            dbg(f"POST {i}/{len(jobs)}  kind={job['kind']}")
            dbg(f"{'='*50}")

            media_path, mime_type = None, None
            tweets = []
            tmp_to_clean = None
            fatal_session_error = False

            try:
                if job["kind"] in ("image", "video"):
                    drive_file = job["drive_file"]
                    data_row = job["data_row"]
                    mime_type = drive_file["mimeType"]
                    link = maybe_link(job["kind"], cfg)
                    text = build_text_from_data_row(data_row, link)
                    tweets = [text]

                    ext = Path(drive_file["name"]).suffix or (".mp4" if job["kind"] == "video" else ".jpg")
                    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                        media_path = tmp.name
                    tmp_to_clean = media_path
                    download_file(drive_service, drive_file["id"], drive_file["name"], media_path)

                else:  # thread
                    row = job["sheet_row"]
                    tweets = split_into_tweets(row["text"], cfg["THREAD_DELIMITER"])

                for j, t in enumerate(tweets, start=1):
                    dbg(f"  Tweet {j}: {t[:80]}{'…' if len(t) > 80 else ''}")

                job_payload = {"kind": job["kind"], "tweets": tweets, "media_path": media_path, "mime_type": mime_type}
                posted, session_error = post_one_job(page, job_payload, post_index=i)

                if posted:
                    dbg(f"POST {i}/{len(jobs)}: SUCCESS ✓")
                    results_summary.append((i, job["kind"], "SUCCESS"))
                    consecutive_session_failures = 0
                    if job["kind"] in ("image", "video"):
                        mark_data_posted(data_ws, job["data_row"]["row_number"])
                        move_to_claimed(drive_service, cfg["CLAIMED_FOLDER_ID"], job["drive_file"]["id"], job["drive_file"]["name"])
                    else:
                        mark_thread_posted(thread_ws, job["sheet_row"]["row_number"])
                elif session_error:
                    consecutive_session_failures += 1
                    dbg(f"POST {i}/{len(jobs)}: FAILED — looked like a session/login problem "
                        f"({consecutive_session_failures}/{MAX_CONSECUTIVE_SESSION_FAILURES} consecutive).")
                    results_summary.append((i, job["kind"], "FAILED (session)"))
                    if consecutive_session_failures >= MAX_CONSECUTIVE_SESSION_FAILURES:
                        dbg("Session repeatedly looks broken — aborting the rest of this run.")
                        dbg("Update X_STORAGE_STATE_JSON in this account's Accounts-sheet row and rerun.")
                        update_cell(accounts_ws, row_number, "ASSIGNED_STATUS", "SESSION_ERROR — needs fresh X_STORAGE_STATE_JSON")
                        fatal_session_error = True
                else:
                    dbg(f"POST {i}/{len(jobs)}: FAILED after retries.")
                    results_summary.append((i, job["kind"], "FAILED"))
                    consecutive_session_failures = 0

            except Exception as e:
                dbg(f"POST {i}/{len(jobs)}: EXCEPTION — {e}")
                results_summary.append((i, job["kind"], f"EXCEPTION: {e}"))

            finally:
                if tmp_to_clean:
                    try:
                        Path(tmp_to_clean).unlink(missing_ok=True)
                    except Exception:
                        pass

            if fatal_session_error:
                break

            if i < len(jobs):
                interval_seconds = int(cfg["INTERVAL_MINUTES"] * 60)
                if time.time() + interval_seconds >= run_deadline:
                    dbg(f"⏱ Next post would land after the {RUN_BUDGET_MINUTES}-min budget — stopping instead of sleeping into it.")
                    break
                dbg(f"Sleeping {interval_seconds}s ({cfg['INTERVAL_MINUTES']} min) before next post…")
                time.sleep(interval_seconds)

        browser.close()

    try:
        update_cell(accounts_ws, row_number, "ASSIGNED_STATUS", f"COMPLETED {now_iso()}")
    except Exception:
        pass

    dbg("")
    dbg("=" * 60)
    dbg("Run complete. Summary:")
    for idx, kind, status in results_summary:
        dbg(f"  Post {idx:>2} [{kind:<6}]: {status}")
    dbg(f"Total run time: {(time.time() - run_start) / 60:.1f} min")
    dbg(f"Debug screenshots saved in: {SCREENSHOT_DIR.resolve()}")
    dbg("=" * 60)


if __name__ == "__main__":
    main()
