"""
Bluesky -> Mega.nz + Google Sheets scraper

Scrapes ORIGINAL posts (no reposts) from a single Bluesky username, downloads
all matching images/videos locally, then uploads them ALL to Mega.nz in one
batch pass at the end, and logs filename + mega link + caption + hashtags to
a Google Sheet in a single batched append call.

Environment variables:
    BSKY_HANDLE, BSKY_APP_PASSWORD   - Bluesky login
    TARGET_USERNAME                  - handle to scrape
    MODE                              - "timeline" or "media"
    CONTENT_TYPE                     - "images", "videos", or "both"
    MAX_POSTS                        - how many posts to scan (default 100)
    HASHTAG_COUNT                    - max hashtags to save per post (default 3)
    MEGA_EMAIL, MEGA_PASSWORD        - Mega.nz login (put in GitHub Secrets!)
    MEGA_FOLDER_NAME                 - Mega folder name (created if missing)
    GOOGLE_APPLICATION_CREDENTIALS   - path to Google OAuth user token JSON
    GOOGLE_SHEET_ID                  - target Google Sheet ID
"""

import json
import os
import re
import subprocess
import sys
import time
from datetime import datetime

import requests
from atproto import Client
from mega import Mega

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# ---------------------------------------------------------------------------
# Config from environment
# ---------------------------------------------------------------------------
_INVISIBLE_CHARS_RE = re.compile(r"[\u200b-\u200f\u202a-\u202e\u2060-\u2064\ufeff]")


def clean_str(value):
    if value is None:
        return ""
    return _INVISIBLE_CHARS_RE.sub("", value).strip()


BSKY_HANDLE = clean_str(os.environ.get("BSKY_HANDLE", ""))
BSKY_APP_PASSWORD = os.environ.get("BSKY_APP_PASSWORD", "").strip()
TARGET_USERNAME = clean_str(os.environ.get("TARGET_USERNAME", ""))
MODE = os.environ.get("MODE", "timeline").strip().lower()
CONTENT_TYPE = os.environ.get("CONTENT_TYPE", "both").strip().lower()
MAX_POSTS = int(os.environ.get("MAX_POSTS", "100") or 100)
HASHTAG_COUNT = int(os.environ.get("HASHTAG_COUNT", "3") or 3)

MEGA_EMAIL = os.environ.get("MEGA_EMAIL", "").strip()
MEGA_PASSWORD = os.environ.get("MEGA_PASSWORD", "").strip()
MEGA_FOLDER_NAME = clean_str(os.environ.get("MEGA_FOLDER_NAME", ""))

GOOGLE_CREDS_PATH = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "google_creds.json")
GOOGLE_SHEET_ID = os.environ.get("GOOGLE_SHEET_ID", "").strip()

DOWNLOAD_DIR = "downloads"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

SHEET_HEADER = ["File Name", "Mega Link", "Type", "Caption", "Hashtags"]


def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


def fail(msg):
    log(f"FATAL: {msg}")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Google Sheets helpers
# ---------------------------------------------------------------------------
def get_sheets_service():
    if not os.path.exists(GOOGLE_CREDS_PATH):
        fail(f"Google credentials file not found at {GOOGLE_CREDS_PATH}")
    with open(GOOGLE_CREDS_PATH) as f:
        info = json.load(f)
    creds = Credentials(
        token=info.get("token"),
        refresh_token=info.get("refresh_token"),
        token_uri=info.get("token_uri"),
        client_id=info.get("client_id"),
        client_secret=info.get("client_secret"),
        scopes=info.get("scopes", SCOPES),
    )
    if not creds.valid:
        if creds.refresh_token:
            creds.refresh(Request())
            log("🔄 Refreshed Google OAuth access token")
        else:
            fail("Google credentials are invalid/expired and no refresh_token is present")
    return build("sheets", "v4", credentials=creds)


def get_first_sheet_title(sheets_service, sheet_id):
    meta = sheets_service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    return meta["sheets"][0]["properties"]["title"]


def ensure_sheet_header(sheets_service, sheet_id, sheet_title):
    existing = (
        sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=sheet_id, range=f"{sheet_title}!A1:E1")
        .execute()
        .get("values", [])
    )
    if not existing:
        sheets_service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{sheet_title}!A1:E1",
            valueInputOption="RAW",
            body={"values": [SHEET_HEADER]},
        ).execute()


def append_sheet_rows_batch(sheets_service, sheet_id, sheet_title, rows):
    """Single API call for every row collected during the run."""
    if not rows:
        log("ℹ️ No rows to log to Sheets.")
        return
    sheets_service.spreadsheets().values().append(
        spreadsheetId=sheet_id,
        range=f"{sheet_title}!A:E",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": rows},
    ).execute()
    log(f"📝 Logged {len(rows)} rows to sheet '{sheet_title}' in one batch.")


# ---------------------------------------------------------------------------
# Mega.nz helpers
# ---------------------------------------------------------------------------
def get_mega_client():
    if not MEGA_EMAIL or not MEGA_PASSWORD:
        fail("MEGA_EMAIL / MEGA_PASSWORD are not set")
    mega = Mega()
    m = mega.login(MEGA_EMAIL, MEGA_PASSWORD)
    log("✅ Mega login successful!")
    return m


def get_or_create_mega_folder(m, folder_name):
    if not folder_name:
        return None
    existing = m.find(folder_name)
    if existing:
        log(f"📁 Found existing Mega folder '{folder_name}'")
        return existing[0]
    created = m.create_folder(folder_name)
    folder_id = created.get(folder_name) if isinstance(created, dict) else None
    log(f"📁 Created new Mega folder '{folder_name}'")
    return folder_id


def upload_file_to_mega(m, filepath, folder_id):
    """Returns (filename, share_link) or (filename, None) if link retrieval fails."""
    fname = os.path.basename(filepath)
    try:
        if folder_id:
            uploaded = m.upload(filepath, folder_id)
        else:
            uploaded = m.upload(filepath)
        try:
            link = m.get_upload_link(uploaded)
        except Exception:
            link = None
        return fname, link
    except Exception as e:
        log(f"  ⚠️ Mega upload failed for {fname}: {e}")
        return fname, None


# ---------------------------------------------------------------------------
# Bluesky helpers (unchanged)
# ---------------------------------------------------------------------------
def get_video_playlist_url(post_view):
    try:
        embed = getattr(post_view, "embed", None)
        if not embed:
            return None
        playlist = getattr(embed, "playlist", None)
        if playlist:
            return str(playlist)
        media = getattr(embed, "media", None)
        if media:
            playlist = getattr(media, "playlist", None)
            if playlist:
                return str(playlist)
    except Exception:
        pass
    return None


def get_image_urls(post_view):
    urls = []
    try:
        embed = getattr(post_view, "embed", None)
        if not embed:
            return urls
        embed_type = getattr(embed, "$type", "") or getattr(embed, "py_type", "") or str(type(embed))
        if "images" in embed_type.lower():
            for img in getattr(embed, "images", []) or []:
                url = getattr(img, "fullsize", None) or getattr(img, "thumb", None)
                if url:
                    urls.append(url)
        media = getattr(embed, "media", None)
        if media:
            media_type = getattr(media, "$type", "") or getattr(media, "py_type", "") or str(type(media))
            if "images" in media_type.lower():
                for img in getattr(media, "images", []) or []:
                    url = getattr(img, "fullsize", None) or getattr(img, "thumb", None)
                    if url:
                        urls.append(url)
    except Exception:
        pass
    return urls


def download_binary(url, filepath, timeout=30):
    try:
        r = requests.get(url, stream=True, timeout=timeout)
        if r.status_code == 200:
            with open(filepath, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
            return True
    except Exception:
        pass
    return False


def download_video(playlist_url, filepath, timeout=120):
    if not playlist_url:
        return False
    try:
        result = subprocess.run(
            ["ffmpeg", "-y", "-loglevel", "error", "-i", playlist_url, "-c", "copy", filepath],
            capture_output=True, text=True, timeout=timeout,
        )
        if result.returncode == 0 and os.path.exists(filepath) and os.path.getsize(filepath) > 10000:
            return True
        if result.stderr:
            log(f"  ffmpeg error: {result.stderr.strip()[-300:]}")
    except subprocess.TimeoutExpired:
        log("  ffmpeg timed out downloading video")
    except FileNotFoundError:
        fail("ffmpeg is not installed on this runner - required to download HLS videos")
    except Exception as e:
        log(f"  ffmpeg exception: {e}")
    if os.path.exists(filepath):
        os.remove(filepath)
    return False


def extract_hashtags(text, limit):
    tags = re.findall(r"#(\w+)", text or "")
    if limit <= 0:
        return tags
    return tags[:limit]


def is_repost(feed_item):
    reason = getattr(feed_item, "reason", None)
    if not reason:
        return False
    reason_type = getattr(reason, "$type", "") or getattr(reason, "py_type", "") or str(type(reason))
    return "reasonRepost" in reason_type


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    if not BSKY_HANDLE or not BSKY_APP_PASSWORD:
        fail("BSKY_HANDLE / BSKY_APP_PASSWORD are not set")
    if not TARGET_USERNAME:
        fail("TARGET_USERNAME is not set")
    if not GOOGLE_SHEET_ID:
        fail("GOOGLE_SHEET_ID is not set")
    if MODE not in ("timeline", "media"):
        fail("MODE must be 'timeline' or 'media'")
    if CONTENT_TYPE not in ("images", "videos", "both"):
        fail("CONTENT_TYPE must be 'images', 'videos', or 'both'")

    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    log("🔑 Setting up Google Sheets service...")
    sheets_service = get_sheets_service()
    sheet_title = get_first_sheet_title(sheets_service, GOOGLE_SHEET_ID)
    ensure_sheet_header(sheets_service, GOOGLE_SHEET_ID, sheet_title)

    log(f"🔑 Logging in to Bluesky as {BSKY_HANDLE}...")
    client = Client()
    client.login(BSKY_HANDLE, BSKY_APP_PASSWORD)
    log("✅ Bluesky login successful!")

    try:
        profile = client.app.bsky.actor.get_profile(params={"actor": TARGET_USERNAME})
        target_did = profile.did
    except Exception as e:
        fail(f"Could not resolve target username '{TARGET_USERNAME}': {e}")

    log(f"🔍 Scraping '{TARGET_USERNAME}' | mode={MODE} | content={CONTENT_TYPE} "
        f"| max_posts={MAX_POSTS} | hashtags={HASHTAG_COUNT}")

    feed_filter = "posts_with_media" if MODE == "media" else "posts_no_replies"

    cursor = None
    scanned = 0
    # Each entry: {"path": ..., "type": ..., "caption": ..., "hashtags": ...}
    pending_uploads = []

    while scanned < MAX_POSTS:
        try:
            resp = client.app.bsky.feed.get_author_feed(
                params={"actor": TARGET_USERNAME, "filter": feed_filter, "limit": 30, "cursor": cursor}
            )
        except Exception as e:
            log(f"⚠️ Feed fetch error: {e}")
            break

        if not resp.feed:
            log(f"ℹ️ No more posts available - stopping (scanned {scanned} of {MAX_POSTS}).")
            break

        for item in resp.feed:
            if scanned >= MAX_POSTS:
                break
            scanned += 1

            if is_repost(item):
                continue

            post_view = item.post
            author_did = getattr(post_view.author, "did", None)
            if author_did != target_did:
                continue

            record = getattr(post_view, "record", None)
            text = getattr(record, "text", "") or ""
            post_cid = getattr(post_view, "cid", None)
            safe_name = re.sub(r"[^a-zA-Z0-9_-]", "_", post_cid or str(scanned))
            hashtags = extract_hashtags(text, HASHTAG_COUNT)
            hashtags_str = ", ".join(f"#{h}" for h in hashtags)

            if CONTENT_TYPE in ("videos", "both"):
                playlist_url = get_video_playlist_url(post_view)
                if playlist_url:
                    fname = f"{safe_name}.mp4"
                    fpath = os.path.join(DOWNLOAD_DIR, fname)
                    log(f"⬇️ Downloading video: {text[:50]}...")
                    if download_video(playlist_url, fpath):
                        pending_uploads.append(
                            {"path": fpath, "type": "video", "caption": text, "hashtags": hashtags_str}
                        )
                    else:
                        log(f"  ❌ Failed to download video for post {safe_name}")

            if CONTENT_TYPE in ("images", "both"):
                img_urls = get_image_urls(post_view)
                for i, url in enumerate(img_urls):
                    fname = f"{safe_name}_img{i + 1}.jpg"
                    fpath = os.path.join(DOWNLOAD_DIR, fname)
                    log(f"⬇️ Downloading image {i + 1}: {text[:50]}...")
                    if download_binary(url, fpath):
                        pending_uploads.append(
                            {"path": fpath, "type": "image", "caption": text, "hashtags": hashtags_str}
                        )
                    else:
                        log(f"  ❌ Failed to download image for post {safe_name}")

        cursor = getattr(resp, "cursor", None)
        if not cursor:
            log(f"ℹ️ Reached the end of '{TARGET_USERNAME}' feed - stopping "
                f"(scanned {scanned} of {MAX_POSTS}).")
            break
        time.sleep(0.5)

    log(f"📦 Download phase complete: {len(pending_uploads)} files ready for Mega upload.")

    # -------------------------------------------------------------
    # Batch upload phase - everything downloaded, now push it all at once
    # -------------------------------------------------------------
    sheet_rows = []
    if pending_uploads:
        m = get_mega_client()
        folder_id = get_or_create_mega_folder(m, MEGA_FOLDER_NAME) if MEGA_FOLDER_NAME else None

        for entry in pending_uploads:
            fpath = entry["path"]
            fname, link = upload_file_to_mega(m, fpath, folder_id)
            log(f"  ✅ Uploaded to Mega: {fname}" + (f" ({link})" if link else ""))
            sheet_rows.append([fname, link or "", entry["type"], entry["caption"], entry["hashtags"]])
            if os.path.exists(fpath):
                os.remove(fpath)

    append_sheet_rows_batch(sheets_service, GOOGLE_SHEET_ID, sheet_title, sheet_rows)

    log(f"🎉 Done! Scanned {scanned} posts, uploaded {len(sheet_rows)} files to Mega, "
        f"logged to sheet '{sheet_title}'.")


if __name__ == "__main__":
    main()
