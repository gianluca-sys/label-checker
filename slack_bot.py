#!/usr/bin/env python3
"""
Healf Label Checker — Slack Bot

Four modes — all triggered by typing the SKU as the message text:

  Mode 1 — Label vs Label (2 PDFs)
    Attach two PDF labels → compares them field by field

  Mode 2 — Label vs Warehouse photos (1 PDF + 1–20 images)
    Attach one PDF + warehouse photos → checks physical product
    against the approved label

  Mode 3 — Label vs Google Drive folder (1 PDF + Drive folder URL in text)
    Type:  SKU-001 https://drive.google.com/drive/folders/FOLDER_ID
    Attach one PDF → bot downloads images from Drive automatically

  Mode 4 — Warehouse vs Warehouse (2 Drive folder URLs, no PDF)
    Type:  SKU-001
           https://drive.google.com/drive/folders/FOLDER1
           https://drive.google.com/drive/folders/FOLDER2
    Compares two sets of warehouse photos against each other

Environment variables required:
  SLACK_BOT_TOKEN   — xoxb-...
  SLACK_APP_TOKEN   — xapp-...  (Socket Mode app-level token)
  ANTHROPIC_API_KEY — sk-ant-...

Optional (for Google Sheets logging):
  GOOGLE_SHEET_ID   — the ID from your sheet's URL
  GOOGLE_CREDS_FILE — path to your service account JSON (default: google_credentials.json)
"""

import base64
import os
import re
import sys
import json
import tempfile
import threading
import urllib.request
from datetime import datetime
from pathlib import Path

import anthropic
from flask import Flask, request, jsonify
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

sys.path.insert(0, str(Path(__file__).parent))
from label_compare import extract, compare


# ── Config ─────────────────────────────────────────────────────────────────────
SLACK_BOT_TOKEN   = os.environ.get("SLACK_BOT_TOKEN", "")
SLACK_APP_TOKEN   = os.environ.get("SLACK_APP_TOKEN", "")
GOOGLE_SHEET_ID   = os.environ.get("GOOGLE_SHEET_ID", "")
GOOGLE_CREDS_FILE = os.environ.get("GOOGLE_CREDS_FILE", "google_credentials.json")
GOOGLE_CREDS_JSON = os.environ.get("GOOGLE_CREDS_JSON", "")  # base64-encoded service account JSON (for cloud)
API_SECRET_KEY    = os.environ.get("API_SECRET_KEY", "")     # shared secret for the /run-checks endpoint

IMAGE_MIMETYPES = {"image/jpeg", "image/png", "image/gif", "image/webp"}
IMAGE_EXTENSIONS = {".jpg": "image/jpeg", ".jpeg": "image/jpeg",
                    ".png": "image/png", ".gif": "image/gif", ".webp": "image/webp"}

app      = App(token=SLACK_BOT_TOKEN)
http_app = Flask(__name__)


# ── Google credential loader (file on disk or base64 env var) ─────────────────
def _get_google_creds(scopes):
    """Returns a Credentials object from GOOGLE_CREDS_JSON env var or local file."""
    from google.oauth2.service_account import Credentials
    if GOOGLE_CREDS_JSON:
        raw  = base64.b64decode(GOOGLE_CREDS_JSON).decode()
        info = json.loads(raw)
        return Credentials.from_service_account_info(info, scopes=scopes)
    if os.path.exists(GOOGLE_CREDS_FILE):
        return Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
    return None


# ── Google Sheets logging ──────────────────────────────────────────────────────
def _get_sheets_client():
    if not GOOGLE_SHEET_ID:
        return None, None
    scopes = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = _get_google_creds(scopes)
    if not creds:
        return None, None
    import gspread
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    return gc, sh


def _ensure_ws(sh, name, rows, cols, headers):
    import gspread
    try:
        ws = sh.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(name, rows=rows, cols=cols)
        ws.append_row(headers)
        ws.format("1", {"textFormat": {"bold": True},
                        "backgroundColor": {"red": 0.176, "green": 0.416, "blue": 0.310}})
    return ws


def _log_to_sheets(sku, name1, name2, result, mode):
    if not GOOGLE_SHEET_ID:
        return False
    if not GOOGLE_CREDS_JSON and not os.path.exists(GOOGLE_CREDS_FILE):
        print("Google Sheets: no credentials available — skipping.")
        return False
    try:
        gc, sh = _get_sheets_client()
        if not sh:
            return False

        ts     = datetime.now().strftime("%Y-%m-%d %H:%M")
        status = "CHANGES" if result["total_differences"] else "MATCH"

        # ── Change Log tab (one row per run) ──────────────────────────────────
        ws_log = _ensure_ws(sh, "Change Log", 1000, 8,
                            ["Timestamp", "SKU", "Mode", "Source 1 / US Label",
                             "Source 2 / UK Label", "Changes Found", "Total Fields", "Status"])
        ws_log.append_row([
            ts, sku, mode, name1, name2,
            result["total_differences"],
            result["total_differences"] + result["total_matches"],
            status,
        ])

        # ── Per-SKU detail tab (full field-by-field comparison) ───────────────
        tab_name = sku[:100]
        import gspread
        try:
            sh.del_worksheet(sh.worksheet(tab_name))
        except gspread.WorksheetNotFound:
            pass

        ws_sku = sh.add_worksheet(tab_name, rows=500, cols=6)

        # Row 1 — title
        ws_sku.append_row([f"Label Check — {sku}", "", "", "", "", ""])
        ws_sku.format("A1", {
            "textFormat": {"bold": True, "fontSize": 14, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "backgroundColor": {"red": 0.176, "green": 0.416, "blue": 0.310},
            "horizontalAlignment": "CENTER",
        })

        # Row 2 — metadata
        uk_mode_log = "US→UK" in mode
        src1_label  = "US Label" if uk_mode_log else "Source 1"
        src2_label  = "UK Label" if uk_mode_log else "Source 2"
        ws_sku.append_row([f"Run: {ts}", f"Mode: {mode}", f"{src1_label}: {name1}",
                           f"{src2_label}: {name2}", f"Status: {status}", ""])
        ws_sku.format("A2:F2", {
            "textFormat": {"italic": True, "fontSize": 10},
            "backgroundColor": {"red": 0.847, "green": 0.953, "blue": 0.863},
        })

        # Row 3 — column headers
        ws_sku.append_row(["Category", "Field", src1_label, src2_label, "Status", "Notes"])
        ws_sku.format("A3:F3", {
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "backgroundColor": {"red": 0.176, "green": 0.416, "blue": 0.310},
            "horizontalAlignment": "CENTER",
        })
        ws_sku.freeze(rows=3)

        # Differences
        if result["differences"]:
            ws_sku.append_row([f"CHANGES FOUND ({result['total_differences']})", "", "", "", "", ""])
            current_row_count = len(ws_sku.get_all_values())
            ws_sku.format(f"A{current_row_count}:F{current_row_count}", {
                "textFormat": {"bold": True, "foregroundColor": {"red": 0.8, "green": 0, "blue": 0}},
                "backgroundColor": {"red": 1, "green": 0.9, "blue": 0.9},
                "horizontalAlignment": "CENTER",
            })

            for d in result["differences"]:
                try:
                    row_num = len(ws_sku.get_all_values()) + 1
                    note_text = d.get("note", "")
                    ws_sku.append_row([d["category"], d["field"], d["current"], d["new"],
                                       "CRITICAL" if d.get("critical") else "CHANGED",
                                       note_text])
                    bg = {"red": 1, "green": 0.878, "blue": 0.878} if d.get("critical") else {"red": 1, "green": 1, "blue": 1}
                    ws_sku.format(f"A{row_num}:F{row_num}", {"backgroundColor": bg})
                    if d.get("critical"):
                        ws_sku.format(f"E{row_num}", {
                            "textFormat": {"bold": True, "foregroundColor": {"red": 0.8, "green": 0, "blue": 0}},
                        })
                    if note_text:
                        ws_sku.format(f"F{row_num}", {
                            "textFormat": {"italic": True, "fontSize": 9,
                                           "foregroundColor": {"red": 0.3, "green": 0.3, "blue": 0.6}},
                        })
                except Exception as row_err:
                    print(f"Sheet row write error (skipping): {row_err}")



        return True
    except Exception as e:
        print(f"Google Sheets error: {e}")
        return False


# ── Warehouse photo extraction ─────────────────────────────────────────────────
WAREHOUSE_PROMPT = """These are warehouse photos of a physical product. Extract every piece of text
visible on the label EXACTLY as printed — word for word, character for character.

Return ONLY a JSON object (no markdown, no explanation) with this structure:
{
  "brand_name": "...",
  "product_name": "...",
  "net_weight": "...",
  "serving_size": "...",
  "servings_per_container": "...",
  "calories": "...",
  "nutrients": [{"name": "...", "amount": "...", "dv_percent": "..."}],
  "ingredients": "...",
  "allergens": "...",
  "suggested_use": "...",
  "other_claims": ["..."]
}

Rules:
- If text is split across multiple photos, combine what is visible.
- Use null for any field not visible in any photo.
- Extract nutrients in the order they appear on the label.
- suggested_use: full text of 'Suggested Use', 'Directions', 'Recommended Use', or 'How To Use'.
- allergens: include any allergen warnings embedded within the ingredients text."""


def extract_from_images(client, image_paths):
    content = []
    for path in image_paths:
        ext = Path(path).suffix.lower()
        media_type = IMAGE_EXTENSIONS.get(ext, "image/jpeg")
        with open(path, "rb") as f:
            data = base64.standard_b64encode(f.read()).decode()
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": media_type, "data": data},
        })
    content.append({"type": "text", "text": WAREHOUSE_PROMPT})

    response = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=4096,
        messages=[{"role": "user", "content": content}],
    )
    text = response.content[0].text
    m = re.search(r"\{.*\}", text, re.DOTALL)
    if not m:
        raise ValueError(f"No JSON found in Claude response for warehouse photos")
    return json.loads(m.group())


# ── Slack result formatter ─────────────────────────────────────────────────────
def _format_result(sku, name1, name2, result, mode, comp_mode="us_us"):
    total_d = result["total_differences"]
    total_m = result["total_matches"]
    diffs   = result["differences"]

    uk_tag = "  🇬🇧 _US → UK comparison_" if comp_mode == "us_uk" else ""

    if mode == "warehouse":
        header2 = f"PDF: `{name1}`   Warehouse: `{name2}`{uk_tag}"
    elif mode == "warehouse_vs_warehouse":
        header2 = f"Set 1: `{name1}`   Set 2: `{name2}`{uk_tag}"
    else:
        header2 = f"US label: `{name1}`   {'UK label' if comp_mode == 'us_uk' else 'New label'}: `{name2}`{uk_tag}"

    lines = [
        f"*Label Check — {sku}*",
        header2,
        "─" * 44,
    ]

    if total_d == 0:
        lines.append(f"✅  *No differences found* ({total_m} fields checked)")
    else:
        critical     = [d for d in diffs if d.get("critical")]
        non_critical = [d for d in diffs if not d.get("critical")]

        if critical:
            lines.append(f"🔴  *{len(critical)} critical change{'s' if len(critical) != 1 else ''}*")
            lines.append("")
            for diff in critical[:10]:
                curr = diff["current"][:120]
                new  = diff["new"][:120]
                lines.append(f"• *{diff['field']}*")
                if mode == "warehouse":
                    lines.append(f"  PDF:       _{curr}_")
                    lines.append(f"  Warehouse: _{new}_")
                elif mode == "warehouse_vs_warehouse":
                    lines.append(f"  Set 1: _{curr}_")
                    lines.append(f"  Set 2: _{new}_")
                else:
                    lines.append(f"  Current: _{curr}_")
                    lines.append(f"  New:     _{new}_")
            if len(critical) > 10:
                lines.append(f"_…and {len(critical) - 10} more critical changes_")

        if non_critical:
            if critical:
                lines.append("")
            lines.append(f"🟡  *{len(non_critical)} other change{'s' if len(non_critical) != 1 else ''}*")
            lines.append("")
            for diff in non_critical[:5]:
                curr = diff["current"][:80]
                new  = diff["new"][:80]
                lines.append(f"• *{diff['field']}*")
                if mode == "warehouse":
                    lines.append(f"  PDF:       _{curr}_")
                    lines.append(f"  Warehouse: _{new}_")
                elif mode == "warehouse_vs_warehouse":
                    lines.append(f"  Set 1: _{curr}_")
                    lines.append(f"  Set 2: _{new}_")
                else:
                    lines.append(f"  Current: _{curr}_")
                    lines.append(f"  New:     _{new}_")
            if len(non_critical) > 5:
                lines.append(f"_…and {len(non_critical) - 5} more non-critical changes_")

    return "\n".join(lines)


# ── File download ──────────────────────────────────────────────────────────────
def _download(url, dest_path):
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {SLACK_BOT_TOKEN}"})
    with urllib.request.urlopen(req) as resp, open(dest_path, "wb") as f:
        f.write(resp.read())


# ── Google Drive folder download ───────────────────────────────────────────────
DRIVE_FOLDER_RE = re.compile(r"drive\.google\.com/drive/folders/([a-zA-Z0-9_-]+)")

def _parse_drive_folder_id(text):
    m = DRIVE_FOLDER_RE.search(text or "")
    return m.group(1) if m else None


def _download_drive_images(folder_id, dest_dir):
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    import io

    scopes = ["https://www.googleapis.com/auth/drive.readonly"]
    creds  = _get_google_creds(scopes)
    if not creds:
        raise ValueError("No Google credentials available. Set GOOGLE_CREDS_JSON environment variable.")
    svc    = build("drive", "v3", credentials=creds, cache_discovery=False)

    resp  = svc.files().list(
        q=f"'{folder_id}' in parents and mimeType contains 'image/' and trashed = false",
        fields="files(id, name, mimeType)",
        pageSize=20,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = resp.get("files", [])

    if not files:
        raise ValueError("No images found in the Google Drive folder. "
                         "Make sure the folder is shared with the service account.")

    paths = []
    for f in files[:20]:
        dest = os.path.join(dest_dir, f["name"])
        req  = svc.files().get_media(fileId=f["id"], supportsAllDrives=True)
        with open(dest, "wb") as fh:
            dl = MediaIoBaseDownload(fh, req)
            done = False
            while not done:
                _, done = dl.next_chunk()
        paths.append(dest)

    return paths


# ── Mode 1: PDF vs PDF ─────────────────────────────────────────────────────────
def _process_pdf_vs_pdf(client, sku, file1, file2, channel, thread_ts, comp_mode="us_us"):
    try:
        ac = anthropic.Anthropic()
        with tempfile.TemporaryDirectory() as tmp:
            path1 = os.path.join(tmp, file1["name"])
            path2 = os.path.join(tmp, file2["name"])
            _download(file1["url_private_download"], path1)
            _download(file2["url_private_download"], path2)
            d1     = extract(ac, path1)
            d2     = extract(ac, path2)
            result = compare(d1, d2, mode=comp_mode)

        display_mode = "Label vs Label (US→UK)" if comp_mode == "us_uk" else "Label vs Label"
        text   = _format_result(sku, file1["name"], file2["name"], result, "pdf", comp_mode)
        logged = _log_to_sheets(sku, file1["name"], file2["name"], result, display_mode)
        if logged:
            text += "\n\n📊  _Logged to Google Sheets_"
        client.chat_postMessage(channel=channel, thread_ts=thread_ts, text=text)

    except Exception as e:
        client.chat_postMessage(channel=channel, thread_ts=thread_ts,
                                text=f"❌  Something went wrong: {e}")


# ── Mode 2: PDF vs Warehouse photos ───────────────────────────────────────────
def _process_pdf_vs_warehouse(client, sku, pdf_file, image_files, channel, thread_ts, comp_mode="us_us"):
    try:
        ac = anthropic.Anthropic()
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = os.path.join(tmp, pdf_file["name"])
            _download(pdf_file["url_private_download"], pdf_path)

            img_paths = []
            for img in image_files:
                p = os.path.join(tmp, img["name"])
                _download(img["url_private_download"], p)
                img_paths.append(p)

            d_pdf       = extract(ac, pdf_path)
            d_warehouse = extract_from_images(ac, img_paths)
            result      = compare(d_pdf, d_warehouse, mode=comp_mode)

        photo_names  = ", ".join(f["name"] for f in image_files)
        display_mode = "Label vs Warehouse (US→UK)" if comp_mode == "us_uk" else "Label vs Warehouse"
        text         = _format_result(sku, pdf_file["name"], photo_names, result, "warehouse", comp_mode)
        logged       = _log_to_sheets(sku, pdf_file["name"], photo_names, result, display_mode)
        if logged:
            text += "\n\n📊  _Logged to Google Sheets_"
        client.chat_postMessage(channel=channel, thread_ts=thread_ts, text=text)

    except Exception as e:
        client.chat_postMessage(channel=channel, thread_ts=thread_ts,
                                text=f"❌  Something went wrong: {e}")


# ── Mode 4: Warehouse vs Warehouse (2 Drive folders) ──────────────────────────
def _process_warehouse_vs_warehouse(client, sku, folder_id1, folder_id2, channel, thread_ts, comp_mode="us_us"):
    try:
        ac = anthropic.Anthropic()
        with tempfile.TemporaryDirectory() as tmp:
            dir1 = os.path.join(tmp, "set1")
            dir2 = os.path.join(tmp, "set2")
            os.makedirs(dir1)
            os.makedirs(dir2)

            paths1 = _download_drive_images(folder_id1, dir1)
            paths2 = _download_drive_images(folder_id2, dir2)

            d1 = extract_from_images(ac, paths1)
            d2 = extract_from_images(ac, paths2)
            result = compare(d1, d2, mode=comp_mode)

        display_mode = "Warehouse vs Warehouse (US→UK)" if comp_mode == "us_uk" else "Warehouse vs Warehouse"
        text   = _format_result(sku,
                                f"Folder 1 ({len(paths1)} photos)",
                                f"Folder 2 ({len(paths2)} photos)",
                                result, "warehouse_vs_warehouse", comp_mode)
        logged = _log_to_sheets(sku,
                                f"Drive:{folder_id1}",
                                f"Drive:{folder_id2}",
                                result, display_mode)
        if logged:
            text += "\n\n📊  _Logged to Google Sheets_"
        client.chat_postMessage(channel=channel, thread_ts=thread_ts, text=text)

    except Exception as e:
        client.chat_postMessage(channel=channel, thread_ts=thread_ts,
                                text=f"❌  Something went wrong: {e}")


# ── Mode 3: PDF vs Google Drive folder ────────────────────────────────────────
def _process_pdf_vs_drive(client, sku, pdf_file, folder_id, channel, thread_ts, comp_mode="us_us"):
    try:
        ac = anthropic.Anthropic()
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path  = os.path.join(tmp, pdf_file["name"])
            _download(pdf_file["url_private_download"], pdf_path)

            img_paths = _download_drive_images(folder_id, tmp)

            d_pdf       = extract(ac, pdf_path)
            d_warehouse = extract_from_images(ac, img_paths)
            result      = compare(d_pdf, d_warehouse, mode=comp_mode)

        display_mode = "Label vs Drive Folder (US→UK)" if comp_mode == "us_uk" else "Label vs Drive Folder"
        text      = _format_result(sku, pdf_file["name"],
                                   f"Drive folder ({len(img_paths)} photos)", result, "warehouse", comp_mode)
        logged    = _log_to_sheets(sku, pdf_file["name"], f"Drive:{folder_id}",
                                   result, display_mode)
        if logged:
            text += "\n\n📊  _Logged to Google Sheets_"
        client.chat_postMessage(channel=channel, thread_ts=thread_ts, text=text)

    except Exception as e:
        client.chat_postMessage(channel=channel, thread_ts=thread_ts,
                                text=f"❌  Something went wrong: {e}")


# ── Slack event handler ────────────────────────────────────────────────────────
@app.event("message")
def handle_message(event, client):
    files     = event.get("files", [])
    raw_text  = (event.get("text") or "").strip()
    channel   = event["channel"]
    ts        = event["ts"]

    pdf_files   = [f for f in files if f.get("mimetype") == "application/pdf"]
    img_files   = [f for f in files if f.get("mimetype") in IMAGE_MIMETYPES]
    folder_ids  = DRIVE_FOLDER_RE.findall(raw_text)

    # Detect comparison type from "US vs US" or "US vs UK" prefix
    is_us_uk = bool(re.search(r'US\s+vs\s+UK', raw_text, re.IGNORECASE))
    is_us_us = bool(re.search(r'US\s+vs\s+US', raw_text, re.IGNORECASE))

    # If files/folders are attached but no mode specified, ask the user to clarify
    has_content = len(pdf_files) > 0 or len(img_files) > 0 or len(folder_ids) > 0
    if has_content and not is_us_uk and not is_us_us:
        client.chat_postMessage(
            channel=channel, thread_ts=ts,
            text=(
                "⚠️  Please specify the comparison type at the start of your message:\n\n"
                "• *US vs US*  `DFH-MGC-120` — comparing two US labels\n"
                "• *US vs UK*  `DFH-MGC-120` — comparing a US label against a UK label\n\n"
                "Attach your files again with the correct prefix and I'll get started."
            ),
        )
        return

    comp_mode = "us_uk" if is_us_uk else "us_us"

    # SKU = message text stripped of mode prefix, Drive URLs, and Slack URL formatting
    sku = re.sub(r"<https?://[^>]*>", "", raw_text)
    sku = DRIVE_FOLDER_RE.sub("", sku)
    sku = re.sub(r'US\s+vs\s+UK', '', sku, flags=re.IGNORECASE)
    sku = re.sub(r'US\s+vs\s+US', '', sku, flags=re.IGNORECASE)
    sku = sku.strip() or "UNKNOWN-SKU"

    mode_label = "  🇬🇧 *US vs UK*" if is_us_uk else "  🇺🇸 *US vs US*"

    # Mode 1: exactly 2 PDFs
    if len(pdf_files) == 2 and len(img_files) == 0:
        client.chat_postMessage(
            channel=channel, thread_ts=ts,
            text=f"📋  Comparing labels for *{sku}*{mode_label} — give me ~30 seconds…",
        )
        threading.Thread(
            target=_process_pdf_vs_pdf,
            args=(client, sku, pdf_files[0], pdf_files[1], channel, ts, comp_mode),
            daemon=True,
        ).start()

    # Mode 2: 1 PDF + 1–20 images attached directly
    elif len(pdf_files) == 1 and 1 <= len(img_files) <= 20:
        client.chat_postMessage(
            channel=channel, thread_ts=ts,
            text=f"📸  Checking warehouse photos against PDF label for *{sku}*{mode_label} — give me ~45 seconds…",
        )
        threading.Thread(
            target=_process_pdf_vs_warehouse,
            args=(client, sku, pdf_files[0], img_files, channel, ts, comp_mode),
            daemon=True,
        ).start()

    # Mode 3: 1 PDF + 1 Google Drive folder link
    elif len(pdf_files) == 1 and len(folder_ids) == 1:
        client.chat_postMessage(
            channel=channel, thread_ts=ts,
            text=f"📂  Fetching images from Drive and checking against PDF for *{sku}*{mode_label} — give me ~45 seconds…",
        )
        threading.Thread(
            target=_process_pdf_vs_drive,
            args=(client, sku, pdf_files[0], folder_ids[0], channel, ts, comp_mode),
            daemon=True,
        ).start()

    # Mode 4: 2 Google Drive folder links, no PDF
    elif len(pdf_files) == 0 and len(folder_ids) == 2:
        client.chat_postMessage(
            channel=channel, thread_ts=ts,
            text=f"📂  Comparing two warehouse photo sets for *{sku}*{mode_label} — give me ~60 seconds…",
        )
        threading.Thread(
            target=_process_warehouse_vs_warehouse,
            args=(client, sku, folder_ids[0], folder_ids[1], channel, ts, comp_mode),
            daemon=True,
        ).start()

    # Unrecognised combination — give the user a hint
    elif len(pdf_files) > 0 or len(img_files) > 0 or folder_ids:
        client.chat_postMessage(
            channel=channel, thread_ts=ts,
            text=(
                "⚠️  I wasn't sure what to do with those files. Here's what I accept:\n\n"
                "*Format:*  `US vs US  SKU`  or  `US vs UK  SKU`\n\n"
                "*Then attach one of:*\n"
                "• 2 PDFs — label vs label\n"
                "• 1 PDF + 1–20 photos — PDF vs warehouse photos\n"
                "• 1 PDF + Drive folder link — PDF vs Drive folder\n"
                "• 2 Drive folder links (no PDF) — folder vs folder\n\n"
                "*US vs UK:* attach US label first, UK label second."
            ),
        )


# ── HTTP endpoints (for Google Sheets button) ──────────────────────────────────
@http_app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@http_app.route("/run-checks", methods=["POST"])
def trigger_run_checks():
    if API_SECRET_KEY and request.headers.get("X-API-Key") != API_SECRET_KEY:
        return jsonify({"error": "Unauthorized"}), 401

    def _run():
        try:
            from run_checks import run
            run()
        except Exception as e:
            print(f"Sheet run-checks error: {e}")

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({
        "status": "queued",
        "message": "Label checks started — results will appear in the sheet shortly.",
    })


# ── Entry point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    missing = [v for v in ("SLACK_BOT_TOKEN", "SLACK_APP_TOKEN", "ANTHROPIC_API_KEY")
               if not os.environ.get(v)]
    if missing:
        sys.exit(f"Error: missing environment variable(s): {', '.join(missing)}")

    print("Starting Healf Label Checker…")

    # Slack bot runs in a background thread (Socket Mode — no port required)
    def _start_slack():
        print("Slack bot connecting…")
        SocketModeHandler(app, SLACK_APP_TOKEN).start()

    threading.Thread(target=_start_slack, daemon=True).start()

    # Flask HTTP server runs on main thread (Railway web service requires PORT)
    port = int(os.environ.get("PORT", 8080))
    print(f"HTTP server listening on port {port}…")
    http_app.run(host="0.0.0.0", port=port, use_reloader=False)
