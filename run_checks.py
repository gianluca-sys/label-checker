#!/usr/bin/env python3
"""
Healf Label Checker — Sheet Runner

Reads the "Approved for Mass print" production queue, finds rows where
"Run?" (column Z) is ticked, runs the label comparison, and writes
"Audit Check" (AA) and "Comments" (AB) back into the sheet.

Usage:
    python run_checks.py

Required environment variables:
    ANTHROPIC_API_KEY
    GOOGLE_CREDS_JSON   base64-encoded service account JSON (cloud / Railway)
    — or —
    GOOGLE_CREDS_FILE   path to service account JSON file (default: google_credentials.json)

Column Z — Check Mode (type one of):
    UK vs Warehouse      compares approved UK label (col F) vs warehouse photos (col G)
    UK vs US Label       compares approved UK label (col F) vs US AI file (col W)
    UK vs Change Mgmt    compares approved UK label (col F) vs change-mgmt AI file (col X)
"""

import base64
import json
import os
import re
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import anthropic

sys.path.insert(0, str(Path(__file__).parent))
from label_compare import extract, compare

# ── Config ─────────────────────────────────────────────────────────────────────
PRODUCTION_SHEET_ID = "1gQNo_8U4r5qBp_nhGQzEssOh9Mzd9jgsemlIZ_y2yDk"
SHEET_TAB           = "Ready to Print"

GOOGLE_CREDS_JSON = os.environ.get("GOOGLE_CREDS_JSON", "")
GOOGLE_CREDS_FILE = os.environ.get("GOOGLE_CREDS_FILE", "google_credentials.json")

# Column positions — 1-based (gspread convention)
COL_SKU          =  4   # D  Variant SKU
COL_NAME         =  5   # E  Product Name
COL_UK_LABEL     =  7   # G  Draft Artwork Link         — approved UK PDF
COL_WAREHOUSE    = 10   # J  Warehouse Photos            — Drive folder link
COL_US_LABEL     = 11   # K  US AI File                 — original US label PDF
COL_CHANGE_MGMT  = 12   # L  Change Management AI file  — new brand label PDF
COL_CHECK_MODE   = 13   # M  comparison mode dropdown
COL_RUN          = 14   # N  tick to queue
COL_AUDIT_CHECK  = 15   # O  result written by script
COL_COMMENTS     = 16   # P  detail written by script

DRIVE_FILE_RE   = re.compile(r"(?:/file/d/|[?&]id=)([a-zA-Z0-9_-]+)")
DRIVE_FOLDER_RE = re.compile(r"drive\.google\.com/drive/folders/([a-zA-Z0-9_-]+)")

IMAGE_EXTENSIONS = {
    ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
    ".png": "image/png",  ".gif":  "image/gif",
    ".webp": "image/webp",
}

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
- dv_percent should be a plain number only (e.g. "80", not "80%" or "80% DV").
- suggested_use: full text of Suggested Use, Directions, Recommended Use, or How To Use.
- allergens: include any allergen warnings embedded within the ingredients text."""


# ── Credentials ────────────────────────────────────────────────────────────────
def _get_creds(scopes):
    from google.oauth2.service_account import Credentials
    if GOOGLE_CREDS_JSON:
        info = json.loads(base64.b64decode(GOOGLE_CREDS_JSON).decode())
        return Credentials.from_service_account_info(info, scopes=scopes)
    if os.path.exists(GOOGLE_CREDS_FILE):
        return Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
    sys.exit("No Google credentials. Set GOOGLE_CREDS_JSON or GOOGLE_CREDS_FILE.")


def _drive_svc():
    from googleapiclient.discovery import build
    return build("drive", "v3",
                 credentials=_get_creds(["https://www.googleapis.com/auth/drive.readonly"]),
                 cache_discovery=False)


def _sheets_client():
    import gspread
    return gspread.authorize(_get_creds([
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]))


# ── Drive helpers ──────────────────────────────────────────────────────────────
def _file_id(url):
    m = DRIVE_FILE_RE.search(url or "")
    return m.group(1) if m else None


def _folder_id(url):
    m = DRIVE_FOLDER_RE.search(url or "")
    return m.group(1) if m else None


def _dl_file(svc, fid, dest):
    from googleapiclient.http import MediaIoBaseDownload
    req = svc.files().get_media(fileId=fid, supportsAllDrives=True)
    with open(dest, "wb") as fh:
        dl = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = dl.next_chunk()


def _dl_folder_images(svc, fid, dest_dir):
    resp = svc.files().list(
        q=f"'{fid}' in parents and mimeType contains 'image/' and trashed = false",
        fields="files(id, name, mimeType)",
        pageSize=20,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = resp.get("files", [])
    if not files:
        raise ValueError(f"No images found in Drive folder {fid}. "
                         "Check the folder is shared with the service account.")
    paths = []
    for f in files[:20]:
        dest = os.path.join(dest_dir, f["name"])
        _dl_file(svc, f["id"], dest)
        paths.append(dest)
    return paths


# ── Warehouse image → label data ───────────────────────────────────────────────
def _extract_images(ac, paths):
    content = []
    for path in paths:
        ext = Path(path).suffix.lower()
        media_type = IMAGE_EXTENSIONS.get(ext, "image/jpeg")
        with open(path, "rb") as f:
            data = base64.standard_b64encode(f.read()).decode()
        content.append({"type": "image",
                         "source": {"type": "base64", "media_type": media_type, "data": data}})
    content.append({"type": "text", "text": WAREHOUSE_PROMPT})
    resp = ac.messages.create(
        model="claude-sonnet-4-6", max_tokens=4096,
        messages=[{"role": "user", "content": content}],
    )
    text = resp.content[0].text
    m = re.search(r"\{.*\}", text, re.DOTALL)
    if not m:
        raise ValueError("No JSON in Claude response for warehouse images")
    return json.loads(m.group())


# ── Row helper ─────────────────────────────────────────────────────────────────
def _cell(row, col_1based):
    idx = col_1based - 1
    return row[idx].strip() if idx < len(row) else ""


# ── Main ───────────────────────────────────────────────────────────────────────
def run():
    print("Connecting to Google Sheets…")
    gc = _sheets_client()
    ws = gc.open_by_key(PRODUCTION_SHEET_ID).worksheet(SHEET_TAB)

    all_rows = ws.get_all_values()
    queued = [
        (i + 2, row)                          # +2: skip header + convert to 1-based
        for i, row in enumerate(all_rows[1:]) # skip header row
        if _cell(row, COL_RUN).upper() in ("TRUE", "YES", "1", "✓", "TRUE")
    ]

    if not queued:
        print("No rows queued. Tick 'Run?' (column Z = TRUE) for the SKUs you want to check.")
        return

    print(f"Found {len(queued)} SKU(s) to process.\n")
    ac  = anthropic.Anthropic()
    svc = _drive_svc()

    for row_num, row in queued:
        sku       = _cell(row, COL_SKU)
        name      = _cell(row, COL_NAME)
        mode_text = _cell(row, COL_CHECK_MODE) or "UK vs Warehouse"

        print(f"▶  [{sku}] {name}  —  {mode_text}")
        ws.update_cell(row_num, COL_AUDIT_CHECK, "⏳ Running…")
        ws.update_cell(row_num, COL_COMMENTS, "")

        try:
            with tempfile.TemporaryDirectory() as tmp:

                # ── Download label 1 (UK label for UK modes; skipped for US vs US) ──
                d_uk = None
                if mode_text != "US vs US":
                    uk_url = _cell(row, COL_UK_LABEL)
                    uk_fid = _file_id(uk_url)
                    if not uk_fid:
                        raise ValueError(f"Cannot extract file ID from UK label URL: {uk_url!r}")
                    uk_path = os.path.join(tmp, f"{sku}_uk.pdf")
                    _dl_file(svc, uk_fid, uk_path)
                    d_uk = extract(ac, uk_path)

                # ── Download comparison target based on mode ───────────────────
                if mode_text == "UK vs Warehouse":
                    wh_url = _cell(row, COL_WAREHOUSE)
                    fol_id = _folder_id(wh_url)
                    if fol_id:
                        img_paths = _dl_folder_images(svc, fol_id, tmp)
                    else:
                        single_fid = _file_id(wh_url)
                        if not single_fid:
                            raise ValueError(f"Cannot parse warehouse URL: {wh_url!r}")
                        img_path = os.path.join(tmp, f"{sku}_wh.jpg")
                        _dl_file(svc, single_fid, img_path)
                        img_paths = [img_path]
                    d_other  = _extract_images(ac, img_paths)
                    source2  = f"Warehouse ({len(img_paths)} photo{'s' if len(img_paths) != 1 else ''})"

                elif mode_text == "UK vs US Label":
                    us_url = _cell(row, COL_US_LABEL)
                    us_fid = _file_id(us_url)
                    if not us_fid:
                        raise ValueError(f"Cannot extract US label file ID from: {us_url!r}")
                    us_path = os.path.join(tmp, f"{sku}_us.pdf")
                    _dl_file(svc, us_fid, us_path)
                    d_other = extract(ac, us_path)
                    source2 = "US Label"

                elif mode_text == "UK vs Change Mgmt":
                    cm_url = _cell(row, COL_CHANGE_MGMT)
                    cm_fid = _file_id(cm_url)
                    if not cm_fid:
                        raise ValueError(f"Cannot extract Change Mgmt file ID from: {cm_url!r}")
                    cm_path = os.path.join(tmp, f"{sku}_cm.pdf")
                    _dl_file(svc, cm_fid, cm_path)
                    d_other = extract(ac, cm_path)
                    source2 = "Change Mgmt Label"

                elif mode_text == "US vs US":
                    us_url = _cell(row, COL_US_LABEL)
                    us_fid = _file_id(us_url)
                    if not us_fid:
                        raise ValueError(f"Cannot extract US AI file ID from: {us_url!r}")
                    us_path = os.path.join(tmp, f"{sku}_us.pdf")
                    _dl_file(svc, us_fid, us_path)
                    d_uk = extract(ac, us_path)  # reuse d_uk slot — first label is US AI file

                    cm_url = _cell(row, COL_CHANGE_MGMT)
                    cm_fid = _file_id(cm_url)
                    if not cm_fid:
                        raise ValueError(f"Cannot extract Change Mgmt file ID from: {cm_url!r}")
                    cm_path = os.path.join(tmp, f"{sku}_cm.pdf")
                    _dl_file(svc, cm_fid, cm_path)
                    d_other = extract(ac, cm_path)
                    source2 = "New Brand Label (Change Mgmt)"

                else:
                    raise ValueError(
                        f"Unknown Check Mode: {mode_text!r}\n"
                        "Valid values: 'UK vs Warehouse', 'UK vs US Label', "
                        "'UK vs Change Mgmt', 'US vs US'"
                    )

                comp_mode = "us_us" if mode_text == "US vs US" else "us_uk"
                result = compare(d_uk, d_other, mode=comp_mode)

            n_diffs  = result["total_differences"]
            critical = [d for d in result["differences"] if d.get("critical")]

            if n_diffs == 0:
                audit    = "✅ Pass — No action required"
                comments = f"All checked fields match ({result['total_matches']} fields verified against {source2})."
            else:
                n_crit   = len(critical)
                n_other  = n_diffs - n_crit
                audit    = f"❌ Fail — {n_crit} critical, {n_other} other"
                lines    = []
                for d in result["differences"][:15]:
                    tag  = "⚠️ CRITICAL" if d.get("critical") else "ℹ️  Change"
                    curr = str(d["current"])[:100]
                    new  = str(d["new"])[:100]
                    note = f"  Note: {d['note']}" if d.get("note") else ""
                    lines.append(f"{tag} — {d['field']}\n  Was: {curr}\n  Now: {new}{note}")
                if n_diffs > 15:
                    lines.append(f"…and {n_diffs - 15} more. Run in Slack for full detail.")
                comments = "\n\n".join(lines)

            ws.update_cell(row_num, COL_AUDIT_CHECK, audit)
            ws.update_cell(row_num, COL_COMMENTS,    comments)
            ws.update_cell(row_num, COL_RUN,         "FALSE")
            print(f"   → {audit}\n")

        except Exception as e:
            ws.update_cell(row_num, COL_AUDIT_CHECK, "❌ Error")
            ws.update_cell(row_num, COL_COMMENTS,    str(e))
            ws.update_cell(row_num, COL_RUN,         "FALSE")
            print(f"   → ERROR: {e}\n")

    print(f"Done — {datetime.now().strftime('%Y-%m-%d %H:%M')}")


if __name__ == "__main__":
    if not os.environ.get("ANTHROPIC_API_KEY"):
        sys.exit("Error: ANTHROPIC_API_KEY environment variable not set.")
    run()
