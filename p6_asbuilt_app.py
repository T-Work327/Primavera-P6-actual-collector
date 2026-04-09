"""
Primavera P6 Asbuilt Data Collector — Streamlit Web App
=========================================================
Run with:
    streamlit run p6_asbuilt_app.py

Requirements:
    pip install streamlit openpyxl

──────────────────────────────────────────────────────────
USER MANAGEMENT
──────────────────────────────────────────────────────────
Edit the USERS dictionary below to add, remove, or change
passwords and roles. Passwords are stored as SHA-256 hashes.

To generate a hash for a new password, run in Python:
    import hashlib
    print(hashlib.sha256("yourpassword".encode()).hexdigest())

Roles:
  "readonly"   — View entries only
  "readwrite"  — View + Submit/Update + Import from Excel + Photo Log
  "admin"      — All of the above + Export to Excel
──────────────────────────────────────────────────────────
"""

import hashlib
import io
import json
import uuid
from datetime import date, datetime, time
from pathlib import Path

import streamlit as st

try:
    from PIL import Image, ImageOps
    _PILLOW = True
except ImportError:
    _PILLOW = False

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("openpyxl is required.  Run:  pip install openpyxl")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# USER DEFINITIONS  —  edit this section to manage users
# ══════════════════════════════════════════════════════════════════════════════
#
# To change a password:
#   1. Run in Python:  import hashlib; print(hashlib.sha256("newpassword".encode()).hexdigest())
#   2. Replace the hash string below
#
# To add a user:  copy any line and change username, hash, role, and display name
# To remove a user:  delete their entry

def _h(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()

USERS = {
    #  username       password hash           role           display name
    "admin":    {"hash": _h("admin123"),    "role": "admin",     "name": "Administrator"},
    "engineer": {"hash": _h("engineer1"),   "role": "readwrite", "name": "Site Engineer"},
    "viewer":   {"hash": _h("viewer1"),     "role": "readonly",  "name": "Project Viewer"},
}

# ── Role Permission Matrix ─────────────────────────────────────────────────────
PERMISSIONS = {
    "readonly":  {"view"},
    "readwrite": {"view", "submit", "import", "photos"},
    "admin":     {"view", "submit", "import", "export", "photos"},
}

def has_permission(perm: str) -> bool:
    return perm in PERMISSIONS.get(st.session_state.get("role", ""), set())

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════

DATA_FILE  = Path("p6_asbuilt_store.json")
PHOTO_DIR  = Path("p6_images")
PHOTO_FILE   = Path("p6_photo_log.json")
ASSIGN_FILE  = Path("p6_photo_assignments.json")

USER_DATA = "DurationQtyType=QT_Day\nShowAsPercentage=0\nSmallScaleQtyType=QT_Hour\nDateFormat=dd/mm/yyyy\nCurrencyFormat=US Dollar"

STATUS_OPTIONS = ["Not Started", "In Progress", "Completed"]

STATUS_COLOUR = {
    "Not Started": "#6b7280",
    "In Progress":  "#d97706",
    "Completed":    "#16a34a",
}

ROLE_LABEL = {
    "readonly":  "Read Only",
    "readwrite": "Read / Write",
    "admin":     "Admin",
}

# P6 internal field key names (row 1 of TASK sheet)
P6_FIELD_KEYS = [
    "task_code", "task_name", "status_code", "act_start_date",
    "act_end_date", "complete_pct", "remain_drtn_hr_cnt",
    "complete_pct_type", "wbs_id", "user_field_910",
]

# Column definitions: (display header, column width, data key)
P6_COLUMNS = [
    ("Activity ID",          14, "activity_id"),
    ("Activity Name",        36, "activity_name"),
    ("Activity Status",      16, "activity_status"),
    ("Actual Start",         20, "actual_start"),
    ("Actual Finish",        20, "actual_finish"),
    ("Duration % Complete",  20, "pct_complete"),
    ("Remaining Duration",   20, "remaining_dur"),
    ("Percent Complete Type",20, "complete_pct_type"),
    ("WBS Code",             36, "wbs_id"),
    ("Comments",             50, "comments_export"),
]

DATE_KEYS = {"actual_start", "actual_finish"}

# ══════════════════════════════════════════════════════════════════════════════
# DATE HELPERS
# Dates stored in JSON as ISO strings: "YYYY-MM-DDTHH:MM:00"
# ══════════════════════════════════════════════════════════════════════════════

def dt_to_iso(dt: datetime) -> str:
    return dt.strftime("%Y-%m-%dT%H:%M:00")

def iso_to_dt(value: str) -> datetime | None:
    if not value or str(value).strip() == "":
        return None
    for fmt in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:00", "%Y-%m-%dT%H:%M",
                "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
                "%d-%b-%y %H:%M", "%d-%b-%y", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(value).strip(), fmt)
        except ValueError:
            continue
    return None

def display_dt(value: str) -> str:
    dt = iso_to_dt(value)
    return dt.strftime("%d/%m/%Y %H:%M") if dt else "—"

def normalise_imported_date(raw_val) -> str:
    if raw_val is None or str(raw_val).strip() == "":
        return ""
    if isinstance(raw_val, datetime):
        return dt_to_iso(raw_val)
    dt = iso_to_dt(str(raw_val))
    return dt_to_iso(dt) if dt else str(raw_val).strip()


# ══════════════════════════════════════════════════════════════════════════════
# COMMENT HELPERS
# Comments are stored on each entry as a list of dicts:
#   _comments: [{"text": "...", "by": "...", "at": "DD/MM/YYYY HH:MM"}, ...]
# Newest first.  Exported to P6 as a single ';'-separated string (no timestamps).
# ══════════════════════════════════════════════════════════════════════════════

def comments_to_export(comments: list[dict]) -> str:
    """Flatten comment list → '; '-joined string, newest first, no timestamps."""
    return "; ".join(c["text"] for c in comments if c.get("text", "").strip())

def import_string_to_comments(raw: str, imported_by: str) -> list[dict]:
    """
    Split a '; '-separated import string into individual comment records.
    Each segment gets the same import timestamp and is marked as imported.
    Order is preserved (P6 exports newest first, so we keep that).
    """
    if not raw or not raw.strip():
        return []
    segments = [s.strip() for s in raw.split(";") if s.strip()]
    ts = datetime.now().strftime("%d/%m/%Y %H:%M")
    return [{"text": seg, "by": f"{imported_by} (imported)", "at": ts}
            for seg in segments]

# ══════════════════════════════════════════════════════════════════════════════
# STORAGE
# ══════════════════════════════════════════════════════════════════════════════

def load_entries() -> list[dict]:
    if DATA_FILE.exists():
        try:
            return json.loads(DATA_FILE.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return []
    return []

def save_entries(entries: list[dict]) -> None:
    DATA_FILE.write_text(json.dumps(entries, ensure_ascii=False, indent=2), encoding="utf-8")

def upsert_entry(entries: list[dict], new: dict) -> tuple:
    idx = next(
        (i for i, e in enumerate(entries)
         if e.get("activity_id", "").upper() == new.get("activity_id", "").upper()),
        None,
    )
    if idx is not None:
        entries[idx] = new
        return entries, "updated"
    entries.append(new)
    return entries, "saved"

# ══════════════════════════════════════════════════════════════════════════════
# DUPLICATE DETECTION
# ══════════════════════════════════════════════════════════════════════════════

# Fields compared when deciding whether incoming data is identical to stored.
# Identity fields (activity_id, name, wbs) and metadata (_submitted_at etc.)
# are intentionally excluded — we only care about progress data changing.
_PROGRESS_FIELDS = (
    "activity_status", "actual_start", "actual_finish",
    "pct_complete", "remaining_dur",
)

def is_exact_duplicate(incoming: dict, stored: dict) -> bool:
    """Return True if all progress fields are identical between incoming and stored."""
    for field in _PROGRESS_FIELDS:
        # Normalise: strip whitespace, treat None and "" as equivalent
        a = str(incoming.get(field) or "").strip()
        b = str(stored.get(field)   or "").strip()
        if a != b:
            return False
    return True

# ══════════════════════════════════════════════════════════════════════════════
# PHOTO STORAGE
#
# Two-file model:
#   p6_photo_log.json         — one record per image file (no activity link)
#   p6_photo_assignments.json — many-to-many: {photo_id, activity_id}
#
# This means one image file is stored once and can be assigned to any number
# of activities without duplication.
# ══════════════════════════════════════════════════════════════════════════════

def ensure_photo_dir() -> None:
    PHOTO_DIR.mkdir(exist_ok=True)

# ── Photos (image records) ─────────────────────────────────────────────────

def load_photos() -> list[dict]:
    if PHOTO_FILE.exists():
        try:
            return json.loads(PHOTO_FILE.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return []
    return []

def save_photos(photos: list[dict]) -> None:
    PHOTO_FILE.write_text(json.dumps(photos, ensure_ascii=False, indent=2), encoding="utf-8")

THUMB_SIZE = (400, 400)   # max thumbnail dimensions

def upload_photo(photo_date: date, comment: str,
                 file_bytes: bytes, original_name: str,
                 uploaded_by: str) -> dict:
    """Save image + thumbnail and create a photo record. Does NOT assign to any activity."""
    ensure_photo_dir()
    base_id  = uuid.uuid4().hex
    ext      = Path(original_name).suffix.lower() or ".jpg"
    filename = f"{base_id}{ext}"
    thumb    = f"{base_id}_thumb.jpg"
    dest     = PHOTO_DIR / filename
    dest_t   = PHOTO_DIR / thumb

    if _PILLOW and ext in (".jpg", ".jpeg", ".png", ".webp"):
        img = Image.open(io.BytesIO(file_bytes))
        img = ImageOps.exif_transpose(img)
        img.save(dest)
        # Generate thumbnail — convert to RGB so PNG/WEBP save as JPEG cleanly
        t = img.copy()
        t.thumbnail(THUMB_SIZE, Image.LANCZOS)
        t.convert("RGB").save(dest_t, "JPEG", quality=75, optimize=True)
    else:
        dest.write_bytes(file_bytes)
        thumb = ""   # no thumbnail for GIF/unsupported

    record = {
        "id":          base_id,
        "photo_date":  photo_date.isoformat(),
        "comment":     comment.strip(),
        "filename":    filename,
        "thumb":       thumb,
        "uploaded_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "uploaded_by": uploaded_by,
    }
    photos = load_photos()
    photos.append(record)
    save_photos(photos)
    return record

def delete_photo_file(photo_id: str) -> None:
    """Delete the image file and all its assignments."""
    photos = load_photos()
    record = next((p for p in photos if p["id"] == photo_id), None)
    if record:
        img_path = PHOTO_DIR / record["filename"]
        if img_path.exists():
            img_path.unlink()
        thumb = record.get("thumb", "")
        if thumb:
            t_path = PHOTO_DIR / thumb
            if t_path.exists():
                t_path.unlink()
        save_photos([p for p in photos if p["id"] != photo_id])
        # Remove all assignments for this photo
        assignments = load_assignments()
        save_assignments([a for a in assignments if a["photo_id"] != photo_id])

# ── Assignments (many-to-many link) ───────────────────────────────────────

def load_assignments() -> list[dict]:
    if ASSIGN_FILE.exists():
        try:
            return json.loads(ASSIGN_FILE.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return []
    return []

def save_assignments(assignments: list[dict]) -> None:
    ASSIGN_FILE.write_text(
        json.dumps(assignments, ensure_ascii=False, indent=2), encoding="utf-8"
    )

@st.cache_data(show_spinner=False)
def load_image_bytes(filename: str) -> bytes | None:
    """Load image file bytes once and cache. Cache cleared on upload/delete."""
    if not filename:
        return None
    path = PHOTO_DIR / filename
    return path.read_bytes() if path.exists() else None

def assign_photo(photo_id: str, activity_ids: list[str], assigned_by: str) -> None:
    """Add assignments for a photo to a list of activities (skip duplicates)."""
    assignments = load_assignments()
    existing    = {(a["photo_id"], a["activity_id"].upper()) for a in assignments}
    new_records = []
    for aid in activity_ids:
        if (photo_id, aid.upper()) not in existing:
            new_records.append({
                "photo_id":    photo_id,
                "activity_id": aid,
                "assigned_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "assigned_by": assigned_by,
            })
    if new_records:
        assignments.extend(new_records)
        save_assignments(assignments)
        # Update session_state cache so gallery reflects change without full rerun
        if "photo_assignments" in st.session_state:
            st.session_state["photo_assignments"] = assignments

def unassign_photo(photo_id: str, activity_id: str) -> None:
    """Remove a single photo→activity assignment. Image file is kept."""
    assignments = load_assignments()
    updated = [
        a for a in assignments
        if not (a["photo_id"] == photo_id
                and a["activity_id"].upper() == activity_id.upper())
    ]
    save_assignments(updated)
    # Update session_state cache so gallery reflects change without full rerun
    if "photo_assignments" in st.session_state:
        st.session_state["photo_assignments"] = updated

def photos_for_activity(activity_id: str) -> list[dict]:
    """Return all photo records assigned to a given activity."""
    assignments = load_assignments()
    photo_map   = {p["id"]: p for p in load_photos()}
    return [
        photo_map[a["photo_id"]]
        for a in assignments
        if a["activity_id"].upper() == activity_id.upper()
        and a["photo_id"] in photo_map
    ]

def activities_for_photo(photo_id: str) -> list[str]:
    """Return list of activity_ids assigned to a given photo."""
    return [a["activity_id"] for a in load_assignments()
            if a["photo_id"] == photo_id]

# ══════════════════════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ══════════════════════════════════════════════════════════════════════════════

# P6 internal status codes — must NOT contain spaces (0x20) or P6 rejects the import.
STATUS_TO_P6 = {
    "Not Started": "TK_NotStart",
    "In Progress":  "TK_Active",
    "Completed":    "TK_Complete",
}

THIN   = Side(style="thin", color="B0B8C8")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

@st.cache_data
def build_excel(entries: list[dict], project_name: str = "") -> bytes:
    """Build P6-ready XLSX.
    If project_name is supplied, the WBS prefix on every row is replaced with it.
    e.g. project_name="ProjectB" turns "ProjectA.1.2.3" → "ProjectB.1.2.3"
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TASK"
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 28
    ws.freeze_panes = "A3"

    # Row 1 — P6 internal field keys
    for col_idx, (key, (_, width, _)) in enumerate(zip(P6_FIELD_KEYS, P6_COLUMNS), start=1):
        c = ws.cell(row=1, column=col_idx, value=key)
        c.font      = Font(name="Arial", italic=True, color="4472C4", size=9)
        c.fill      = PatternFill("solid", fgColor="D9E1F2")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = BORDER
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Row 2 — Human-readable column headers
    for col_idx, (header, _, _) in enumerate(P6_COLUMNS, start=1):
        c = ws.cell(row=2, column=col_idx, value=header)
        c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill      = PatternFill("solid", fgColor="1C3557")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = BORDER

    # Data rows
    for row_idx, entry in enumerate(entries, start=3):
        fill = PatternFill("solid", fgColor="EEF2F8" if row_idx % 2 == 0 else "FFFFFF")
        for col_idx, (_, _, key) in enumerate(P6_COLUMNS, start=1):
            value = entry.get(key, "")
            if key in DATE_KEYS:
                dt_val = iso_to_dt(value)
                c = ws.cell(row=row_idx, column=col_idx, value=dt_val)
                if dt_val:
                    c.number_format = "DD/MM/YYYY HH:MM"
            elif key == "complete_pct_type":
                c = ws.cell(row=row_idx, column=col_idx, value="Physical")
            elif key == "comments_export":
                # Build export string from stored comment list (newest first, no timestamps)
                c = ws.cell(row=row_idx, column=col_idx,
                            value=comments_to_export(entry.get("_comments", [])))
            elif key == "wbs_id" and project_name and value:
                # Replace existing prefix with the supplied project name
                new_wbs = project_name.strip() + "." + strip_wbs_prefix(str(value))
                c = ws.cell(row=row_idx, column=col_idx, value=new_wbs)
            else:
                c = ws.cell(row=row_idx, column=col_idx, value=str(value) if value != "" else "")
            c.fill   = fill
            c.border = BORDER
            c.alignment = Alignment(vertical="center", wrap_text=False)
            if col_idx <= 2:
                c.fill = PatternFill("solid", fgColor="F2F2F2")
                c.font = Font(name="Arial", size=10, bold=(col_idx == 1))
            else:
                c.font = Font(name="Arial", size=10, color="1F4E79")
            if key in ("pct_complete", "remaining_dur") and value != "":
                c.alignment = Alignment(horizontal="right", vertical="center")

    # USERDATA sheet — do not modify this section, P6 is very particular about it
    wu  = wb.create_sheet("USERDATA")
    wu.column_dimensions["A"].width = 60
    b2  = Border(left=Side(style="thin", color="B0B8C8"), right=Side(style="thin", color="B0B8C8"),
                 top=Side(style="thin",  color="B0B8C8"), bottom=Side(style="thin", color="B0B8C8"))
    # Row 1: field key identifier (no spaces — required by P6)
    r1 = wu.cell(row=1, column=1, value="user_data")
    r1.font   = Font(name="Arial", bold=True, size=9, color="4472C4")
    r1.fill   = PatternFill("solid", fgColor="D9E1F2")
    r1.border = b2
    # Row 2: section label
    r2 = wu.cell(row=2, column=1, value="UserSettings Do Not Edit")
    r2.font   = Font(name="Arial", bold=True, size=11, color="1C3557")
    r2.fill   = PatternFill("solid", fgColor="D9E1F2")
    r2.border = b2
    # Row 3: settings values
    r3 = wu.cell(row=3, column=1,
                 value=USER_DATA)
    r3.font      = Font(name="Arial", size=10)
    r3.fill      = PatternFill("solid", fgColor="F8F9FB")
    r3.alignment = Alignment(vertical="top", wrap_text=True)
    r3.border    = b2
    wu.row_dimensions[3].height = 80

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# EXCEL IMPORT
# ══════════════════════════════════════════════════════════════════════════════

P6_KEY_MAP = {
    "task_code": "activity_id", "task_name": "activity_name",
    "status_code": "activity_status", "act_start_date": "actual_start",
    "act_end_date": "actual_finish", "complete_pct": "pct_complete",
    "remain_drtn_hr_cnt": "remaining_dur", "complete_pct_type": "complete_pct_type",
    "wbs_id": "wbs_id", "user_field_910": "comments_import",
}
HEADER_KEY_MAP = {
    "activity id": "activity_id", "activity name": "activity_name",
    "activity status": "activity_status", "actual start": "actual_start",
    "actual finish": "actual_finish", "duration % complete": "pct_complete",
    "remaining duration": "remaining_dur", "percent complete type": "complete_pct_type",
    "wbs code": "wbs_id", "comments": "comments_import",
}

def read_p6_excel(file_bytes: bytes) -> tuple:
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    warnings_list = []
    sheet_name = next((s for s in wb.sheetnames if s.upper() == "TASK"), None)
    if sheet_name is None:
        sheet_name = wb.sheetnames[0]
        warnings_list.append(f"No TASK sheet found — reading from '{sheet_name}' instead.")
    ws = wb[sheet_name]
    rows_iter = list(ws.iter_rows(values_only=True))
    if not rows_iter:
        return [], ["The sheet appears to be empty."]
    col_map, data_start = {}, 1
    for row_idx in range(min(3, len(rows_iter))):
        row = rows_iter[row_idx]
        mapping = {}
        for col_idx, cell_val in enumerate(row):
            if cell_val is None:
                continue
            cs = str(cell_val).strip().lower()
            if cs in P6_KEY_MAP:
                dk = P6_KEY_MAP[cs]
                if dk != "complete_pct_type":  # always force Physical on export; ignore on import
                    mapping[col_idx] = dk
            elif cs in HEADER_KEY_MAP:
                dk = HEADER_KEY_MAP[cs]
                if dk != "complete_pct_type":
                    mapping[col_idx] = dk
        if mapping:
            col_map, data_start = mapping, row_idx + 1
            if row_idx == 0 and len(rows_iter) > 1:
                next_str = [str(v).strip().lower() for v in rows_iter[1] if v is not None]
                if any(h in HEADER_KEY_MAP for h in next_str):
                    data_start = 2
            break
    if not col_map:
        return [], ["Could not detect column headers."]
    entries = []
    for row in rows_iter[data_start:]:
        if all(v is None or str(v).strip() == "" for v in row):
            continue
        entry = {
            "activity_id": "", "activity_name": "", "activity_status": "",
            "actual_start": "", "actual_finish": "", "pct_complete": "",
            "remaining_dur": "", "complete_pct_type": "Physical", "wbs_id": "",
            "comments_import": "",
        }
        for col_idx, data_key in col_map.items():
            if col_idx >= len(row):
                continue
            raw_val = row[col_idx]
            if data_key in DATE_KEYS:
                entry[data_key] = normalise_imported_date(raw_val)
            elif data_key == "pct_complete":
                vs = str(raw_val).replace("%", "").strip() if raw_val is not None else ""
                try:
                    entry[data_key] = str(int(float(vs))) if vs else ""
                except ValueError:
                    entry[data_key] = vs
            else:
                entry[data_key] = "" if raw_val is None else str(raw_val).strip()
        if not entry["activity_id"]:
            continue
        if not entry["complete_pct_type"]:
            entry["complete_pct_type"] = "Physical"
        entry["_submitted_at"] = datetime.now().strftime("%d/%m/%Y %H:%M")
        entries.append(entry)
    return entries, warnings_list


# ══════════════════════════════════════════════════════════════════════════════
# MICROSOFT PROJECT IMPORT
# ══════════════════════════════════════════════════════════════════════════════
# MS Project exports don't carry stable Activity IDs, so rows are matched to
# stored activities by (Name + WBS suffix).
# P6 prefixes WBS with the project name:  "ProjectX.1.2.3"
# MS Project stores WBS without prefix:   "1.2.3"
# We strip everything up to and including the first '.' before comparing.

MSP_KEY_MAP = {
    # MS Project XML/XLSX field headers (lowercase)
    "task name":          "activity_name",
    "name":               "activity_name",
    "wbs":                "wbs_id",
    "outline number":     "wbs_id",
    "% complete":         "pct_complete",
    "percent complete":   "pct_complete",
    "% work complete":    "pct_complete",
    "actual start":       "actual_start",
    "actual finish":      "actual_finish",
    "actual duration":    "remaining_dur",
    "remaining duration": "remaining_dur",
    "duration":           "remaining_dur",
    "status":             "activity_status",
    "notes":              "comments_import",
}

# MS Project status strings → our internal values
MSP_STATUS_MAP = {
    "complete":     "Completed",
    "completed":    "Completed",
    "in progress":  "In Progress",
    "not started":  "Not Started",
    "future task":  "Not Started",
    "on schedule":  "In Progress",
    "late":         "In Progress",
    "":             "Not Started",
}

def strip_wbs_prefix(wbs: str) -> str:
    """Strip the P6 project-name prefix from a stored WBS code.
    'ProjectX.1.2.3'  →  '1.2.3'
    '1.2.3'           →  '1.2.3'  (already clean, first segment is numeric)
    """
    if not wbs:
        return ""
    parts = wbs.strip().split(".", 1)
    if len(parts) == 2 and not parts[0].isdigit():
        return parts[1]
    return wbs.strip()


def strip_msp_wbs(wbs: str) -> str:
    """Normalise an MSP WBS code for comparison against stored P6 WBS.
    MSP has one extra level vs P6, and single-segment codes are WBS titles.
    '1.2.3.4'  →  '1.2.3'   (drop last segment)
    '1'        →  ''          (WBS title row — ignored)
    '1.2'      →  '1'
    """
    if not wbs:
        return ""
    segments = wbs.strip().split(".")
    if len(segments) <= 1:
        return ""
    return ".".join(segments[:-1])

def read_msp_excel(file_bytes: bytes) -> tuple:
    """
    Read a Microsoft Project XLSX export.
    Returns (rows, warnings) where rows use our internal field names.
    Rows do NOT have activity_id set — matching is done in the UI.
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    warnings_list = []
    # Use the first sheet (MSP exports vary in sheet naming)
    ws = wb[wb.sheetnames[0]]
    rows_iter = list(ws.iter_rows(values_only=True))
    if not rows_iter:
        return [], ["The sheet appears to be empty."]

    # Detect header row
    col_map, data_start = {}, 1
    for row_idx in range(min(5, len(rows_iter))):
        row = rows_iter[row_idx]
        mapping = {}
        for col_idx, cell_val in enumerate(row):
            if cell_val is None:
                continue
            cs = str(cell_val).strip().lower()
            if cs in MSP_KEY_MAP:
                mapping[col_idx] = MSP_KEY_MAP[cs]
        if mapping:
            col_map, data_start = mapping, row_idx + 1
            break

    if not col_map:
        return [], ["Could not detect Microsoft Project column headers. "
                    "Expected columns like 'Task Name', 'WBS', '% Complete', "
                    "'Actual Start', 'Actual Finish'."]

    if "activity_name" not in col_map.values():
        warnings_list.append("No Task Name / Name column found.")
    if "wbs_id" not in col_map.values():
        warnings_list.append("No WBS / Outline Number column found — name-only matching will be used.")

    entries = []
    for row in rows_iter[data_start:]:
        if all(v is None or str(v).strip() == "" for v in row):
            continue
        entry = {
            "activity_id":       "",   # filled by matching logic in the UI
            "activity_name":     "",
            "activity_status":   "",
            "actual_start":      "",
            "actual_finish":     "",
            "pct_complete":      "",
            "remaining_dur":     "",
            "complete_pct_type": "Physical",
            "wbs_id":            "",
            "comments_import":   "",
        }
        for col_idx, data_key in col_map.items():
            if col_idx >= len(row):
                continue
            raw_val = row[col_idx]
            if data_key in DATE_KEYS:
                entry[data_key] = normalise_imported_date(raw_val)
            elif data_key == "pct_complete":
                vs = str(raw_val).replace("%", "").strip() if raw_val is not None else ""
                try:
                    entry[data_key] = str(int(float(vs))) if vs else ""
                except ValueError:
                    entry[data_key] = vs
            elif data_key == "activity_status":
                raw_str = str(raw_val).strip().lower() if raw_val else ""
                entry[data_key] = MSP_STATUS_MAP.get(raw_str, "Not Started")
            elif data_key == "remaining_dur":
                # MSP duration strings like "5 days", "5d", "5" — extract number
                if raw_val is None:
                    entry[data_key] = ""
                else:
                    dur_str = str(raw_val).strip()
                    import re as _re
                    m = _re.search(r"(\d+\.?\d*)", dur_str)
                    entry[data_key] = m.group(1) if m else ""
            else:
                entry[data_key] = "" if raw_val is None else str(raw_val).strip()

        # Skip summary rows (WBS with no sub-level or name only rows)
        if not entry["activity_name"]:
            continue

        # Strip P6 prefix from WBS for consistent comparison
        entry["wbs_id"] = strip_msp_wbs(entry["wbs_id"])
        entry["_submitted_at"] = datetime.now().strftime("%d/%m/%Y %H:%M")
        entries.append(entry)

    return entries, warnings_list


def match_msp_to_stored(msp_rows: list[dict], stored: list[dict]) -> tuple:
    """
    Match MSP rows to stored activities by (name + WBS suffix).
    Returns:
      matched       — list of (msp_row, stored_entry)   — exactly one match
      unmatched     — list of msp_row                    — no stored match found
      duplicates    — list of (msp_row, [stored_entries]) — ambiguous (2+ matches)
    """
    # Build lookup: (name.lower(), wbs_suffix.lower()) → [stored entries]
    lookup: dict[tuple, list] = {}
    for e in stored:
        key = (
            e.get("activity_name", "").strip().lower(),
            strip_wbs_prefix(e.get("wbs_id", "")).lower(),
        )
        lookup.setdefault(key, []).append(e)

    matched, unmatched, duplicates = [], [], []
    for row in msp_rows:
        key = (
            row.get("activity_name", "").strip().lower(),
            row.get("wbs_id", "").lower(),
        )
        hits = lookup.get(key, [])
        if len(hits) == 1:
            matched.append((row, hits[0]))
        elif len(hits) == 0:
            unmatched.append(row)
        else:
            duplicates.append((row, hits))

    return matched, unmatched, duplicates

# ══════════════════════════════════════════════════════════════════════════════
# WBS OFFSET DETECTION
# ══════════════════════════════════════════════════════════════════════════════

def detect_wbs_offset(unmatched: list[dict], stored: list[dict]) -> list[dict]:
    """
    For each unmatched MSP row, check whether the same activity name exists in
    stored with a WBS that differs only in the last numeric segment by ±1.

    Returns a list of offset suggestions, each a dict:
      {
        "msp_row":      the unmatched MSP row,
        "stored_entry": the stored entry found at the adjacent WBS,
        "msp_wbs":      e.g. "1.2.3"
        "stored_wbs":   e.g. "1.2.2"   (stored WBS after stripping prefix)
        "depth":        the segment index where they differ (0-based),
        "delta":        +1 or -1  (MSP value minus stored value at that depth),
      }

    Only suggests when the name match is unique and unambiguous.
    """
    # Build name → list of stored entries
    name_lookup: dict[str, list] = {}
    for e in stored:
        name_lookup.setdefault(e.get("activity_name","").strip().lower(), []).append(e)

    suggestions = []
    for row in unmatched:
        msp_name = row.get("activity_name","").strip().lower()
        msp_wbs  = row.get("wbs_id","").strip()
        if not msp_wbs:
            continue
        msp_segs = msp_wbs.split(".")
        if len(msp_segs) < 2:
            continue

        candidates = name_lookup.get(msp_name, [])
        if not candidates:
            continue

        for delta in (+1, -1):
            for depth in range(len(msp_segs) - 1, -1, -1):
                try:
                    adj_val = int(msp_segs[depth]) - delta  # stored value = msp - delta
                    if adj_val < 0:
                        continue
                except ValueError:
                    continue

                adj_segs = msp_segs[:]
                adj_segs[depth] = str(adj_val)
                adj_wbs = ".".join(adj_segs)

                # Compare against stored WBS after stripping P6 prefix
                for e in candidates:
                    stored_wbs_clean = strip_wbs_prefix(e.get("wbs_id","")).lower()
                    if stored_wbs_clean == adj_wbs.lower():
                        suggestions.append({
                            "msp_row":      row,
                            "stored_entry": e,
                            "msp_wbs":      msp_wbs,
                            "stored_wbs":   stored_wbs_clean,
                            "depth":        depth,
                            "delta":        delta,
                        })

    return suggestions


def apply_wbs_offset(stored: list[dict], prefix: str, depth: int,
                     delta: int, from_val: int) -> tuple[list[dict], int]:
    """
    Shift WBS codes in stored entries where:
      - The WBS prefix up to `depth` matches `prefix`
      - The segment at `depth` is >= from_val  (for positive delta)
        or <= from_val  (for negative delta)

    `prefix` is the dot-joined segments BEFORE depth, e.g. "1.2" for depth 2.
    `delta`  is the amount to add to segment at depth (+1 or -1).
    `from_val` is the stored segment value where the shift starts.

    Returns (updated_entries, count_changed).
    """
    changed = 0
    for entry in stored:
        raw_wbs   = entry.get("wbs_id", "")
        clean_wbs = strip_wbs_prefix(raw_wbs)
        segs      = clean_wbs.split(".")
        if len(segs) <= depth:
            continue
        # Check prefix matches
        if prefix and ".".join(segs[:depth]) != prefix:
            continue
        try:
            seg_val = int(segs[depth])
        except ValueError:
            continue
        # Only shift entries at or beyond the insertion/deletion point
        if delta > 0 and seg_val < from_val:
            continue
        if delta < 0 and seg_val > from_val:
            continue

        new_segs       = segs[:]
        new_segs[depth] = str(seg_val + delta)
        new_clean_wbs  = ".".join(new_segs)

        # Reconstruct: if original had a prefix, re-attach it
        original_segs = raw_wbs.split(".")
        if len(original_segs) > len(segs):
            # There was a non-numeric prefix segment
            prefix_part  = ".".join(original_segs[:len(original_segs)-len(segs)])
            entry["wbs_id"] = prefix_part + "." + new_clean_wbs
        else:
            entry["wbs_id"] = new_clean_wbs
        changed += 1

    return stored, changed


# ══════════════════════════════════════════════════════════════════════════════
# DATE INPUT WIDGET HELPER
# Uses st.datetime_input (Streamlit >= 1.43).
# Returns a datetime, or None if the user has not enabled an optional field.
# ══════════════════════════════════════════════════════════════════════════════

def datetime_inputs(label: str, key: str, required: bool = True,
                    default_dt: datetime | None = None) -> datetime | None:
    """
    Render a single datetime picker (st.datetime_input).
    For optional fields a checkbox gates the widget; returns None when unchecked.
    """
    default_val = default_dt if default_dt else datetime.combine(date.today(), time(8, 0))

    if not required:
        enabled = st.checkbox(f"Set {label}", key=f"{key}_enabled",
                              value=(default_dt is not None))
        if not enabled:
            return None

    return st.datetime_input(label, value=default_val, key=f"{key}_dt",
                              step=60 * 15)   # 15-minute steps

# ══════════════════════════════════════════════════════════════════════════════
# PAGE SETUP
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(page_title="P6 Asbuilt Collector", page_icon="🏗️", layout="wide")

# ══════════════════════════════════════════════════════════════════════════════
# AUTH — SESSION STATE & LOGIN SCREEN
# ══════════════════════════════════════════════════════════════════════════════

if "authenticated" not in st.session_state:
    st.session_state.update({"authenticated": False, "username": "",
                              "display_name": "", "role": ""})

if not st.session_state.authenticated:
    st.title("P6 Asbuilt Collector")
    st.caption("Sign in to continue")
    st.divider()

    _, col_m, _ = st.columns([1, 1.1, 1])
    with col_m:
        with st.container(border=True):
            st.subheader("Sign In")
            username = st.text_input("Username", placeholder="Enter your username")
            password = st.text_input("Password", type="password", placeholder="Enter your password")
            if st.button("Log In", type="primary", use_container_width=True):
                user = USERS.get(username)
                if user and user["hash"] == hashlib.sha256(password.encode()).hexdigest():
                    st.session_state.update({
                        "authenticated": True, "username": username,
                        "display_name": user["name"], "role": user["role"],
                    })
                    st.rerun()
                else:
                    st.error("Incorrect username or password.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR (authenticated)
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.title("🏗️ P6 Asbuilt")
    st.divider()
    st.write(f"**{st.session_state.display_name}**")
    st.caption(ROLE_LABEL.get(st.session_state.role, st.session_state.role))
    if st.button("Log Out", use_container_width=True):
        st.session_state.update({"authenticated": False, "username": "",
                                  "display_name": "", "role": ""})
        st.rerun()
    st.divider()
    st.caption("Entries: p6_asbuilt_store.json")
    st.caption("Photos:  p6_images/")

# ══════════════════════════════════════════════════════════════════════════════
# HEADER & DYNAMIC TABS
# ══════════════════════════════════════════════════════════════════════════════

st.title("🏗️ Primavera P6 — Asbuilt Data Collector")
st.caption("Submit and update asbuilt progress entries, then export a P6-compatible spreadsheet.")
st.divider()
logo=Path("Tricertus_logo.jpg")
st.logo(logo,size="large")

TAB_DEFS = [
    ("📋  View All Entries",  "view"),
    ("📝  Submit / Update",   "submit"),
    ("📤  Import from Excel", "import"),
    ("📥  Export to Excel",   "export"),
    ("📸  Photo Log",         "photos"),
]
visible   = [(lbl, perm) for lbl, perm in TAB_DEFS if has_permission(perm)]
tab_objs  = st.tabs([lbl for lbl, _ in visible])
tab_index = {perm: tab_objs[i] for i, (_, perm) in enumerate(visible)}

# ══════════════════════════════════════════════════════════════════════════════
# TAB: VIEW ALL ENTRIES
# ══════════════════════════════════════════════════════════════════════════════

with tab_index["view"]:
    entries = load_entries()
    st.subheader(f"All Entries ({len(entries)})")

    if not entries:
        st.info("No entries yet.")
    else:
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total",       len(entries))
        m2.metric("Completed",   sum(1 for e in entries if e.get("activity_status") == "Completed"))
        m3.metric("In Progress", sum(1 for e in entries if e.get("activity_status") == "In Progress"))
        m4.metric("Not Started", sum(1 for e in entries if e.get("activity_status") == "Not Started"))
        st.divider()

        # ── Search + Sort controls ─────────────────────────────────────────
        search_col, sort_col, dir_col = st.columns([3, 2, 1])
        with search_col:
            search_text = st.text_input(
                "Search",
                placeholder="Activity ID or name…",
                key="view_search",
            ).strip().lower()
        with sort_col:
            sort_by = st.selectbox(
                "Sort by",
                options=["WBS Code", "Actual Start", "Actual Finish", "Activity ID"],
                key="view_sort_by",
            )
        with dir_col:
            sort_asc = st.radio(
                "Order", options=["↑ Asc", "↓ Desc"],
                key="view_sort_dir", horizontal=True,
            ) == "↑ Asc"

        # Apply search filter before sorting
        if search_text:
            entries = [
                e for e in entries
                if search_text in e.get("activity_id",   "").lower()
                or search_text in e.get("activity_name", "").lower()
                or search_text in e.get("wbs_id",        "").lower()
            ]

        def wbs_key(e: dict):
            parts = e.get("wbs_id", "").split(".")
            segments = []
            for p in parts:
                try:
                    segments.append((0, int(p)))
                except ValueError:
                    segments.append((1, p.lower()))
            return segments or [(1, "")]

        def date_key(field: str):
            def _key(e: dict):
                dt = iso_to_dt(e.get(field, ""))
                return dt if dt else (datetime.min if sort_asc else datetime.max)
            return _key

        if sort_by == "WBS Code":
            sorted_entries = sorted(entries, key=wbs_key, reverse=not sort_asc)
        elif sort_by == "Actual Start":
            sorted_entries = sorted(entries, key=date_key("actual_start"), reverse=not sort_asc)
        elif sort_by == "Actual Finish":
            sorted_entries = sorted(entries, key=date_key("actual_finish"), reverse=not sort_asc)
        else:
            sorted_entries = sorted(entries, key=lambda e: e.get("activity_id", "").upper(), reverse=not sort_asc)

        st.divider()

        can_edit   = has_permission("submit")
        can_delete = has_permission("submit")

        # ── Pagination ─────────────────────────────────────────────────────
        PAGE_SIZE = 25
        total_pages = max(1, (len(sorted_entries) + PAGE_SIZE - 1) // PAGE_SIZE)
        page = st.number_input(
            f"Page (1 – {total_pages})",
            min_value=1, max_value=total_pages, value=1, step=1,
            key="view_page",
        ) - 1  # zero-based
        page_entries = sorted_entries[page * PAGE_SIZE : (page + 1) * PAGE_SIZE]
        total_label  = f"{len(sorted_entries)} match{'es' if len(sorted_entries) != 1 else ''}" if search_text else f"{len(sorted_entries)} entries"
        st.caption(
            f"Showing {page * PAGE_SIZE + 1}–{min((page + 1) * PAGE_SIZE, len(sorted_entries))} "
            f"of {total_label}"
        )
        st.divider()

        # Build a lookup so we don't scan the list for every card
        id_to_index = {e.get("activity_id", "").upper(): idx for idx, e in enumerate(entries)}

        for entry in page_entries:
            i     = id_to_index.get(entry.get("activity_id", "").upper(), 0)
            status    = entry.get("activity_status", "")
            act_id    = entry.get("activity_id", "")
            act_name  = entry.get("activity_name", "")
            wbs       = entry.get("wbs_id", "—")
            pct       = entry.get("pct_complete", "0")
            rem       = entry.get("remaining_dur", "—")
            a_start   = display_dt(entry.get("actual_start", ""))
            a_finish  = display_dt(entry.get("actual_finish", ""))
            subm_at   = entry.get("_submitted_at", "")
            submitter = entry.get("_submitted_by", "")
            n_comments = len(entry.get("_comments", []))

            with st.container(border=True):
                # ── Header row ─────────────────────────────────────────────
                head_left, head_left2, head_right = st.columns([2, 2, 1])
                with head_left:
                    st.write(f"Activity ID: {act_id}")
                with head_left2:
                    st.write(f"Activity Name: {act_name}")
                with head_right:
                    st.write(status)

                # ── Detail row ─────────────────────────────────────────────
                c1, c2, c3, c4, c5 = st.columns([1, 1, 1, 0.3, 0.4], gap="xsmall")
                c1.write("WBS");      c1.write(wbs)
                c2.write("Start");    c2.write(a_start)
                c3.write("Finish");   c3.write(a_finish)
                c4.metric("% Complete", f"{pct}%")
                c5.metric("Remaining", f"{rem} days" if rem and rem != "—" else "—")

                # ── Footer ─────────────────────────────────────────────────
                footer = f"Last updated: {subm_at}"
                if submitter:
                    footer += f"  ·  By: {submitter}"
                if n_comments:
                    footer += f"  ·  💬 {n_comments} comment{'s' if n_comments != 1 else ''}"
                st.caption(footer)

                # ── Inline edit expander (submit/admin only) ───────────────
                if can_edit:
                    with st.expander("✏️  Edit name / Add comment"):
                        # Activity name
                        st.write("**Edit Activity Name**")
                        new_name = st.text_input(
                            "Activity Name",
                            value=act_name,
                            key=f"edit_name_{i}",
                            label_visibility="collapsed",
                        ).strip()

                        st.divider()

                        # Existing comments
                        st.write("**Comments**")
                        existing_comments = entry.get("_comments", [])
                        if existing_comments:
                            for c in existing_comments:
                                st.write(f"**{c['at']}** — {c['by']}")
                                st.write(c["text"])
                                st.divider()
                        else:
                            st.caption("No comments yet.")

                        new_comment_text = st.text_area(
                            "Add comment",
                            placeholder="Enter progress notes, observations, or issues...",
                            height=100,
                            key=f"view_comment_{i}",
                            label_visibility="collapsed",
                        ).strip()

                        # Save button — only active if something changed
                        name_changed    = new_name != act_name and new_name != ""
                        comment_entered = bool(new_comment_text)

                        if st.button(
                            "💾  Save changes",
                            key=f"edit_save_{i}",
                            type="primary",
                            disabled=not (name_changed or comment_entered),
                        ):
                            updated = entry.copy()
                            if name_changed:
                                updated["activity_name"] = new_name
                            if comment_entered:
                                new_record = {
                                    "text": new_comment_text,
                                    "by":   st.session_state.display_name,
                                    "at":   datetime.now().strftime("%d/%m/%Y %H:%M"),
                                }
                                updated["_comments"] = [new_record] + existing_comments
                            updated["_submitted_at"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                            updated["_submitted_by"] = st.session_state.display_name
                            entries[i] = updated
                            save_entries(entries)
                            st.success("Saved.")
                            st.rerun()

                # ── Delete button ──────────────────────────────────────────
                if can_delete and st.button(f"🗑 Delete {act_id}", key=f"del_{i}"):
                    entries.pop(i)
                    save_entries(entries)
                    st.rerun()

        st.dataframe(sorted_entries)

# ══════════════════════════════════════════════════════════════════════════════
# TAB: SUBMIT / UPDATE
# ══════════════════════════════════════════════════════════════════════════════

if "submit" in tab_index:
    with tab_index["submit"]:
        entries   = load_entries()
        known_ids = {e["activity_id"].upper(): e for e in entries}

        st.subheader("Submit or Update an Asbuilt Entry")
        st.caption("Enter the Activity ID — if it already exists the name is filled automatically.")

        col_id, col_wbs = st.columns(2)
        with col_id:
            activity_id_raw = st.text_input("Activity ID *", placeholder="e.g. A1000").strip()
        with col_wbs:
            wbs_input = st.text_input("WBS ID *", placeholder="e.g. 1.2.3").strip()

        existing = known_ids.get(activity_id_raw.upper()) if activity_id_raw else None
        if existing:
            st.info(
                f"**Existing entry found:** {existing['activity_name']}  \n"
                f"Status: **{existing['activity_status']}** | "
                f"**{existing.get('pct_complete', 0)}%** complete  \n"
                f"Submitting will **update** this entry.", icon="ℹ️",
            )

        if existing:
            st.text_input("Activity Name", value=existing["activity_name"], disabled=True)
            activity_name = existing["activity_name"]
        else:
            activity_name = st.text_input(
                "Activity Name *", placeholder="e.g. Concrete Pour - Foundations"
            ).strip()

        activity_status = st.selectbox("Activity Status *", STATUS_OPTIONS)

        def _existing_dt(key: str) -> datetime | None:
            return iso_to_dt(existing.get(key, "")) if existing else None

        actual_start_dt  = None
        actual_finish_dt = None
        pct_complete     = 0
        remaining_dur    = ""

        if activity_status == "Not Started":
            st.info(
                "% Complete is set to 0 and Remaining Duration left blank automatically "
                "for 'Not Started' activities.", icon="ℹ️"
            )

        elif activity_status == "In Progress":
            actual_start_dt = datetime_inputs(
                "Actual Start *", key="start_ip", required=True,
                default_dt=_existing_dt("actual_start"),
            )
            col_p, col_r = st.columns(2)
            with col_p:
                pct_complete = st.number_input(
                    "Duration % Complete *", min_value=0, max_value=99, step=5,
                    value=int(existing.get("pct_complete", 0)) if existing else 0,
                )
            with col_r:
                remaining_dur = st.text_input(
                    "Remaining Duration (days) *", placeholder="e.g. 5",
                    value=existing.get("remaining_dur", "") if existing else "",
                ).strip()

        elif activity_status == "Completed":
            actual_start_dt = datetime_inputs(
                "Actual Start *", key="start_c", required=True,
                default_dt=_existing_dt("actual_start"),
            )
            actual_finish_dt = datetime_inputs(
                "Actual Finish *", key="finish_c", required=True,
                default_dt=_existing_dt("actual_finish"),
            )
            pct_complete  = 100
            remaining_dur = "0"
            st.info("% Complete set to 100 and Remaining Duration to 0 automatically.", icon="✅")

        # ── Comments section ──────────────────────────────────────────────
        st.divider()
        st.subheader("Comments")

        existing_comments = existing.get("_comments", []) if existing else []
        if existing_comments:
            st.caption(f"{len(existing_comments)} existing comment{'s' if len(existing_comments) != 1 else ''} stored for this activity:")
            for c in existing_comments:
                st.write(f"**{c['at']}** — {c['by']}")
                st.write(c["text"])
                st.divider()

        new_comment_text = st.text_area(
            "Add a new comment (optional)",
            placeholder="Enter progress notes, observations, or issues...",
            height=120,
            key="submit_new_comment",
        ).strip()

        if st.button("Submit Entry", type="primary"):
            errors = []
            if not activity_id_raw:                                                      errors.append("Activity ID is required.")
            if not wbs_input:                                                            errors.append("WBS ID is required.")
            if not existing and not activity_name:                                       errors.append("Activity Name is required for new activities.")
            if activity_status in ("In Progress", "Completed") and not actual_start_dt: errors.append("Actual Start is required.")
            if activity_status == "Completed" and not actual_finish_dt:                  errors.append("Actual Finish is required when status is Completed.")
            if activity_status == "In Progress" and not remaining_dur:                   errors.append("Remaining Duration is required when In Progress.")
            if actual_start_dt and actual_finish_dt and actual_finish_dt < actual_start_dt:
                errors.append("Actual Finish cannot be before Actual Start.")

            if errors:
                for e in errors:
                    st.error(e)
            else:
                # Build updated comment list — prepend new comment if provided
                updated_comments = list(existing.get("_comments", [])) if existing else []
                if new_comment_text:
                    updated_comments.insert(0, {
                        "text": new_comment_text,
                        "by":   st.session_state.display_name,
                        "at":   datetime.now().strftime("%d/%m/%Y %H:%M"),
                    })

                entry = {
                    "activity_id":       activity_id_raw,
                    "activity_name":     activity_name,
                    "activity_status":   activity_status,
                    "actual_start":      dt_to_iso(actual_start_dt)  if actual_start_dt  else "",
                    "actual_finish":     dt_to_iso(actual_finish_dt) if actual_finish_dt else "",
                    "pct_complete":      str(pct_complete),
                    "remaining_dur":     str(remaining_dur),
                    "complete_pct_type": "Physical",
                    "wbs_id":            wbs_input,
                    "_comments":         updated_comments,
                    "_submitted_at":     datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "_submitted_by":     st.session_state.display_name,
                }
                entries, action = upsert_entry(entries, entry)
                save_entries(entries)
                icon = "✅" if action == "saved" else "🔄"
                st.success(f"{icon} Entry **{action}** successfully!")
                with st.expander("View saved data"):
                    display = entry.copy()
                    if display["actual_start"]:  display["actual_start"]  = display_dt(display["actual_start"])
                    if display["actual_finish"]: display["actual_finish"] = display_dt(display["actual_finish"])
                    st.json(display)

# ══════════════════════════════════════════════════════════════════════════════
# TAB: IMPORT FROM EXCEL
# ══════════════════════════════════════════════════════════════════════════════

if "import" in tab_index:
    with tab_index["import"]:
        entries   = load_entries()
        known_ids = {e["activity_id"].upper(): e for e in entries}

        st.subheader("Import from Excel")

        import_mode = st.radio(
            "Source format",
            options=["Primavera P6 Export", "Microsoft Project Export"],
            horizontal=True,
            key="import_mode",
        )

        if import_mode == "Primavera P6 Export":
            st.caption("Upload a P6-format XLSX (TASK sheet). Conflicts matched by Activity ID.")
        else:
            st.caption(
                "Upload a Microsoft Project XLSX export. "
                "Activities are matched to stored entries by **Name + WBS**. "
                "MS Project WBS codes are compared after stripping the P6 project prefix."
            )

        uploaded = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"], key="import_file")

        if uploaded:
            file_bytes = uploaded.read()
            with st.spinner("Reading Excel file..."):
                if import_mode == "Primavera P6 Export":
                    imported_rows, warnings_list = read_p6_excel(file_bytes)
                else:
                    imported_rows, warnings_list = read_msp_excel(file_bytes)

            for w in warnings_list:
                st.warning(w)

            if not imported_rows:
                st.error("No valid rows found. Check the file format and column headers.")

            elif import_mode == "Microsoft Project Export":
                # ── MSP matching flow ──────────────────────────────────────
                st.divider()
                matched_all, unmatched, duplicates = match_msp_to_stored(imported_rows, entries)

                # Drop matched rows where the incoming data is identical to stored
                matched          = [(r, s) for r, s in matched_all
                                    if not is_exact_duplicate(r, s)]
                identical_skipped = len(matched_all) - len(matched)

                msg = (
                    f"Found **{len(imported_rows)}** rows: "
                    f"**{len(matched)}** matched with changes, "
                    f"**{len(unmatched)}** unmatched, "
                    f"**{len(duplicates)}** ambiguous"
                )
                if identical_skipped:
                    msg += f", **{identical_skipped}** identical (skipped)"
                st.success(msg)

                # ── Duplicate name+WBS conflicts ──────────────────────────
                if duplicates:
                    st.divider()
                    st.subheader("⚠️ Ambiguous matches — multiple activities share the same Name and WBS")
                    st.caption(
                        "The activities below could not be matched uniquely. "
                        "Rename one of the stored activities to make it unique, "
                        "then re-import."
                    )
                    for msp_row, hits in duplicates:
                        with st.expander(
                            f"❓  {msp_row['activity_name']}  —  WBS: {msp_row['wbs_id']}",
                            expanded=True,
                        ):
                            st.write(f"**{len(hits)} stored activities match this name and WBS:**")
                            for h in hits:
                                st.write(
                                    f"- `{h['activity_id']}`  {h['activity_name']}  "
                                    f"WBS: {h['wbs_id']}  Status: {h['activity_status']}"
                                )
                            st.warning(
                                "Rename one of these activities in the Submit/Update tab "
                                "or View All Entries before re-importing."
                            )

                # ── WBS offset detection ─────────────────────────────────
                # Run before rendering unmatched cards so suggestions can be
                # applied to stored entries before the user handles each row.
                wbs_applied_offsets = {}  # suggestion_key → bool (applied)

                if unmatched:
                    offset_suggestions = detect_wbs_offset(unmatched, entries)

                    # Deduplicate suggestions by (depth, delta, prefix, from_val)
                    # so we only show one prompt per unique shift
                    seen_shifts  = {}   # shift_key → suggestion
                    for sg in offset_suggestions:
                        segs     = sg["stored_wbs"].split(".")
                        prefix   = ".".join(segs[:sg["depth"]])
                        from_val = int(segs[sg["depth"]])
                        shift_key = (prefix, sg["depth"], sg["delta"], from_val)
                        if shift_key not in seen_shifts:
                            seen_shifts[shift_key] = sg

                    if seen_shifts:
                        st.divider()
                        st.subheader("🔀 Possible WBS offset detected")
                        st.caption(
                            "The following unmatched activities were found at adjacent WBS codes "
                            "with the same name — this may indicate a WBS level was inserted or "
                            "deleted, shifting all subsequent codes. Review each suggestion and "
                            "apply the shift if correct. **Applying a shift updates the stored "
                            "WBS codes immediately and cannot be undone here — save a backup first.**"
                        )

                        for shift_key, sg in seen_shifts.items():
                            prefix, depth, delta, from_val = shift_key
                            direction = "up (+1)" if delta > 0 else "down (−1)"
                            parent    = prefix if prefix else "root"
                            affected  = [
                                e for e in entries
                                if (lambda c, s, d, fv, dl:
                                    len(s) > d
                                    and (not c or ".".join(s[:d]) == c)
                                    and (int(s[d]) >= fv if dl > 0 else int(s[d]) <= fv)
                                )(
                                    prefix,
                                    strip_wbs_prefix(e.get("wbs_id","")).split("."),
                                    depth,
                                    from_val,
                                    delta,
                                )
                            ]

                            with st.expander(
                                f"Shift WBS segment {depth+1} {direction} "
                                f"under '{parent}' from position {from_val} onwards "
                                f"— affects ~{len(affected)} activities",
                                expanded=True,
                            ):
                                c_msp, c_stored = st.columns(2)
                                with c_msp:
                                    st.write("**Unmatched MSP activity:**")
                                    st.write(f"- Name: {sg['msp_row'].get('activity_name','')}")
                                    st.write(f"- WBS: `{sg['msp_wbs']}`")
                                with c_stored:
                                    st.write("**Stored activity found at adjacent WBS:**")
                                    st.write(f"- Name: {sg['stored_entry'].get('activity_name','')}")
                                    st.write(f"- Stored WBS: `{sg['stored_entry'].get('wbs_id','')}`")
                                    st.write(f"- ID: `{sg['stored_entry'].get('activity_id','')}`")

                                st.write(f"**Activities that would be shifted ({len(affected)}):**")
                                preview = affected[:5]
                                for e in preview:
                                    old_w = strip_wbs_prefix(e.get("wbs_id",""))
                                    segs2 = old_w.split(".")
                                    segs2[depth] = str(int(segs2[depth]) + delta)
                                    st.caption(
                                        f"- `{e['activity_id']}` {e['activity_name']}  "
                                        f"{old_w} → {'.'.join(segs2)}"
                                    )
                                if len(affected) > 5:
                                    st.caption(f"  … and {len(affected)-5} more")

                                if st.button(
                                    f"✅ Apply this WBS shift",
                                    key=f"apply_wbs_shift_{'_'.join(str(x) for x in shift_key)}",
                                    type="primary",
                                ):
                                    entries, n_changed = apply_wbs_offset(
                                        entries, prefix, depth, delta, from_val
                                    )
                                    save_entries(entries)
                                    # Rebuild known_ids after the shift
                                    known_ids = {e["activity_id"].upper(): e for e in entries}
                                    st.success(
                                        f"WBS shift applied — {n_changed} activities updated. "
                                        f"Re-run the import to see updated matches."
                                    )
                                    st.rerun()

                # ── Unmatched rows — add as new or manually map to existing ─
                # new_activity_ids:   {row_idx → new Activity ID string}
                # manual_overwrites:  {row_idx → stored activity_id to overwrite}
                new_activity_ids  = {}
                manual_overwrites = {}

                if unmatched:
                    st.divider()
                    st.subheader(f"➕ {len(unmatched)} unmatched rows")
                    st.caption(
                        "These rows had no automatic Name + WBS match. "
                        "For each row you can: **Add as new** (enter a new Activity ID), "
                        "**Overwrite existing** (manually select the stored activity it should update), "
                        "or leave both blank to skip."
                    )

                    # Build a readable option list for the overwrite selectbox
                    overwrite_options = ["— Select stored activity —"] + [
                        f"{e['activity_id']}  —  {e['activity_name']}  (WBS: {e['wbs_id']})"
                        for e in entries
                    ]
                    # Map display string back to activity_id
                    option_to_aid = {
                        f"{e['activity_id']}  —  {e['activity_name']}  (WBS: {e['wbs_id']})": e["activity_id"]
                        for e in entries
                    }

                    for idx, row in enumerate(unmatched):
                        with st.container(border=True):
                            # Row info
                            st.write(f"**{row['activity_name']}**")
                            st.caption(
                                f"WBS: {row['wbs_id'] or '—'}"
                                f"  ·  Status: {row.get('activity_status','—')}"
                                f"  ·  % Complete: {row.get('pct_complete','—')}"
                                f"  ·  Start: {display_dt(row.get('actual_start',''))}"
                                f"  ·  Finish: {display_dt(row.get('actual_finish',''))}"
                            )

                            # Mode toggle
                            mode = st.radio(
                                "Action",
                                ["Skip", "Add as new", "Overwrite existing"],
                                key=f"unmatched_mode_{idx}",
                                horizontal=True,
                            )

                            if mode == "Add as new":
                                c_id, c_status = st.columns([2, 1])
                                with c_id:
                                    proposed_id = st.text_input(
                                        "New Activity ID",
                                        placeholder="e.g. A1050",
                                        key=f"new_act_id_{idx}",
                                    ).strip()
                                with c_status:
                                    st.write("")  # spacer
                                    if proposed_id:
                                        if proposed_id.upper() in known_ids:
                                            st.error("ID already exists.")
                                        elif proposed_id.upper() in {
                                            v.upper() for v in new_activity_ids.values() if v
                                        }:
                                            st.error("ID used twice above.")
                                        else:
                                            st.success("✓")
                                            new_activity_ids[idx] = proposed_id

                            elif mode == "Overwrite existing":
                                st.caption(
                                    "Use this when the MSP name differs slightly from the stored name "
                                    "(e.g. 'ActivityName (detail)' vs 'ActivityName'). "
                                    "Only progress fields will be updated — the stored name and ID are kept."
                                )
                                selected_opt = st.selectbox(
                                    "Select stored activity to overwrite",
                                    options=overwrite_options,
                                    key=f"overwrite_select_{idx}",
                                )
                                if selected_opt != "— Select stored activity —":
                                    target_aid = option_to_aid[selected_opt]
                                    # Warn if this target is already being overwritten by another row
                                    already_used = [
                                        i for i, a in manual_overwrites.items()
                                        if a.upper() == target_aid.upper() and i != idx
                                    ]
                                    if already_used:
                                        st.error(
                                            f"This activity is already targeted by another row above."
                                        )
                                    else:
                                        manual_overwrites[idx] = target_aid
                                        stored_e = known_ids[target_aid.upper()]
                                        c_cur, c_new = st.columns(2)
                                        with c_cur:
                                            st.write("**Stored (will be kept):**")
                                            st.write(f"- Name: {stored_e.get('activity_name','')}")
                                            st.write(f"- Status: {stored_e.get('activity_status','—')}")
                                            st.write(f"- Start: {display_dt(stored_e.get('actual_start',''))}")
                                            st.write(f"- Finish: {display_dt(stored_e.get('actual_finish',''))}")
                                            st.write(f"- % Complete: {stored_e.get('pct_complete','—')}")
                                        with c_new:
                                            st.write("**Incoming (will overwrite progress):**")
                                            st.write(f"- Name: {row.get('activity_name','')}")
                                            st.write(f"- Status: {row.get('activity_status','—')}")
                                            st.write(f"- Start: {display_dt(row.get('actual_start',''))}")
                                            st.write(f"- Finish: {display_dt(row.get('actual_finish',''))}")
                                            st.write(f"- % Complete: {row.get('pct_complete','—')}")

                # ── Matched rows resolution ───────────────────────────────
                msp_resolutions = {}
                if matched:
                    st.divider()
                    st.subheader(f"✅ {len(matched)} matched rows — review and confirm")
                    st.caption("Each MSP row has been matched to a stored activity. Review the changes below.")

                    fields = [
                        ("Status",          "activity_status", False),
                        ("Actual Start",     "actual_start",    True),
                        ("Actual Finish",    "actual_finish",   True),
                        ("% Complete",       "pct_complete",    False),
                        ("Remaining (days)", "remaining_dur",   False),
                    ]
                    for msp_row, stored_entry in matched:
                        aid = stored_entry["activity_id"].upper()
                        label = (
                            f"🔁  `{stored_entry['activity_id']}`  "
                            f"{stored_entry['activity_name']}  —  WBS: {stored_entry['wbs_id']}"
                        )
                        with st.expander(label, expanded=False):
                            c_cur, c_new = st.columns(2)
                            with c_cur:
                                st.write("**Current (stored)**")
                                for lbl, k, is_date in fields:
                                    val = stored_entry.get(k, "") or ""
                                    st.write(f"- **{lbl}:** {display_dt(val) if is_date else (val or '—')}")
                            with c_new:
                                st.write("**Incoming (MS Project)**")
                                for lbl, k, is_date in fields:
                                    val = msp_row.get(k, "") or ""
                                    st.write(f"- **{lbl}:** {display_dt(val) if is_date else (val or '—')}")

                            choice = st.radio(
                                "Action",
                                ["Overwrite with incoming", "Keep current"],
                                key=f"msp_conflict_{aid}",
                                horizontal=True,
                            )
                            msp_resolutions[aid] = {
                                "action":  "overwrite" if choice == "Overwrite with incoming" else "skip",
                                "msp_row": msp_row,
                                "stored":  stored_entry,
                            }

                    st.divider()
                    ow = sum(1 for v in msp_resolutions.values() if v["action"] == "overwrite")
                    sk = sum(1 for v in msp_resolutions.values() if v["action"] == "skip")
                    parts = []
                    if ow: parts.append(f"{ow} activities will be updated")
                    if sk: parts.append(f"{sk} matches will be skipped")
                    if duplicates: parts.append(f"{len(duplicates)} ambiguous rows skipped")
                    st.info("  ·  ".join(parts) if parts else "Nothing to import.")

                    # ── Validation ───────────────────────────────────────
                    _ids_entered   = list(new_activity_ids.values())
                    _id_dupes      = len(_ids_entered) != len(set(i.upper() for i in _ids_entered))
                    _id_clash      = any(i.upper() in known_ids for i in _ids_entered)
                    _ow_dupes      = len(manual_overwrites) != len(set(
                                         a.upper() for a in manual_overwrites.values()))
                    _can_confirm   = bool(msp_resolutions) or bool(new_activity_ids) or bool(manual_overwrites)
                    _blocked       = _id_dupes or _id_clash or _ow_dupes

                    # ── Summary ───────────────────────────────────────────
                    ow       = sum(1 for v in msp_resolutions.values() if v["action"] == "overwrite")
                    sk       = sum(1 for v in msp_resolutions.values() if v["action"] == "skip")
                    n_new    = len(new_activity_ids)
                    n_manual = len(manual_overwrites)
                    n_skip   = len(unmatched) - n_new - n_manual
                    parts    = []
                    if ow:       parts.append(f"{ow} auto-matched activities will be updated")
                    if sk:       parts.append(f"{sk} auto-matches skipped")
                    if n_new:    parts.append(f"{n_new} new activities will be added")
                    if n_manual: parts.append(f"{n_manual} manual overwrites will be applied")
                    if n_skip:   parts.append(f"{n_skip} unmatched rows skipped")
                    if duplicates: parts.append(f"{len(duplicates)} ambiguous rows skipped")
                    st.info("  ·  ".join(parts) if parts else "Nothing to import.")

                    if _id_dupes:  st.error("Duplicate Activity IDs in the 'Add as new' fields above.")
                    if _id_clash:  st.error("One or more new Activity IDs already exist in the store.")
                    if _ow_dupes:  st.error("The same stored activity is targeted by more than one manual overwrite.")

                    if st.button(
                        "✅  Confirm MSP Import", type="primary",
                        disabled=(not _can_confirm or _blocked),
                    ):
                        entries       = load_entries()
                        entries_index = {e.get("activity_id", "").upper(): idx
                                         for idx, e in enumerate(entries)}
                        updated = added = manually_updated = 0

                        # Helper: merge MSP progress into a stored entry in-place
                        def _apply_msp(stored: dict, msp_row: dict) -> dict:
                            stored["activity_status"] = msp_row.get("activity_status") or stored["activity_status"]
                            stored["actual_start"]    = msp_row.get("actual_start")    or stored["actual_start"]
                            stored["actual_finish"]   = msp_row.get("actual_finish")   or stored["actual_finish"]
                            stored["pct_complete"]    = msp_row.get("pct_complete")    or stored["pct_complete"]
                            stored["remaining_dur"]   = msp_row.get("remaining_dur")   or stored["remaining_dur"]
                            imp_str = msp_row.pop("comments_import", "") or ""
                            if imp_str:
                                stored["_comments"] = (
                                    import_string_to_comments(imp_str, st.session_state.display_name)
                                    + stored.get("_comments", [])
                                )
                            stored["_submitted_at"] = datetime.now().strftime("%d/%m/%Y %H:%M")
                            stored["_submitted_by"] = st.session_state.display_name
                            return stored

                        # Auto-matched updates
                        for aid, res in msp_resolutions.items():
                            if res["action"] == "overwrite":
                                idx = entries_index.get(aid)
                                if idx is not None:
                                    entries[idx] = _apply_msp(entries[idx], res["msp_row"])
                                    updated += 1

                        # Manual overwrites from unmatched rows
                        for row_idx, target_aid in manual_overwrites.items():
                            idx = entries_index.get(target_aid.upper())
                            if idx is not None:
                                entries[idx] = _apply_msp(entries[idx], unmatched[row_idx])
                                manually_updated += 1

                        # New activities from unmatched rows
                        for idx, new_id in new_activity_ids.items():
                            if not new_id:
                                continue
                            row     = unmatched[idx]
                            imp_str = row.pop("comments_import", "") or ""
                            entries.append({
                                "activity_id":       new_id,
                                "activity_name":     row.get("activity_name", ""),
                                "activity_status":   row.get("activity_status", "Not Started"),
                                "actual_start":      row.get("actual_start", ""),
                                "actual_finish":     row.get("actual_finish", ""),
                                "pct_complete":      row.get("pct_complete", "0"),
                                "remaining_dur":     row.get("remaining_dur", ""),
                                "complete_pct_type": "Physical",
                                "wbs_id":            row.get("wbs_id", ""),
                                "_comments":         import_string_to_comments(imp_str, st.session_state.display_name),
                                "_submitted_at":     datetime.now().strftime("%d/%m/%Y %H:%M"),
                                "_submitted_by":     st.session_state.display_name,
                            })
                            added += 1

                        save_entries(entries)
                        msg = "MSP import complete:"
                        if updated:          msg += f" **{updated}** auto-matched updated"
                        if manually_updated: msg += f", **{manually_updated}** manually overwritten"
                        if added:            msg += f", **{added}** new activities added"
                        st.success(msg)
                        st.rerun()

            else:
                # ── P6 matching flow (existing logic) ─────────────────────
                clean      = [r for r in imported_rows if r["activity_id"].upper() not in known_ids]
                all_match  = [r for r in imported_rows if r["activity_id"].upper() in known_ids]
                # Split conflicts into true changes vs exact duplicates
                conflicts  = [r for r in all_match
                              if not is_exact_duplicate(r, known_ids[r["activity_id"].upper()])]
                duplicates_skipped = len(all_match) - len(conflicts)
                msg = (
                    f"Found **{len(imported_rows)}** rows: **{len(clean)}** new, "
                    f"**{len(conflicts)}** conflict{'s' if len(conflicts) != 1 else ''}"
                )
                if duplicates_skipped:
                    msg += f", **{duplicates_skipped}** identical (skipped)"
                st.success(msg)

                resolutions = {}
                if conflicts:
                    st.divider()
                    st.subheader("⚠️ Conflicts — Activity ID already exists")
                    st.caption("Review each conflict and choose an action before confirming the import.")
                    for row in conflicts:
                        aid      = row["activity_id"].upper()
                        existing = known_ids[aid]
                        label    = f"🔁  {row['activity_id']}  —  {row.get('activity_name', existing.get('activity_name', ''))}"
                        with st.expander(label, expanded=True):
                            fields = [
                                ("Status",          "activity_status", False),
                                ("Actual Start",     "actual_start",    True),
                                ("Actual Finish",    "actual_finish",   True),
                                ("% Complete",       "pct_complete",    False),
                                ("Remaining (days)", "remaining_dur",   False),
                                ("WBS",              "wbs_id",          False),
                            ]
                            # Existing stored comments summary
                            existing_comment_count = len(existing.get("_comments", []))
                            incoming_comment_str   = row.get("comments_import", "").strip()
                            c_cur, c_new = st.columns(2)
                            with c_cur:
                                st.write("**Current (stored)**")
                                for lbl, k, is_date in fields:
                                    val = existing.get(k, "") or ""
                                    st.write(f"- **{lbl}:** {display_dt(val) if is_date else (val or '—')}")
                                st.write(f"- **Comments:** {existing_comment_count} stored comment{'s' if existing_comment_count != 1 else ''}")
                            with c_new:
                                st.write("**Incoming (from file)**")
                                for lbl, k, is_date in fields:
                                    val = row.get(k, "") or ""
                                    st.write(f"- **{lbl}:** {display_dt(val) if is_date else (val or '—')}")
                                st.write(f"- **Comments (user_field_910):** {incoming_comment_str or '—'}")

                            choice = st.radio(
                                "Activity data", ["Overwrite with incoming", "Keep current"],
                                key=f"conflict_{aid}", horizontal=True,
                            )
                            resolutions[aid] = "overwrite" if choice == "Overwrite with incoming" else "skip"

                            # Separate resolution for comments if both sides have data
                            if incoming_comment_str and existing_comment_count > 0:
                                comment_choice = st.radio(
                                    "Comments",
                                    ["Append imported comments to existing", "Keep existing only", "Replace with imported only"],
                                    key=f"comment_conflict_{aid}", horizontal=True,
                                )
                                resolutions[f"comment_{aid}"] = comment_choice
                            elif incoming_comment_str:
                                resolutions[f"comment_{aid}"] = "Append imported comments to existing"
                            else:
                                resolutions[f"comment_{aid}"] = "Keep existing only"

                st.divider()
                ow    = sum(1 for v in resolutions.values() if v == "overwrite")
                sk    = sum(1 for v in resolutions.values() if v == "skip")
                parts = [f"{len(clean)} new entries will be added"]
                if ow: parts.append(f"{ow} existing entries will be overwritten")
                if sk: parts.append(f"{sk} conflicts will be skipped")
                st.info("  ·  ".join(parts))

                if st.button("✅  Confirm Import", type="primary"):
                    entries = load_entries()
                    added = overwritten = skipped = 0
                    for row in clean:
                        # Convert any imported comment string to structured list
                        imp_str = row.pop("comments_import", "") or ""
                        row["_comments"] = import_string_to_comments(imp_str, st.session_state.display_name)
                        row["_submitted_by"] = st.session_state.display_name
                        entries.append(row)
                        added += 1
                    # Build index once before the loop — O(n) not O(n²)
                    entries_index = {e.get("activity_id", "").upper(): idx
                                     for idx, e in enumerate(entries)}
                    for row in conflicts:
                        aid = row["activity_id"].upper()
                        if resolutions.get(aid) == "overwrite":
                            idx = entries_index.get(aid, 0)
                            imp_str = row.pop("comments_import", "") or ""
                            comment_res = resolutions.get(f"comment_{aid}", "Keep existing only")
                            existing_comments = entries[idx].get("_comments", [])
                            imported_comments = import_string_to_comments(imp_str, st.session_state.display_name)
                            if comment_res == "Append imported comments to existing":
                                # Imported go after existing (existing are newer)
                                row["_comments"] = existing_comments + imported_comments
                            elif comment_res == "Replace with imported only":
                                row["_comments"] = imported_comments
                            else:
                                row["_comments"] = existing_comments
                            row["_submitted_by"] = st.session_state.display_name
                            entries[idx] = row
                            overwritten += 1
                        else:
                            skipped += 1
                    save_entries(entries)
                    msg = f"Import complete: **{added}** added"
                    if overwritten: msg += f", **{overwritten}** overwritten"
                    if skipped:     msg += f", **{skipped}** skipped"
                    st.success(msg)
                    st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# TAB: EXPORT TO EXCEL
# ══════════════════════════════════════════════════════════════════════════════

if "export" in tab_index:
    with tab_index["export"]:
        entries = load_entries()
        st.subheader("Export P6-Ready Excel File")
        st.write("The exported file is formatted for direct import into Primavera P6:")
        st.write("- **Row 1:** P6 internal field key names (`task_code`, `act_start_date`, etc.)")
        st.write("- **Row 2:** Column headers · **Sheet name:** `TASK`")
        st.write("- **Date format:** `DD/MM/YYYY HH:MM` stored as proper Excel datetime cells")
        st.write("- **% Complete:** Plain integer")
        st.write("- **Complete Type:** Always Physical as to not overide remaining durations with calulated values from percent complete")
        st.divider()

        if not entries:
            st.warning("No entries to export yet.")
        else:
            st.info(f"{len(entries)} {'entry' if len(entries) == 1 else 'entries'} ready to export.")

            # ── Project name / WBS prefix ─────────────────────────────────
            # Detect the current prefix from the first entry that has one
            _current_prefix = ""
            for _e in entries:
                _wbs = _e.get("wbs_id", "")
                if _wbs and "." in _wbs:
                    _first_seg = _wbs.split(".")[0]
                    if not _first_seg.isdigit():
                        _current_prefix = _first_seg
                        break

            rename_wbs = st.checkbox(
                "Update project name in WBS codes",
                value=False,
                key="export_rename_wbs",
                help="Use this when the P6 project name has changed since activities were entered.",
            )

            project_name_out = ""
            if rename_wbs:
                if _current_prefix:
                    st.caption(f"Current prefix detected: **{_current_prefix}**")
                else:
                    st.caption("No non-numeric prefix detected in stored WBS codes.")

                project_name_out = st.text_input(
                    "New P6 project name *",
                    placeholder="e.g. ProjectB",
                    key="export_project_name",
                ).strip()

                st.warning(
                    "⚠️ The project name entered here must exactly match the project name "
                    "in Primavera P6, including capitalisation. If it does not match, "
                    "the import into P6 may fail or assign activities to the wrong WBS."
                )

                if project_name_out and _current_prefix:
                    st.caption(
                        f"Preview: **{_current_prefix}.1.2.3** → "
                        f"**{project_name_out}.1.2.3**"
                    )

            st.divider()
            _ready = not rename_wbs or bool(project_name_out)
            if not _ready:
                st.error("Enter the new project name above before downloading.")

            excel_bytes = build_excel(entries, project_name=project_name_out)
            fname = f"p6_asbuilt_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            st.download_button(
                label="⬇️  Download P6-Ready XLSX",
                data=excel_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
                disabled=not _ready,
            )
            st.divider()
            st.write("**How to import into P6:**")
            st.write("1. Open **Primavera P6 Professional**")
            st.write("2. `File` → `Import` → `Spreadsheet (XLSX)` → **Next**")
            st.write("3. Browse to the downloaded file → **Next**")
            st.write("4. Select **Activities** → choose your project → **Finish**")
            st.write("5. Press **F9** to reschedule")

# ══════════════════════════════════════════════════════════════════════════════
# TAB: PHOTO LOG
# ══════════════════════════════════════════════════════════════════════════════

if "photos" in tab_index:
    with tab_index["photos"]:
        ensure_photo_dir()
        entries           = load_entries()
        known_ids_ordered = [e["activity_id"] for e in entries]
        known_map         = {e["activity_id"].upper(): e for e in entries}
        can_upload        = has_permission("photos")

        # ── Cache photos list and all lookups in session_state ─────────────
        # These are only reloaded from disk when a photo is uploaded or
        # deleted (st.rerun with cache cleared). All other interactions
        # — dropdown changes, filter changes, assignment changes — read
        # entirely from session_state, avoiding any disk I/O or list rebuilds.

        if "photo_list" not in st.session_state:
            st.session_state["photo_list"] = load_photos()
        photos = st.session_state["photo_list"]

        if "photo_assignments" not in st.session_state:
            st.session_state["photo_assignments"] = load_assignments()
        assignments = st.session_state["photo_assignments"]

        # Lookups are rebuilt only when assignments change (cheap, in-memory)
        # We store them in session_state so a dropdown change doesn't rebuild them
        # Hash photo_id+activity_id pairs so any change — not just count — triggers rebuild
        assign_sig = hash(tuple((a["photo_id"], a["activity_id"]) for a in assignments))
        if (st.session_state.get("_assign_sig") != assign_sig
                or "photo_to_aids" not in st.session_state):
            _pta, _atp, _pm = {}, {}, {}
            for a in assignments:
                _pta.setdefault(a["photo_id"], []).append(a["activity_id"])
                _atp.setdefault(a["activity_id"].upper(), []).append(a["photo_id"])
            for p in photos:
                _pm[p["id"]] = p
            st.session_state["photo_to_aids"] = _pta
            st.session_state["aid_to_pids"]   = _atp
            st.session_state["photo_map"]     = _pm
            st.session_state["_assign_sig"]   = assign_sig

        photo_to_aids = st.session_state["photo_to_aids"]
        aid_to_pids   = st.session_state["aid_to_pids"]
        photo_map     = st.session_state["photo_map"]

        # ── Step 1: Upload ─────────────────────────────────────────────────
        if can_upload:
            st.subheader("Step 1 — Upload Photo")
            st.caption("Upload the image and set its date and comment. You will assign it to activities in Step 2.")

            up_col1, up_col2 = st.columns(2)
            with up_col1:
                photo_date = st.date_input(
                    "Date of Photo *",
                    value=date.today(),
                    format="DD/MM/YYYY",
                    key="photo_upload_date",
                )
            with up_col2:
                comment = st.text_input(
                    "Comment",
                    placeholder="e.g. North elevation — formwork complete (max 100 chars)",
                    max_chars=100,
                    key="photo_upload_comment",
                )

            uploaded_file = st.file_uploader(
                "Choose image *",
                type=["jpg", "jpeg", "png", "webp", "gif"],
                key="photo_upload_file",
            )

            if uploaded_file:
                # Read bytes once — reused for both preview and upload
                _file_bytes = uploaded_file.read()
                if _PILLOW:
                    _prev = ImageOps.exif_transpose(Image.open(io.BytesIO(_file_bytes)))
                    st.image(_prev, caption="Preview", width=400)
                else:
                    st.image(io.BytesIO(_file_bytes), caption="Preview", width=400)

            if st.button("📤  Upload Photo", type="primary", disabled=not uploaded_file):
                record = upload_photo(
                    photo_date    = photo_date,
                    comment       = comment,
                    file_bytes    = _file_bytes,
                    original_name = uploaded_file.name,
                    uploaded_by   = st.session_state.display_name,
                )
                st.success(f"✅ Photo uploaded — now assign it to activities in Step 2.")
                # Store the new photo id so Step 2 pre-selects it
                st.session_state["last_uploaded_photo_id"] = record["id"]
                load_image_bytes.clear()
                for _k in ("photo_list","photo_map","photo_to_aids","aid_to_pids","_assign_sig"):
                    st.session_state.pop(_k, None)
                st.rerun()

            st.divider()

            # ── Step 2: Assign ─────────────────────────────────────────────
            st.subheader("Step 2 — Assign Photo to Activities")
            st.caption("Select a photo from the library and assign it to one or more activities.")

            if not photos:
                st.info("No photos uploaded yet — upload one above first.")
            elif not known_ids_ordered:
                st.warning("No activities found — submit some entries first.")
            else:
                # Photo selector — default to most recently uploaded
                last_id   = st.session_state.get("last_uploaded_photo_id", photos[-1]["id"])
                photo_ids = [p["id"] for p in sorted(photos, key=lambda p: p["uploaded_at"], reverse=True)]

                def photo_label(pid: str) -> str:
                    p = photo_map.get(pid, {})
                    try:
                        dt_str = datetime.strptime(p.get("photo_date",""), "%Y-%m-%d").strftime("%d/%m/%Y")
                    except ValueError:
                        dt_str = p.get("photo_date", "")
                    n = len(photo_to_aids.get(pid, []))
                    suffix = f" — {n} assignment{'s' if n != 1 else ''}"
                    comment_preview = (p.get("comment","")[:30] + "…") if p.get("comment","") else "no comment"
                    return f"{dt_str}  ·  {comment_preview}{suffix}"

                default_idx = photo_ids.index(last_id) if last_id in photo_ids else 0
                selected_photo_id = st.selectbox(
                    "Select photo",
                    options=photo_ids,
                    index=default_idx,
                    format_func=photo_label,
                    key="assign_photo_select",
                )

                selected_photo = photo_map.get(selected_photo_id, {})

                # Show thumbnail only — avoids loading full image in Step 2
                _thumb = selected_photo.get("thumb", "")
                _thumb_bytes = load_image_bytes(_thumb) if _thumb else load_image_bytes(selected_photo.get("filename",""))
                if _thumb_bytes:
                    st.image(_thumb_bytes, width=200)

                # Current assignments for this photo
                current_aids = set(a.upper() for a in photo_to_aids.get(selected_photo_id, []))

                # Multi-select for activities
                assign_ids = st.multiselect(
                    "Assign to activities",
                    options=known_ids_ordered,
                    default=[aid for aid in known_ids_ordered if aid.upper() in current_aids],
                    format_func=lambda aid: f"{aid}  —  {known_map[aid.upper()]['activity_name']}",
                    key="assign_activity_select",
                )

                if st.button("💾  Save Assignments", type="primary"):
                    # Add new assignments
                    new_aids = [aid for aid in assign_ids if aid.upper() not in current_aids]
                    if new_aids:
                        assign_photo(selected_photo_id, new_aids, st.session_state.display_name)
                    # Remove deselected assignments
                    removed_aids = [aid for aid in known_ids_ordered
                                    if aid.upper() in current_aids and aid not in assign_ids]
                    for aid in removed_aids:
                        unassign_photo(selected_photo_id, aid)
                    # session_state updated inside assign/unassign — no rerun needed
                    st.success(
                        f"Assignments saved — "
                        f"{len(new_aids)} added, {len(removed_aids)} removed."
                    )

            st.divider()

        # ── Gallery (all roles) ────────────────────────────────────────────
        st.subheader("Photo Gallery")

        if not photos:
            st.info("No photos uploaded yet.")
        else:
            # Filter controls
            f_col1, f_col2 = st.columns([2, 3])
            with f_col1:
                # Only show activity IDs that have at least one assignment
                assigned_aids = sorted({a["activity_id"] for a in assignments})
                filter_id = st.selectbox(
                    "Filter by Activity",
                    options=["— All —"] + assigned_aids,
                    key="photo_filter_id",
                )
            with f_col2:
                filter_text = st.text_input(
                    "Search comments",
                    placeholder="Type to search…",
                    key="photo_filter_text",
                ).strip().lower()

            # Apply filters
            if filter_id != "— All —":
                visible_pids = set(aid_to_pids.get(filter_id.upper(), []))
                filtered = [p for p in photos if p["id"] in visible_pids]
            else:
                filtered = list(photos)

            if filter_text:
                filtered = [p for p in filtered
                            if filter_text in p.get("comment", "").lower()
                            or any(filter_text in aid.lower()
                                   for aid in photo_to_aids.get(p["id"], []))]

            filtered = sorted(filtered, key=lambda p: p["photo_date"], reverse=True)

            if not filtered:
                st.info("No photos match the current filter.")
            else:
                PHOTO_PAGE_SIZE = 18  # 6 rows of 3 — adjust as needed
                total_photo_pages = max(1, (len(filtered) + PHOTO_PAGE_SIZE - 1) // PHOTO_PAGE_SIZE)
                photo_page = st.number_input(
                    f"Page (1 – {total_photo_pages})",
                    min_value=1, max_value=total_photo_pages, value=1, step=1,
                    key="photo_gallery_page",
                ) - 1  # zero-based
                page_start = photo_page * PHOTO_PAGE_SIZE
                page_end   = page_start + PHOTO_PAGE_SIZE
                page_photos = filtered[page_start:page_end]

                st.caption(
                    f"Showing {page_start + 1}–{min(page_end, len(filtered))} "
                    f"of {len(filtered)} photo{'s' if len(filtered) != 1 else ''}"
                )
                st.divider()

                COLS = 3
                for row_start in range(0, len(page_photos), COLS):
                    cols = st.columns(COLS)
                    for col_idx, photo in enumerate(page_photos[row_start:row_start + COLS]):
                        with cols[col_idx]:
                            with st.container(border=True):
                                # Use thumbnail for fast loading
                                _thumb = photo.get("thumb", "")
                                _tb    = load_image_bytes(_thumb) if _thumb else load_image_bytes(photo.get("filename",""))
                                if _tb:
                                    st.image(_tb, width=350)
                                else:
                                    st.caption("Image missing")

                                try:
                                    dt_str = datetime.strptime(
                                        photo["photo_date"], "%Y-%m-%d"
                                    ).strftime("%d/%m/%Y")
                                except ValueError:
                                    dt_str = photo["photo_date"]

                                st.caption(f"📅 {dt_str}")
                                if photo.get("comment"):
                                    st.caption(photo["comment"])

                                # Assigned activities
                                aids = photo_to_aids.get(photo["id"], [])
                                if aids:
                                    for aid in aids:
                                        act = known_map.get(aid.upper(), {})
                                        st.caption(f"📌 {aid} — {act.get('activity_name','')} ({act.get('activity_status','')})")
                                else:
                                    st.caption("Not assigned")

                                st.caption(f"By {photo.get('uploaded_by','—')}  ·  {photo['uploaded_at']}")

                                if can_upload:
                                    with st.expander("✏️ Assign / Remove / Delete"):
                                        new_assign = st.multiselect(
                                            "Assign to:",
                                            options=[a for a in known_ids_ordered
                                                     if a.upper() not in {x.upper() for x in aids}],
                                            format_func=lambda a: f"{a} — {known_map[a.upper()]['activity_name']}",
                                            key=f"gallery_assign_{photo['id']}",
                                        )
                                        if st.button("＋ Assign", key=f"gallery_assign_btn_{photo['id']}",
                                                     disabled=not new_assign):
                                            assign_photo(photo["id"], new_assign, st.session_state.display_name)
                                            st.session_state["_assign_sig"] = None  # force lookup rebuild
                                            st.toast(f"Assigned to {len(new_assign)} activity/activities.")

                                        if aids:
                                            for aid in aids:
                                                if st.button(f"Remove from {aid}",
                                                             key=f"unassign_{photo['id']}_{aid}"):
                                                    unassign_photo(photo["id"], aid)
                                                    st.session_state["_assign_sig"] = None  # force lookup rebuild
                                                    st.toast(f"Removed from {aid}.")

                                        if st.button("🗑 Delete permanently",
                                                     key=f"photo_del_{photo['id']}", type="primary"):
                                            delete_photo_file(photo["id"])
                                            load_image_bytes.clear()
                                            for _k in ("photo_list","photo_map","photo_assignments",
                                                       "photo_to_aids","aid_to_pids","_assign_sig"):
                                                st.session_state.pop(_k, None)
                                            st.rerun()
