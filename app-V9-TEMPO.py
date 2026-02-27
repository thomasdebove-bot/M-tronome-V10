# =========================================================
# TEMPO • METRONOME • COMPTE-RENDU SYNTHESE (HTML / PRINT-FIRST) — V3.2+
# =========================================================
# ✅ Bleu = sujets traités dans la réunion sélectionnée (Meeting/ID)
# ✅ Rappels = tâches non clôturées ET en retard à la DATE DE SEANCE (pas "aujourd’hui")
# ✅ À suivre = tâches non clôturées NON en retard à la date de séance (inclut réunions précédentes)
# ✅ Rappels + À suivre classés PAR ZONE
# ✅ KPI "Rappels par entreprise" (logo + compteur)
# ✅ Bandeau projet via Projects.csv (image + infos)
# ✅ Images dans TÂCHES/MEMOS/RAPPELS/ÀSUIVRE (détection automatique colonne + parsing robuste)
# ✅ Commentaires tâches si dispo
# ✅ Ajout de mémos épinglés par zone (modal) — dispo aussi en "version imprimable"
# ✅ Plus de "badges" instables : colonne dédiée (UI) + colonne "Type" (PDF)
#
# INSTALL
#   python -m pip install fastapi uvicorn pandas openpyxl
#
# RUN
#   python -m uvicorn app:app --host 0.0.0.0 --port 8090 --reload
# =========================================================

from __future__ import annotations

import base64
import json
import os
import re
import urllib.parse
import urllib.request
from datetime import date, timedelta
from typing import Dict, List, Optional, Tuple

import pandas as pd
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import HTMLResponse, JSONResponse

app = FastAPI(title="TEMPO • CR Synthèse (METRONOME)")

# -------------------------
# PATHS (UNC)
# -------------------------
ENTRIES_PATH = os.getenv(
    "METRONOME_ENTRIES",
    r"\\192.168.10.100\02 - affaires\02.2 - SYNTHESE\ZZ - METRONOME\Entries (Tasks & Memos).csv",
)
MEETINGS_PATH = os.getenv(
    "METRONOME_MEETINGS",
    r"\\192.168.10.100\02 - affaires\02.2 - SYNTHESE\ZZ - METRONOME\Meetings.csv",
)
COMPANIES_PATH = os.getenv(
    "METRONOME_COMPANIES",
    r"\\192.168.10.100\02 - affaires\02.2 - SYNTHESE\ZZ - METRONOME\Companies.csv",
)
PROJECTS_PATH = os.getenv(
    "METRONOME_PROJECTS",
    r"\\192.168.10.100\02 - affaires\02.2 - SYNTHESE\ZZ - METRONOME\Projects.csv",
)
LOGO_TEMPO_PATH = os.getenv(
    "METRONOME_LOGO",
    r"\\192.168.10.100\02 - affaires\02.2 - SYNTHESE\ZZ - METRONOME\Content\Logo TEMPO.png",
)
LOGO_RYTHME_PATH = os.getenv(
    "METRONOME_LOGO_RYTHME",
    r"\\192.168.10.100\02 - affaires\02.2 - SYNTHESE\ZZ - METRONOME\Content\Rythme.png",
)
LOGO_T_MARK_PATH = os.getenv(
    "METRONOME_LOGO_TMARK",
    r"\\192.168.10.100\02 - affaires\02.2 - SYNTHESE\ZZ - METRONOME\Content\T logo.png",
)
LOGO_QR_PATH = os.getenv(
    "METRONOME_QR",
    r"\\192.168.10.100\02 - affaires\02.2 - SYNTHESE\ZZ - METRONOME\Content\QR CODE.png",
)
DOCUMENTS_PATH = os.getenv(
    "METRONOME_DOCUMENTS",
    r"\\192.168.10.100\02 - affaires\02.2 - SYNTHESE\ZZ - METRONOME\Documents.csv",
)
COMMENTS_PATH = os.getenv(
    "METRONOME_COMMENTS",
    r"\\192.168.10.100\02 - affaires\02.2 - SYNTHESE\ZZ - METRONOME\Comments.csv",
)

# -------------------------
# COLUMN NAMES (METRONOME EXPORTS)
# -------------------------
# Entries
E_COL_ID = "🔒 Row ID"
E_COL_TITLE = "Title"
E_COL_PROJECT_TITLE = "Project/Title"
E_COL_MEETING_ID = "Meeting/ID"
E_COL_IS_TASK = "Category/Task"
E_COL_CATEGORY = "Category/Name to display"
E_COL_AREAS = "Areas/Names"
E_COL_PACKAGES = "Packages/Names"
E_COL_COMPANY_TASK = "Company/Name for Tasks"
E_COL_OWNER = "Owner for Tasks/Full Name"
E_COL_CREATED = "Declaration Date/Editable"
E_COL_DEADLINE = "Deadline & Status for Tasks/Deadline"
E_COL_STATUS = "Deadline & Status for Tasks/Status Emoji + Text"
E_COL_COMPLETED = "Completed/true/false"
E_COL_COMPLETED_END = "Completed/Declared End"
E_COL_IMAGES_URLS = "Images/Autom input as text (dev)"  # nominal (may vary in exports)

E_COL_TASK_COMMENT_TEXT = "Comment for Tasks/Text"
E_COL_TASK_COMMENT_FULL = "Comment for Tasks/Full text to display if existing (dev)"
E_COL_TASK_COMMENT_AUTHOR = "Comment for Tasks/Editor Name (dev)"
E_COL_TASK_COMMENT_DATE = "Comment for Tasks/Date"

# Meetings
M_COL_ID = "🔒 Row ID"
M_COL_DATE = "Date/Editable"
M_COL_DATE_DISPLAY = "Date/To display (dev)"
M_COL_PROJECT_TITLE = "Project/Title (dev)"
M_COL_ATT_IDS = "Companies/Attending IDs"
M_COL_MISS_IDS = "Companies/Missing IDs"
M_COL_MISS_CALC_IDS = "Companies/Missing Calculated IDs (dev)"
M_COL_TASKS_COUNT = "Entries/Tasks Count"
M_COL_MEMOS_COUNT = "Entries/Memos Count"

# Companies
C_COL_ID = "🔒 Row ID"
C_COL_NAME = "Name"
C_COL_LOGO = "Logo"

# Projects
P_COL_TITLE = "Title"
P_COL_DESC = "Description"
P_COL_IMAGE = "Image"
P_COL_START_SENT = "Timeline/Start Sentence"
P_COL_END_SENT = "Timeline/End Sentence"
P_COL_ARCHIVED = "Archived/Text"

# -------------------------
# CACHE
# -------------------------
_cache = {
    "entries": (None, None),
    "meetings": (None, None),
    "companies": (None, None),
    "projects": (None, None),
    "documents": (None, None),
}


def _mtime(path: str) -> float:
    try:
        return os.path.getmtime(path)
    except Exception:
        return -1.0


class MissingDataError(RuntimeError):
    def __init__(self, label: str, path: str, env_var: str):
        super().__init__(f"Fichier manquant pour {label}: {path} (env: {env_var})")
        self.label = label
        self.path = path
        self.env_var = env_var


def _load_csv(path: str) -> pd.DataFrame:
    return pd.read_csv(path, encoding="utf-8-sig")


def _require_csv(path: str, label: str, env_var: str) -> None:
    if not os.path.exists(path):
        raise MissingDataError(label=label, path=path, env_var=env_var)


def get_entries() -> pd.DataFrame:
    m = _mtime(ENTRIES_PATH)
    old_m, df = _cache["entries"]
    if df is None or m != old_m:
        _require_csv(ENTRIES_PATH, "Entries", "METRONOME_ENTRIES")
        df = _load_csv(ENTRIES_PATH)
        _cache["entries"] = (m, df)
    return df


def get_meetings() -> pd.DataFrame:
    m = _mtime(MEETINGS_PATH)
    old_m, df = _cache["meetings"]
    if df is None or m != old_m:
        _require_csv(MEETINGS_PATH, "Meetings", "METRONOME_MEETINGS")
        df = _load_csv(MEETINGS_PATH)
        _cache["meetings"] = (m, df)
    return df


def get_companies() -> pd.DataFrame:
    m = _mtime(COMPANIES_PATH)
    old_m, df = _cache["companies"]
    if df is None or m != old_m:
        _require_csv(COMPANIES_PATH, "Companies", "METRONOME_COMPANIES")
        df = _load_csv(COMPANIES_PATH)
        _cache["companies"] = (m, df)
    return df


def get_projects() -> pd.DataFrame:
    m = _mtime(PROJECTS_PATH)
    old_m, df = _cache["projects"]
    if df is None or m != old_m:
        _require_csv(PROJECTS_PATH, "Projects", "METRONOME_PROJECTS")
        df = _load_csv(PROJECTS_PATH)
        _cache["projects"] = (m, df)
    return df


def get_documents() -> pd.DataFrame:
    m = _mtime(DOCUMENTS_PATH)
    old_m, df = _cache["documents"]
    if df is None or m != old_m:
        _require_csv(DOCUMENTS_PATH, "Documents", "METRONOME_DOCUMENTS")
        df = _load_csv(DOCUMENTS_PATH)
        _cache["documents"] = (m, df)
    return df


# -------------------------
# UTILITIES
# -------------------------
def _escape(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s)
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#039;")
    )


def _series(df: pd.DataFrame, col: str, default) -> pd.Series:
    if col in df.columns:
        data = df[col]
        if isinstance(data, pd.DataFrame):
            return data.iloc[:, 0]
        return data
    return pd.Series([default] * len(df), index=df.index)


def _filter_entries_by_created_range(
    df: pd.DataFrame, start_date: Optional[date], end_date: Optional[date]
) -> pd.DataFrame:
    if df.empty or (start_date is None and end_date is None):
        return df
    created = _series(df, E_COL_CREATED, None).apply(_parse_date_any)
    mask = pd.Series(True, index=df.index)
    if start_date is not None:
        mask &= created.notna() & (created >= start_date)
    if end_date is not None:
        mask &= created.notna() & (created <= end_date)
    return df.loc[mask].copy()


def _safe_int(v) -> int:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0
    try:
        return int(v)
    except (TypeError, ValueError):
        return 0


def _parse_ids(v) -> List[str]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return []
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return []
    return [p.strip() for p in s.split(",") if p.strip()]


def _bool_true(v) -> bool:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return False
    s = str(v).strip().lower()
    return s in {"true", "1", "yes", "y", "vrai"}


def _parse_date_any(v) -> Optional[date]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    if not s or s.lower() in {"nan", "none"}:
        return None
    m = re.match(r"^(\d{2})/(\d{2})/(\d{2,4})", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000
        try:
            return date(y, mo, d)
        except ValueError:
            return None
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return date(y, mo, d)
        except ValueError:
            return None
    try:
        dt = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.isna(dt):
            return None
        return dt.date()
    except Exception:
        return None


def _fmt_date(d: Optional[date]) -> str:
    return d.strftime("%d/%m/%y") if d else ""


def _norm_name(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip()).lower()


def _trigram(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    raw = re.sub(r"[^A-Za-z0-9]", "", str(s).strip())
    if not raw:
        return ""
    return raw[:3].upper()


def _lot_abbrev(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    text = str(s).strip()
    if not text:
        return ""
    rules = [
        ("Électricité", "ELE"),
        ("Courants forts", "CFO"),
        ("Courants faibles", "CFA"),
        ("Plomberie", "PLB"),
        ("CVC", "CVC"),
        ("Structure", "STR"),
        ("Gros Oeuvre", "GOE"),
        ("Synthèse", "SYN"),
        ("Entreprise Générale", "EG"),
        ("Sprinklage", "SPK"),
    ]
    text_lower = text.lower()
    for label, abbrev in rules:
        if label.lower() in text_lower:
            return abbrev
    return _trigram(text)


def _lot_abbrev_list(value: str) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    raw = str(value).strip()
    if not raw:
        return ""
    parts = [p.strip() for p in re.split(r"[,;/]+", raw) if p.strip()]
    if not parts:
        return ""
    mapped = [_lot_abbrev(p) for p in parts]
    mapped = [m for m in mapped if m]
    if len(mapped) > 1:
        return "PL"
    return " / ".join(mapped)


def _concerne_trigram(value: str) -> str:
    trigram = _trigram(value)
    return trigram or "PE"


def _has_multiple_companies(value: str) -> bool:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return False
    raw = str(value).strip()
    if not raw:
        return False
    parts = [p.strip() for p in re.split(r"[,;/]+", raw) if p.strip()]
    return len(parts) > 1


def _split_words(value: str) -> set[str]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return set()
    raw = str(value).strip()
    if not raw:
        return set()
    return {part for part in re.split(r"[^\w]+", raw) if part}


def _logo_data_url(path: str) -> str:
    if not path:
        return ""
    normalized = os.path.normpath(path)
    if not os.path.exists(normalized):
        return ""
    try:
        with open(normalized, "rb") as f:
            data = base64.b64encode(f.read()).decode("utf-8")
        ext = os.path.splitext(normalized)[1].lower()
        if ext in {".jpg", ".jpeg"}:
            mime = "image/jpeg"
        elif ext == ".svg":
            mime = "image/svg+xml"
        else:
            mime = "image/png"
        return f"data:{mime};base64,{data}"
    except Exception:
        return ""


def _meeting_sequence_for_project(
    meetings_df: pd.DataFrame, meeting_id: str
) -> Tuple[int, int]:
    if meetings_df.empty:
        return 1, 1
    df = meetings_df.copy()
    df["__mid__"] = _series(df, M_COL_ID, "").fillna("").astype(str).str.strip()
    df["__mdate__"] = _series(df, M_COL_DATE, None).apply(_parse_date_any)
    df = df.loc[df["__mid__"] != ""].copy()
    if df.empty:
        return 1, 1
    df = df.sort_values(by=["__mdate__", "__mid__"], ascending=[True, True])
    ids = df["__mid__"].tolist()
    total = len(ids)
    if str(meeting_id) in ids:
        index = ids.index(str(meeting_id)) + 1
    else:
        index = total
    index = max(1, index)
    total = max(1, total)
    return index, total


# -------------------------
# IMAGES (robust)
# -------------------------
def detect_images_column(df: pd.DataFrame) -> Optional[str]:
    """Return likely image URL column name."""
    if df is None or df.empty:
        return None
    if E_COL_IMAGES_URLS in df.columns:
        return E_COL_IMAGES_URLS
    candidates = [c for c in df.columns if "images" in str(c).lower()]
    if not candidates:
        return None
    candidates.sort(key=lambda c: (0 if "autom" in str(c).lower() else 1, len(str(c))))
    return candidates[0]


def detect_memo_images_column(df: pd.DataFrame) -> Optional[str]:
    if df is None or df.empty:
        return None
    memo_candidates = [
        c
        for c in df.columns
        if "image" in str(c).lower() and "memo" in str(c).lower()
    ]
    if memo_candidates:
        memo_candidates.sort(key=lambda c: len(str(c)))
        return memo_candidates[0]
    return detect_images_column(df)


def parse_image_urls_any(v) -> List[str]:
    """Parse robust URLs (http/https) from a cell."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return []
    s = str(v)
    if not s.strip() or s.strip().lower() == "nan":
        return []
    urls = re.findall(r"https?://[^\s,\]\)\"\'<>]+", s)
    out, seen = [], set()
    for u in urls:
        u = u.strip()
        if u and u not in seen:
            out.append(u)
            seen.add(u)
    return out


def _format_entry_text_html(v) -> str:
    """Normalize text for tasks/memos and preserve bullet-like line breaks in HTML."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v)
    if not s.strip() or s.strip().lower() == "nan":
        return ""
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n[ \t]+", "\n", s)
    s = re.sub(r"(?<!\n)\s*(•|●|◦|▪|‣|\*)\s+", r"\n\1 ", s)
    s = re.sub(r"(?<!\n)(?<!\w)-\s+(?=\S)", r"\n- ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return _escape(s.strip()).replace("\n", "<br>")


def render_images_gallery(urls: List[str], print_mode: bool) -> str:
    if not urls:
        return ""
    max_imgs = 3 if print_mode else 10
    thumbs = []
    for u in urls[:max_imgs]:
        uu = _escape(u)
        thumbs.append(
            f"""
          <a class="imgThumb" href="{uu}" target="_blank" rel="noopener">
            <img src="{uu}" loading="lazy" alt="" referrerpolicy="no-referrer" />
            <span class="imgGrip" title="Redimensionner"></span>
          </a>
        """
        )
    return f"""<div class="imgRow">{''.join(thumbs)}</div>"""


# -------------------------
# COMMENTS (TASKS)
# -------------------------
def render_task_comment(r) -> str:
    txt = r.get(E_COL_TASK_COMMENT_FULL)
    if txt is None or (isinstance(txt, float) and pd.isna(txt)) or str(txt).strip() == "":
        txt = r.get(E_COL_TASK_COMMENT_TEXT)
    if txt is None or (isinstance(txt, float) and pd.isna(txt)) or str(txt).strip() == "":
        return ""
    author = _escape(r.get(E_COL_TASK_COMMENT_AUTHOR, ""))
    d = _fmt_date(_parse_date_any(r.get(E_COL_TASK_COMMENT_DATE)))
    body = _format_entry_text_html(txt)
    meta = " • ".join([x for x in [author, d] if x])
    return f"""
      <div class="topicComment">
        <div class="metaLabel">Commentaire</div>
        <div class="metaVal">{meta or "—"}</div>
        <div style="margin-top:6px">{body}</div>
      </div>
    """


def render_entry_comment(r) -> str:
    txt = r.get(E_COL_TASK_COMMENT_FULL)
    if txt is None or (isinstance(txt, float) and pd.isna(txt)) or str(txt).strip() == "":
        txt = r.get(E_COL_TASK_COMMENT_TEXT)
    if txt is None or (isinstance(txt, float) and pd.isna(txt)) or str(txt).strip() == "":
        return ""
    author = _escape(r.get(E_COL_TASK_COMMENT_AUTHOR, ""))
    d = _fmt_date(_parse_date_any(r.get(E_COL_TASK_COMMENT_DATE)))
    company = _escape(r.get(E_COL_COMPANY_TASK, ""))
    body = _format_entry_text_html(txt)
    meta = " • ".join([x for x in [author, company, d] if x])
    return f"""
      <div class="entryComment">
        <div class="metaVal">{meta or "—"}</div>
        <div style="margin-top:6px">{body}</div>
      </div>
    """


# -------------------------
# COMPANIES
# -------------------------
def companies_map_by_id() -> Dict[str, Dict[str, str]]:
    c = get_companies()
    mp = {}
    for _, r in c.iterrows():
        cid = str(r.get(C_COL_ID, "")).strip()
        if not cid:
            continue
        mp[cid] = {
            "name": str(r.get(C_COL_NAME, "")).strip(),
            "logo": str(r.get(C_COL_LOGO, "")).strip(),
        }
    return mp


def companies_logo_by_name() -> Dict[str, str]:
    c = get_companies()
    out = {}
    for _, r in c.iterrows():
        name = str(r.get(C_COL_NAME, "")).strip()
        logo = str(r.get(C_COL_LOGO, "")).strip()
        if name:
            out[_norm_name(name)] = logo
    return out


# -------------------------
# PROJECT INFO
# -------------------------
def project_info_by_title(project_title: str) -> Dict[str, str]:
    p = get_projects().copy()
    p[P_COL_TITLE] = p[P_COL_TITLE].fillna("").astype(str).str.strip()
    row = p.loc[p[P_COL_TITLE] == project_title]
    if row.empty:
        return {"title": project_title, "desc": "", "image": "", "start": "", "end": "", "status": ""}
    r = row.iloc[0]
    return {
        "title": str(r.get(P_COL_TITLE, "")).strip() or project_title,
        "desc": str(r.get(P_COL_DESC, "")).strip(),
        "image": str(r.get(P_COL_IMAGE, "")).strip(),
        "start": str(r.get(P_COL_START_SENT, "")).strip(),
        "end": str(r.get(P_COL_END_SENT, "")).strip(),
        "status": str(r.get(P_COL_ARCHIVED, "")).strip(),
    }


# -------------------------
# MEETING + ENTRIES
# -------------------------
def meeting_row(meeting_id: str) -> pd.Series:
    m = get_meetings()
    row = m.loc[m[M_COL_ID].astype(str) == str(meeting_id)]
    if row.empty:
        raise HTTPException(status_code=404, detail="Meeting not found")
    return row.iloc[0]


def entries_for_meeting(meeting_id: str) -> pd.DataFrame:
    e = get_entries()
    return e.loc[e[E_COL_MEETING_ID].astype(str) == str(meeting_id)].copy()


def compute_presence_lists(mrow: pd.Series) -> Tuple[List[Dict], List[Dict]]:
    mp = companies_map_by_id()
    attending_ids = _parse_ids(mrow.get(M_COL_ATT_IDS))
    missing_ids = _parse_ids(mrow.get(M_COL_MISS_IDS))
    if not missing_ids:
        missing_ids = _parse_ids(mrow.get(M_COL_MISS_CALC_IDS))

    def _to_items(ids: List[str]) -> List[Dict]:
        items = []
        for cid in ids:
            info = mp.get(cid, {"name": f"ID:{cid}", "logo": ""})
            items.append({"id": cid, "name": info.get("name", ""), "logo": info.get("logo", "")})
        items.sort(key=lambda x: (x["name"] or "").lower())
        return items

    return _to_items(attending_ids), _to_items(missing_ids)


# -------------------------
# KPI
# -------------------------
def kpis(mrow: pd.Series, edf: pd.DataFrame, ref_date: date) -> Dict[str, int]:
    tasks_count = _safe_int(mrow.get(M_COL_TASKS_COUNT))
    memos_count = _safe_int(mrow.get(M_COL_MEMOS_COUNT))
    total = len(edf)

    is_task = _series(edf, E_COL_IS_TASK, False).apply(_bool_true)
    tasks = edf[is_task].copy()
    completed = _series(tasks, E_COL_COMPLETED, False).apply(_bool_true)
    open_tasks = tasks[~completed]
    closed_tasks = tasks[completed]

    deadlines = _series(open_tasks, E_COL_DEADLINE, None).apply(_parse_date_any)
    late = (deadlines.notna()) & (deadlines < ref_date)
    late_count = int(late.sum())

    return {
        "total_entries": int(total),
        "tasks_meeting": int(tasks_count),
        "memos_meeting": int(memos_count),
        "open_tasks": int(len(open_tasks)),
        "closed_tasks": int(len(closed_tasks)),
        "late_tasks": int(late_count),
    }


# -------------------------
# REMINDERS / FOLLOW UPS (PROJECT-WIDE) — based on ref_date (date de séance)
# -------------------------
def reminder_level(deadline: Optional[date], completed: bool, ref_date: date) -> Optional[int]:
    """Rappel = tâche non clôturée et en retard à la date de séance (ref_date)."""
    if completed or not deadline:
        return None
    days_late = (ref_date - deadline).days
    if days_late <= 0:
        return None
    return ((days_late - 1) // 7) + 1


def reminder_level_at_done(deadline: Optional[date], done_date: Optional[date]) -> Optional[int]:
    """Rappel historique à la clôture: retard constaté à la date de fin."""
    if not deadline or not done_date:
        return None
    days_late = (done_date - deadline).days
    if days_late <= 0:
        return None
    return ((days_late - 1) // 7) + 1


def _explode_areas(df: pd.DataFrame) -> pd.DataFrame:
    if E_COL_AREAS in df.columns:
        df["__area__"] = df[E_COL_AREAS].fillna("").astype(str).str.strip()
        df.loc[df["__area__"] == "", "__area__"] = "Général"
    else:
        df["__area__"] = "Général"
    df["__area_list__"] = df["__area__"].apply(lambda s: [x.strip() for x in s.split(",")] if "," in s else [s])
    df = df.explode("__area_list__")
    df["__area_list__"] = df["__area_list__"].fillna("Général").astype(str).str.strip()
    df.loc[df["__area_list__"] == "", "__area_list__"] = "Général"
    return df


def reminders_for_project(
    project_title: str,
    ref_date: date,
    max_level: int = 8,
    start_date: Optional[date] = None,
    end_date: Optional[date] = None,
) -> pd.DataFrame:
    e = get_entries().copy()
    e = e.loc[e[E_COL_PROJECT_TITLE].fillna("").astype(str).str.strip() == project_title].copy()
    e = _filter_entries_by_created_range(e, start_date, end_date)

    e["__is_task__"] = _series(e, E_COL_IS_TASK, False).apply(_bool_true)
    e = e.loc[e["__is_task__"] == True].copy()

    e["__completed__"] = _series(e, E_COL_COMPLETED, False).apply(_bool_true)
    e["__done__"] = _series(e, E_COL_COMPLETED_END, None).apply(_parse_date_any)
    e.loc[e["__done__"].notna(), "__completed__"] = True
    e["__deadline__"] = _series(e, E_COL_DEADLINE, None).apply(_parse_date_any)
    e["__reminder__"] = e.apply(lambda r: reminder_level(r["__deadline__"], r["__completed__"], ref_date), axis=1)

    e = e.loc[e["__reminder__"].notna()].copy()
    e["__reminder__"] = e["__reminder__"].astype(int)
    e = e.loc[e["__reminder__"] <= max_level].copy()

    e = _explode_areas(e)

    e["__company__"] = _series(e, E_COL_COMPANY_TASK, "").fillna("").astype(str).str.strip()
    e.loc[e["__company__"] == "", "__company__"] = "Non renseigné"

    e = e.sort_values(["__reminder__", "__deadline__"], ascending=[False, True])
    return e


def followups_for_project(
    project_title: str,
    ref_date: date,
    exclude_entry_ids: set[str],
    start_date: Optional[date] = None,
    end_date: Optional[date] = None,
) -> pd.DataFrame:
    """À suivre = tâches non clôturées NON en retard à ref_date (deadline >= ref_date ou deadline vide)."""
    e = get_entries().copy()
    e = e.loc[e[E_COL_PROJECT_TITLE].fillna("").astype(str).str.strip() == project_title].copy()
    e = _filter_entries_by_created_range(e, start_date, end_date)

    e["__id__"] = _series(e, E_COL_ID, "").fillna("").astype(str).str.strip()
    if exclude_entry_ids:
        e = e.loc[~e["__id__"].isin(exclude_entry_ids)].copy()

    e["__is_task__"] = _series(e, E_COL_IS_TASK, False).apply(_bool_true)
    e = e.loc[e["__is_task__"] == True].copy()

    e["__completed__"] = _series(e, E_COL_COMPLETED, False).apply(_bool_true)
    e["__done__"] = _series(e, E_COL_COMPLETED_END, None).apply(_parse_date_any)
    e.loc[e["__done__"].notna(), "__completed__"] = True
    e = e.loc[e["__completed__"] == False].copy()

    e["__deadline__"] = _series(e, E_COL_DEADLINE, None).apply(_parse_date_any)
    e = e.loc[e["__deadline__"].isna() | (e["__deadline__"] >= ref_date)].copy()

    e = _explode_areas(e)

    e["__company__"] = _series(e, E_COL_COMPANY_TASK, "").fillna("").astype(str).str.strip()
    e.loc[e["__company__"] == "", "__company__"] = "Non renseigné"

    e["__deadline_sort__"] = e["__deadline__"].apply(lambda d: date(2999, 12, 31) if d is None else d)
    e = e.sort_values(["__deadline_sort__", "__company__"], ascending=[True, True])
    return e


def reminders_by_company(rem_df: pd.DataFrame) -> List[Dict]:
    if rem_df.empty:
        return []
    logo_map = companies_logo_by_name()
    g = rem_df.groupby("__company__", dropna=False).size().reset_index(name="count")
    g["__norm__"] = g["__company__"].astype(str).apply(_norm_name)
    g["logo"] = g["__norm__"].apply(lambda k: logo_map.get(k, ""))
    g = g.sort_values("count", ascending=False)
    out = []
    for _, r in g.iterrows():
        out.append({"name": str(r["__company__"]), "logo": str(r["logo"] or "").strip(), "count": int(r["count"])})
    return out


# -------------------------
# ZONES (for meeting entries)
# -------------------------
def group_meeting_by_area(edf: pd.DataFrame) -> List[Tuple[str, pd.DataFrame]]:
    df = edf.copy()
    df = _explode_areas(df)
    areas: List[Tuple[str, pd.DataFrame]] = []
    for area, g in df.groupby("__area_list__", sort=True):
        areas.append((str(area), g.copy()))
    areas.sort(key=lambda x: (0 if x[0].lower() == "général" else 1, x[0].lower()))
    return areas


# -------------------------
# MEMO MODAL (UI)
# -------------------------
EDITOR_MEMO_MODAL_CSS = r"""
.btnAddMemo{margin-left:auto; font-size:12px; padding:6px 10px; border:1px solid #ddd; border-radius:10px; background:#fff; cursor:pointer}
.btnAddMemo:hover{background:#f7f7f7}
.memoModal{position:fixed; inset:0; padding:16px 16px 16px 290px; background:rgba(0,0,0,.35); display:none; align-items:flex-start; justify-content:center; overflow:auto; z-index:9999}
.memoModal .panel{background:#fff; width:min(720px, calc(100vw - 330px)); max-height:calc(100vh - 32px); overflow:auto; border-radius:14px; box-shadow:0 20px 60px rgba(0,0,0,.25)}
@media (max-width:1200px){.memoModal{padding:16px}.memoModal .panel{width:min(720px,94vw)}}
.memoModal .head{display:flex; gap:12px; align-items:center; padding:14px 16px; border-bottom:1px solid #eee}
.memoModal .list{padding:10px 16px}
.memoModal .item{display:block; padding:10px 10px; border:1px solid #eee; border-radius:12px; margin:8px 0}
.memoModal .item:hover{background:#fafafa}
.memoModal .actions{display:flex; gap:10px; justify-content:flex-end; padding:12px 16px; border-top:1px solid #eee}
.memoBtn{padding:8px 12px; border:1px solid #ddd; background:#fff; border-radius:10px; cursor:pointer}
.memoBtnPrimary{border-color:#111; background:#111; color:#fff}
"""

EDITOR_MEMO_MODAL_HTML = r"""
<div class="memoModal" id="memoModal">
  <div class="panel">
    <div class="head">
      <h3 id="memoModalTitle" style="margin:0">Ajouter des mémos</h3>
      <span class="muted" id="memoModalSub"></span>
      <div style="margin-left:auto"></div>
      <button class="memoBtn" id="memoModalClose" type="button">Fermer</button>
    </div>
    <div class="list" id="memoModalList"></div>
    <div class="actions">
      <button class="memoBtn" id="memoModalCancel" type="button">Annuler</button>
      <button class="memoBtn memoBtnPrimary" id="memoModalAdd" type="button">Ajouter</button>
    </div>
  </div>
</div>
"""

EDITOR_MEMO_MODAL_JS = r"""
(function(){
  const qs = (k) => new URLSearchParams(window.location.search).get(k) || "";
  const modal = document.getElementById('memoModal');
  if(!modal) return;
  const listEl = document.getElementById('memoModalList');
  const subEl = document.getElementById('memoModalSub');
  let currentArea = "";

  function open(area){
    currentArea = area;
    subEl.textContent = "Zone : " + area;
    listEl.innerHTML = "<div class='muted'>Chargement…</div>";
    modal.style.display = "flex";
    const project = qs("project") || "";
    fetch(`/api/memos?project=${encodeURIComponent(project)}&area=${encodeURIComponent(area)}`)
      .then(r => r.json())
      .then(data => {
        const pinned = (qs("pinned_memos")||"").split(",").map(s=>s.trim()).filter(Boolean);
        if(!data || !data.items || data.items.length===0){
          listEl.innerHTML = "<div class='muted'>Aucun mémo disponible pour cette zone.</div>";
          return;
        }
        listEl.innerHTML = data.items.map(it => {
          const checked = pinned.includes(it.id) ? "checked" : "";
          const meta = [it.created||"", it.company||"", it.owner||""].filter(Boolean).join(" • ");
          return `<label class="item">
            <div style="display:flex; gap:10px; align-items:flex-start">
              <input type="checkbox" data-id="${it.id}" ${checked} style="margin-top:3px"/>
              <div>
                <div style="font-weight:800">${it.title||"(Sans titre)"}</div>
                <div class="muted" style="margin-top:2px">${meta}</div>
              </div>
            </div>
          </label>`;
        }).join("");
      })
      .catch(()=>{ listEl.innerHTML = "<div class='muted'>Erreur de chargement.</div>"; });
  }

  function close(){ modal.style.display = "none"; currentArea = ""; }
  document.getElementById('memoModalClose').onclick = close;
  document.getElementById('memoModalCancel').onclick = close;
  modal.addEventListener('click', (e)=>{ if(e.target===modal) close(); });

  document.getElementById('memoModalAdd').onclick = function(){
    const ids = Array.from(listEl.querySelectorAll("input[type=checkbox][data-id]"))
      .filter(x=>x.checked).map(x=>x.getAttribute("data-id")).filter(Boolean);
    const u = new URL(window.location.href);
    const existing = (u.searchParams.get("pinned_memos")||"").split(",").map(s=>s.trim()).filter(Boolean);
    const merged = Array.from(new Set(existing.concat(ids))).join(",");
    if(merged) u.searchParams.set("pinned_memos", merged);
    else u.searchParams.delete("pinned_memos");
    window.location.href = u.toString();
  };

  document.addEventListener("click", (e) => {
    const btn = e.target.closest(".btnAddMemo");
    if(!btn) return;
    open(btn.getAttribute("data-area")||"");
  });
})();
"""

QUALITY_MODAL_CSS = r"""
.qualityModal{position:fixed; inset:0; padding:16px 16px 16px 290px; background:rgba(0,0,0,.35); display:none; align-items:flex-start; justify-content:center; overflow:auto; z-index:9998}
.qualityModal .panel{background:#fff; width:min(980px, calc(100vw - 330px)); max-height:calc(100vh - 32px); overflow:auto; border-radius:16px; box-shadow:0 20px 60px rgba(0,0,0,.25)}
@media (max-width:1200px){.qualityModal{padding:16px}.qualityModal .panel{width:min(980px,94vw)}}
.qualityModal .head{display:flex; gap:12px; align-items:center; padding:16px 18px; border-bottom:1px solid #eee}
.qualityModal .list{padding:14px 18px}
.qualityModal .item{border:1px solid #e2e8f0; border-radius:14px; padding:12px; margin:10px 0; background:#fff}
.qualityModal .meta{color:#475569; font-weight:700; font-size:12px}
.qualityScore{font-size:28px; font-weight:1000}
.qualityBadge{display:inline-flex; align-items:center; gap:8px; padding:6px 10px; border-radius:999px; background:#fff1f2; border:1px solid #fecdd3; font-weight:900; color:#b91c1c}
.qualityGrid{display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin-top:10px}
.qualityCard{border:1px solid #e2e8f0; border-radius:12px; padding:10px; background:#f8fafc}
.qualityHighlight{background:#fee2e2; padding:0 4px; border-radius:4px; font-weight:900; color:#b91c1c; position:relative; cursor:help}
.qualityHighlight:hover::after{content:attr(data-suggestion); position:absolute; left:0; top:100%; margin-top:6px; background:#111827; color:#fff; padding:6px 8px; border-radius:6px; font-size:11px; white-space:pre-wrap; z-index:20; min-width:140px; max-width:240px}
.qualityHighlight:hover::before{content:""; position:absolute; left:10px; top:100%; border:6px solid transparent; border-bottom-color:#111827}
.qualityFullText{margin-top:6px; line-height:1.4}
.qualityTips{border-left:4px solid #b91c1c; padding:10px 12px; background:#fff1f2; border-radius:10px; margin-top:12px}
.qualityItemTitle{color:#b91c1c; font-weight:900}
"""

QUALITY_MODAL_HTML = r"""
<div class="qualityModal" id="qualityModal">
  <div class="panel">
    <div class="head">
      <h3 style="margin:0">Qualité orthographique &amp; grammaticale</h3>
      <div style="margin-left:auto"></div>
      <button class="memoBtn" id="qualityModalClose" type="button">Fermer</button>
    </div>
    <div class="list" id="qualityModalList"></div>
  </div>
</div>
"""

QUALITY_MODAL_JS = r"""
(function(){
  const modal = document.getElementById('qualityModal');
  const listEl = document.getElementById('qualityModalList');
  if(!modal || !listEl) return;

  function open(){
    listEl.innerHTML = "<div class='muted'>Analyse en cours…</div>";
    modal.style.display = "flex";
    const qs = new URLSearchParams(window.location.search);
    const meetingId = qs.get("meeting_id") || "";
    const project = qs.get("project") || "";
    fetch(`/api/quality?meeting_id=${encodeURIComponent(meetingId)}&project=${encodeURIComponent(project)}`)
      .then(r => r.json())
      .then(data => {
        if(data.error){
          listEl.innerHTML = `<div class='muted'>${data.error}</div>`;
          return;
        }
        const score = data.score ?? 0;
        const total = data.total ?? 0;
        const issuesByArea = data.issues_by_area || {};
        const issueAreas = Object.keys(issuesByArea);
        const strengths = [
          score >= 95 ? "Très bonne qualité générale." : "Qualité perfectible, corrections recommandées.",
          total === 0 ? "Aucune faute détectée." : "Des corrections sont nécessaires.",
          "Objectif : un texte clair et professionnel."
        ];
        const summary = `
          <div class="qualityBadge">Score: <span class="qualityScore">${score}</span>/100</div>
          <div class="qualityGrid">
            <div class="qualityCard"><div class="meta">Erreurs détectées</div><div style="font-weight:900;font-size:18px">${total}</div></div>
            <div class="qualityCard"><div class="meta">Impact</div><div style="font-weight:900;font-size:18px">${score >= 90 ? "Faible" : score >= 75 ? "Moyen" : "Fort"}</div></div>
            <div class="qualityCard"><div class="meta">Relecture</div><div style="font-weight:900;font-size:18px">${score >= 90 ? "OK" : "Recommandée"}</div></div>
          </div>
          <div class="qualityTips">
            <div style="font-weight:900">Conseils pédagogiques</div>
            <ul style="margin:6px 0 0 16px">
              <li>${strengths[0]}</li>
              <li>${strengths[1]}</li>
              <li>${strengths[2]}</li>
            </ul>
            <div class="meta" style="margin-top:6px">Corrige les libellés directement dans METRONOME pour améliorer la qualité globale.</div>
          </div>
        `;
        if(!issueAreas.length){
          listEl.innerHTML = summary + "<div class='muted' style='margin-top:10px'>Aucune faute détectée.</div>";
          return;
        }
        const escapeHtml = (v) => String(v || "").replace(/[&<>"']/g, (m) => ({
          "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"
        })[m]);
        const sections = issueAreas.map(area => {
          const items = (issuesByArea[area] || []).map(it => {
            const text = it.text || it.context || "";
            const offset = it.offset ?? it.context_offset;
            const length = it.length ?? it.context_length;
            const suggestion = it.replacements || it.message || "Suggestion";
            let highlight = escapeHtml(text);
            if(text && offset != null && length != null){
              const safeText = escapeHtml(text);
              const before = safeText.slice(0, offset);
              const mid = safeText.slice(offset, offset + length);
              const after = safeText.slice(offset + length);
              highlight = `${before}<span class="qualityHighlight" data-suggestion="${escapeHtml(suggestion)}">${mid}</span>${after}`;
            }
            return `
              <div class="item">
                <div class="qualityItemTitle">${escapeHtml(it.category || "Suggestion")}</div>
                <div class="qualityFullText">${highlight || "—"}</div>
              </div>
            `;
          }).join("");
          return `
            <div style="margin-top:16px;font-weight:900">Zone : ${escapeHtml(area)}</div>
            ${items}
          `;
        }).join("");
        listEl.innerHTML = summary + "<div style='margin-top:12px;font-weight:900'>Points à corriger</div>" + sections;
      })
      .catch(() => {
        listEl.innerHTML = "<div class='muted'>Impossible d'analyser pour le moment.</div>";
      });
  }

  function close(){ modal.style.display = "none"; }
  document.getElementById('qualityModalClose').onclick = close;
  modal.addEventListener('click', (e)=>{ if(e.target===modal) close(); });
  document.getElementById('btnQualityCheck')?.addEventListener('click', open);
})();
"""

ANALYSIS_MODAL_CSS = r"""
.analysisModal{position:fixed; inset:0; padding:16px 16px 16px 290px; background:rgba(0,0,0,.35); display:none; align-items:flex-start; justify-content:center; overflow:auto; z-index:9997}
.analysisModal .panel{background:#fff; width:min(980px, calc(100vw - 330px)); max-height:calc(100vh - 32px); overflow:auto; border-radius:16px; box-shadow:0 20px 60px rgba(0,0,0,.25)}
@media (max-width:1200px){.analysisModal{padding:16px}.analysisModal .panel{width:min(980px,94vw)}}
.analysisModal .head{display:flex; gap:12px; align-items:center; padding:16px 18px; border-bottom:1px solid #eee}
.analysisModal .list{padding:14px 18px}
.analysisCard{border:1px solid #e2e8f0; border-radius:14px; padding:12px; margin:10px 0; background:#fff}
.analysisGrid{display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin-top:10px}
.analysisKpi{border:1px solid #e2e8f0; border-radius:12px; padding:10px; background:#f8fafc}
"""

ANALYSIS_MODAL_HTML = r"""
<div class="analysisModal" id="analysisModal">
  <div class="panel">
    <div class="head">
      <h3 style="margin:0">Analyse du compte rendu</h3>
      <div style="margin-left:auto"></div>
      <button class="memoBtn" id="analysisModalClose" type="button">Fermer</button>
    </div>
    <div class="list" id="analysisModalList"></div>
  </div>
</div>
"""

ANALYSIS_MODAL_JS = r"""
(function(){
  const modal = document.getElementById('analysisModal');
  const listEl = document.getElementById('analysisModalList');
  if(!modal || !listEl) return;

  function open(){
    listEl.innerHTML = "<div class='muted'>Analyse en cours…</div>";
    modal.style.display = "flex";
    const qs = new URLSearchParams(window.location.search);
    const meetingId = qs.get("meeting_id") || "";
    const project = qs.get("project") || "";
    fetch(`/api/analysis?meeting_id=${encodeURIComponent(meetingId)}&project=${encodeURIComponent(project)}`)
      .then(r => r.json())
      .then(data => {
        if(data.error){
          listEl.innerHTML = `<div class='muted'>${data.error}</div>`;
          return;
        }
        const k = data.kpis || {};
        const bullets = (data.points || []).map(p => `<li>${p}</li>`).join("");
        const risks = (data.risks || []).map(p => `<li>${p}</li>`).join("");
        const follow = (data.follow_ups || []).map(p => `<li>${p}</li>`).join("");
        const least = (data.least_responsive || []).map(it => `<li>${it.name} (${it.count})</li>`).join("");
        const byArea = data.followups_by_area || {};
        const areaSections = Object.keys(byArea).map(a => {
          const items = (byArea[a] || []).map(t => `<li>${t}</li>`).join("");
          return `<div class="analysisCard"><div style="font-weight:900">Zone : ${a}</div><ul style="margin:6px 0 0 18px">${items || "<li>Aucune action prioritaire.</li>"}</ul></div>`;
        }).join("");
        listEl.innerHTML = `
          <div class="analysisGrid">
            <div class="analysisKpi"><div class="meta">Rappels en retard</div><div style="font-weight:900;font-size:18px">${k.late_tasks ?? 0}</div></div>
            <div class="analysisKpi"><div class="meta">Tâches ouvertes</div><div style="font-weight:900;font-size:18px">${k.open_tasks ?? 0}</div></div>
            <div class="analysisKpi"><div class="meta">À suivre</div><div style="font-weight:900;font-size:18px">${k.followups ?? 0}</div></div>
          </div>
          <div class="analysisCard">
            <div style="font-weight:900">Synthèse rapide</div>
            <ul style="margin:6px 0 0 18px">${bullets || "<li>Aucun point marquant détecté.</li>"}</ul>
          </div>
          <div class="analysisCard">
            <div style="font-weight:900">Points de vigilance</div>
            <ul style="margin:6px 0 0 18px">${risks || "<li>Aucun risque majeur identifié.</li>"}</ul>
          </div>
          <div class="analysisCard">
            <div style="font-weight:900">À relancer à la prochaine réunion</div>
            <ul style="margin:6px 0 0 18px">${follow || "<li>Rien de critique à relancer.</li>"}</ul>
          </div>
          <div class="analysisCard">
            <div style="font-weight:900">Entreprises les moins réactives</div>
            <ul style="margin:6px 0 0 18px">${least || "<li>Aucune entreprise à relancer en priorité.</li>"}</ul>
          </div>
          <div style="margin-top:12px;font-weight:900">Actions attendues par zone</div>
          ${areaSections || "<div class='analysisCard'>Aucune action par zone.</div>"}
        `;
      })
      .catch(() => {
        listEl.innerHTML = "<div class='muted'>Impossible d'analyser pour le moment.</div>";
      });
  }

  function close(){ modal.style.display = "none"; }
  document.getElementById('analysisModalClose').onclick = close;
  modal.addEventListener('click', (e)=>{ if(e.target===modal) close(); });
  document.getElementById('btnAnalysis')?.addEventListener('click', open);
})();
"""

RESIZE_TOP_JS = r"""
(function(){
  const root = document.documentElement;
  const grip = document.getElementById('topPageGrip');
  if(!grip) return;
  let startX = 0;
  let startScale = 1;
  function onMove(e){
    const dx = e.clientX - startX;
    const next = Math.max(0.8, Math.min(1.1, startScale + dx / 500));
    root.style.setProperty('--top-scale', next.toFixed(2));
  }
  function onUp(){
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
  }
  grip.addEventListener('mousedown', (e) => {
    startX = e.clientX;
    const current = parseFloat(getComputedStyle(root).getPropertyValue('--top-scale').trim() || '1');
    startScale = current;
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  });
})();
"""
RESIZE_COLUMNS_JS = r"""
(function(){
  const root = document.documentElement;
  const map = {
    type: '--col-type',
    comment: '--col-comment',
    date: '--col-date',
    date2: '--col-date',
    date3: '--col-date',
    lot: '--col-lot',
    who: '--col-who',
  };
  let active = null;
  let startX = 0;
  let startPct = 0;
  function onMove(e){
    if(!active) return;
    const table = active.closest('table');
    const width = table.getBoundingClientRect().width || 1;
    const dx = e.clientX - startX;
    const deltaPct = (dx / width) * 100;
    const next = Math.max(3, startPct + deltaPct);
    root.style.setProperty(map[active.dataset.col], `${next}%`);
  }
  function onUp(){
    active = null;
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
  }
  document.addEventListener('mousedown', (e) => {
    const grip = e.target.closest('.colGrip');
    if(!grip) return;
    active = grip;
    startX = e.clientX;
    const current = getComputedStyle(root).getPropertyValue(map[grip.dataset.col]).trim().replace('%','');
    startPct = parseFloat(current || '0');
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  });
})();
"""

SYNC_EDITABLE_JS = r"""
(function(){
  function syncAll(){
    const groups = new Map();
    document.querySelectorAll('[data-sync]').forEach(el => {
      const key = el.getAttribute('data-sync') || '';
      if(!key || groups.has(key)) return;
      groups.set(key, el.textContent);
    });
    groups.forEach((value, key) => {
      document.querySelectorAll(`[data-sync="${key}"]`).forEach(el => {
        if(el.textContent !== value){ el.textContent = value; }
      });
    });
  }

  function syncValue(el){
    const key = el.getAttribute('data-sync') || '';
    if(!key) return;
    const value = el.textContent;
    document.querySelectorAll(`[data-sync="${key}"]`).forEach(target => {
      if(target !== el){ target.textContent = value; }
    });
  }

  document.addEventListener('input', (e) => {
    const el = e.target.closest('[data-sync]');
    if(el){ syncValue(el); }
  });
  document.addEventListener('blur', (e) => {
    const el = e.target.closest('[data-sync]');
    if(el){ syncValue(el); }
  }, true);
  window.addEventListener('DOMContentLoaded', syncAll);
})();
"""

RANGE_PICKER_JS = r"""
function toggleRangePanel(){
  const panel = document.getElementById('rangePanel');
  if(!panel){ return; }
  const current = panel.style.display;
  panel.style.display = (!current || current === 'none') ? 'flex' : 'none';
}

function applyRange(){
  const start = document.getElementById('rangeStart')?.value || "";
  const end = document.getElementById('rangeEnd')?.value || "";
  const url = new URL(window.location.href);
  if(start){ url.searchParams.set('range_start', start); }
  else{ url.searchParams.delete('range_start'); }
  if(end){ url.searchParams.set('range_end', end); }
  else{ url.searchParams.delete('range_end'); }
  window.location.href = url.toString();
}

function clearRange(){
  const startEl = document.getElementById('rangeStart');
  const endEl = document.getElementById('rangeEnd');
  if(startEl){ startEl.value = ""; }
  if(endEl){ endEl.value = ""; }
  const url = new URL(window.location.href);
  url.searchParams.delete('range_start');
  url.searchParams.delete('range_end');
  window.location.href = url.toString();
}

window.addEventListener('DOMContentLoaded', () => {
  document.getElementById('btnRange')?.addEventListener('click', toggleRangePanel);
});

document.addEventListener('click', (e) => {
  const btn = e.target.closest('#btnRange');
  if(!btn) return;
  toggleRangePanel();
});
"""

PRINT_PREVIEW_TOGGLE_JS = r"""
(function(){
  const btn = document.getElementById('btnPrintPreview');
  if(!btn) return;
  const STORAGE_KEY = 'tempo.print.preview.enabled.v1';

  function loadState(){
    try{ return localStorage.getItem(STORAGE_KEY) === '1'; }
    catch(_){ return false; }
  }

  function saveState(v){
    try{ localStorage.setItem(STORAGE_KEY, v ? '1' : '0'); }
    catch(_){ }
  }

  function apply(enabled){
    document.body.classList.toggle('printPreviewMode', enabled);
    document.body.classList.toggle('printOptimized', enabled);
    btn.textContent = enabled ? 'Aperçu impression : ON' : 'Aperçu impression : OFF';
    btn.classList.toggle('active', enabled);
    if(window.repaginateReport){ window.repaginateReport(); }
  }

  let enabled = loadState();
  apply(enabled);

  btn.addEventListener('click', () => {
    enabled = !enabled;
    saveState(enabled);
    apply(enabled);
  });
})();
"""

CONSTRAINT_TOGGLES_JS = r"""
(function(){
  const panel = document.getElementById('constraintsPanel');
  if(!panel) return;
  const root = document.documentElement;
  const body = document.body;
  const STORAGE_KEY = 'tempo.constraint.toggles.v1';

  const defaultState = {
    fixedA4: true,
    fixedPageHeight: true,
    pageBreaks: true,
    bodyOffset: true,
    pagePadding: true,
    footerReserve: true,
    tableFixed: true,
    printHideUi: true,
    printStickyHeader: true,
    printCompactRows: true,
    printAvoidSplitRows: true,
    keepSessionHeaderWithNext: true,
    printAutoOptimize: true,
    topScale: true,
  };

  function loadState(){
    try{
      const raw = localStorage.getItem(STORAGE_KEY);
      const parsed = raw ? JSON.parse(raw) : {};
      return {...defaultState, ...parsed};
    }catch(_){
      return {...defaultState};
    }
  }

  function saveState(state){
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }

  function applyConstraint(name, active){
    body.classList.toggle(`constraint-off-${name}`, !active);
  }

  function applyAll(state){
    Object.entries(state).forEach(([k, v]) => applyConstraint(k, !!v));
  }

  const state = loadState();
  panel.querySelectorAll('[data-constraint]').forEach(input => {
    const name = input.getAttribute('data-constraint');
    if(!(name in state)) return;
    input.checked = !!state[name];
    input.addEventListener('change', () => {
      state[name] = !!input.checked;
      applyConstraint(name, state[name]);
      saveState(state);
      if(window.repaginateReport){ window.repaginateReport(); }
    });
  });

  function updateFooterReserveFactor(){
    const input = document.getElementById('footerReserveFactor');
    const value = document.getElementById('footerReserveFactorValue');
    if(!input || !value) return;
    const pct = Math.max(-100, Math.min(150, parseFloat(input.value || '100')));
    const factor = pct / 100;
    value.textContent = `${Math.round(pct)} %`;
    root.style.setProperty('--footer-reserve-factor', factor.toFixed(2));
    try{ localStorage.setItem('tempo.footer.reserve.factor.v1', String(Math.round(pct))); }catch(_){ }
    if(window.repaginateReport){ window.repaginateReport(); }
  }

  const footerReserveInput = document.getElementById('footerReserveFactor');
  if(footerReserveInput){
    let savedPct = null;
    try{ savedPct = localStorage.getItem('tempo.footer.reserve.factor.v1'); }catch(_){ }
    if(savedPct !== null && savedPct !== ''){ footerReserveInput.value = savedPct; }
    footerReserveInput.addEventListener('input', updateFooterReserveFactor);
  }

  document.getElementById('btnConstraints')?.addEventListener('click', () => {
    panel.style.display = panel.style.display === 'none' ? 'flex' : 'none';
  });

  applyAll(state);
  updateFooterReserveFactor();
})();
"""

LAYOUT_CONTROLS_JS = r"""
(function(){
  function closestZone(el){ return el.closest('.zoneBlock'); }
  function move(zone, dir){
    if(!zone) return;
    if(dir === 'up'){
      const prev = zone.previousElementSibling;
      if(prev && prev.classList.contains('zoneBlock')){
        zone.parentNode.insertBefore(zone, prev);
      }
    }else if(dir === 'down'){
      const next = zone.nextElementSibling;
      if(next && next.classList.contains('zoneBlock')){
        zone.parentNode.insertBefore(next, zone);
      }
    }
    if(window.repaginateReport){ window.repaginateReport(); }
  }
  document.addEventListener('click', (e) => {
    const btn = e.target.closest('.zoneBtn');
    if(!btn) return;
    const action = btn.dataset.action || '';
    const zone = closestZone(btn);
    if(!zone) return;
    if(action === 'highlight'){
      zone.classList.toggle('highlight');
    }else if(action === 'move-up'){
      move(zone, 'up');
    }else if(action === 'move-down'){
      move(zone, 'down');
    }
  });
})();
"""


DRAGGABLE_IMAGES_JS = r"""
(function(){
  function ensureThumbWrapper(imgSrc){
    return `<span class="thumbAWrap" data-thumb draggable="true"><a class="thumbA" href="${imgSrc}" target="_blank" rel="noopener"><img class="thumb" src="${imgSrc}" alt="" /></a><button type="button" class="thumbRemove noPrint" title="Supprimer">×</button><span class="thumbHandle" title="Redimensionner"></span></span>`;
  }

  function attachResizeBehavior(wrap){
    if(!wrap || wrap.dataset.resizeReady === '1') return;
    wrap.dataset.resizeReady = '1';
    const handle = wrap.querySelector('.thumbHandle');
    const img = wrap.querySelector('.thumb');
    if(!handle || !img) return;

    handle.addEventListener('pointerdown', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const startX = e.clientX;
      const startWidth = img.getBoundingClientRect().width || 160;
      wrap.classList.add('resizing');

      const onMove = (ev) => {
        const nextWidth = Math.min(520, Math.max(70, startWidth + (ev.clientX - startX)));
        img.style.width = `${nextWidth}px`;
        img.style.height = 'auto';
      };

      const onUp = () => {
        wrap.classList.remove('resizing');
        document.removeEventListener('pointermove', onMove);
        document.removeEventListener('pointerup', onUp);
      };

      document.addEventListener('pointermove', onMove);
      document.addEventListener('pointerup', onUp);
    });
  }

  function initGallery(gallery){
    if(!gallery) return;
    gallery.querySelectorAll('.thumbAWrap').forEach(wrap => {
      wrap.setAttribute('draggable', 'true');
      attachResizeBehavior(wrap);
    });
    if(gallery.dataset.dragReady === '1') return;
    gallery.dataset.dragReady = '1';

    let dragEl = null;

    gallery.addEventListener('dragstart', (e) => {
      if(e.target.closest('.thumbHandle')){
        e.preventDefault();
        return;
      }
      const wrap = e.target.closest('.thumbAWrap');
      if(!wrap || wrap.classList.contains('resizing')) return;
      dragEl = wrap;
      wrap.classList.add('dragging');
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', 'thumb');
    });

    gallery.addEventListener('dragend', () => {
      if(dragEl){ dragEl.classList.remove('dragging'); }
      dragEl = null;
    });

    gallery.addEventListener('dragover', (e) => {
      if(!dragEl) return;
      e.preventDefault();
      const over = e.target.closest('.thumbAWrap');
      if(!over || over === dragEl) return;
      const rect = over.getBoundingClientRect();
      const before = e.clientX < (rect.left + rect.width / 2);
      gallery.insertBefore(dragEl, before ? over : over.nextSibling);
    });

    gallery.addEventListener('click', (e) => {
      const removeBtn = e.target.closest('.thumbRemove');
      if(!removeBtn) return;
      const wrap = removeBtn.closest('.thumbAWrap');
      if(wrap){ wrap.remove(); }
    });
  }

  function ensureRowGallery(cell){
    let gallery = cell.querySelector('.thumbs[data-gallery]');
    if(!gallery){
      gallery = document.createElement('div');
      gallery.className = 'thumbs';
      gallery.setAttribute('data-gallery', '1');
      const comment = cell.querySelector('.commentText');
      if(comment && comment.nextSibling){
        comment.parentNode.insertBefore(gallery, comment.nextSibling);
      }else{
        cell.appendChild(gallery);
      }
    }
    initGallery(gallery);
    return gallery;
  }

  function setupImageButtons(){
    document.querySelectorAll('.colComment').forEach(cell => {
      const btn = cell.querySelector('.btnAddImage');
      const input = cell.querySelector('.imageInput');
      if(!btn || !input || btn.dataset.ready === '1') return;
      btn.dataset.ready = '1';
      btn.addEventListener('click', () => input.click());
      input.addEventListener('change', (e) => {
        const files = Array.from(e.target.files || []).filter(f => f.type.startsWith('image/'));
        if(!files.length) return;
        const gallery = ensureRowGallery(cell);
        files.forEach(file => {
          const reader = new FileReader();
          reader.onload = () => {
            const src = String(reader.result || '');
            if(!src) return;
            gallery.insertAdjacentHTML('beforeend', ensureThumbWrapper(src));
            const inserted = gallery.lastElementChild;
            if(inserted){ attachResizeBehavior(inserted); }
          };
          reader.readAsDataURL(file);
        });
        input.value = '';
      });
    });
  }

  window.enableDraggableThumbs = function(){
    document.querySelectorAll('.thumbs').forEach(initGallery);
    setupImageButtons();
  };

  window.addEventListener('load', () => {
    window.enableDraggableThumbs();
  });
})();
"""


PRINT_OPTIMIZE_JS = r"""
(function(){
  function optimizeWhitespaceForPrint(){
    if(document.body.classList.contains('constraint-off-printAutoOptimize')){ return; }
    if(document.body.classList.contains('printPreviewMode')){ return; }
    document.body.classList.add('printOptimized');
    if(window.repaginateReport){
      window.repaginateReport();
    }
  }
  function restoreAfterPrint(){
    if(document.body.classList.contains('constraint-off-printAutoOptimize')){ return; }
    if(document.body.classList.contains('printPreviewMode')){ return; }
    document.body.classList.remove('printOptimized');
    if(window.repaginateReport){
      window.repaginateReport();
    }
  }

  window.addEventListener('beforeprint', optimizeWhitespaceForPrint);
  window.addEventListener('afterprint', restoreAfterPrint);
})();
"""

PAGINATION_JS = r"""
(function(){
  function px(value){
    const n = parseFloat(value || "0");
    return Number.isNaN(n) ? 0 : n;
  }

  function calcAvailable(page, includePresence){
    const pageContent = page.querySelector('.pageContent');
    const footer = page.querySelector('.docFooter');
    const header = page.querySelector('.reportHeader');
    const presence = page.querySelector('.presenceWrap');
    const pageRect = page.getBoundingClientRect();
    if(!pageContent) return pageRect.height;
    const styles = window.getComputedStyle(pageContent);
    let available = pageRect.height - px(styles.paddingTop) - px(styles.paddingBottom);
    const reserveFooter = !document.body.classList.contains('constraint-off-footerReserve');
    const rootStyles = getComputedStyle(document.documentElement);
    const reserveFactorRaw = parseFloat((rootStyles.getPropertyValue('--footer-reserve-factor') || '1').trim());
    const reserveFactor = Number.isNaN(reserveFactorRaw) ? 1 : reserveFactorRaw;
    if(reserveFooter && footer){ available -= (footer.getBoundingClientRect().height * reserveFactor); }
    if(header){ available -= header.getBoundingClientRect().height; }
    if(includePresence && presence){ available -= presence.getBoundingClientRect().height; }
    return available;
  }

  function clearExtraPages(container){
    const pages = Array.from(container.querySelectorAll('.page--report'));
    pages.slice(1).forEach(page => page.remove());
  }

  function mergeZoneBlocks(container){
    const zones = Array.from(container.querySelectorAll('.zoneBlock'));
    const grouped = new Map();
    zones.forEach(zone => {
      const key = zone.getAttribute('data-zone-id') || '';
      if(!grouped.has(key)){ grouped.set(key, []); }
      grouped.get(key).push(zone);
    });
    grouped.forEach(group => {
      if(group.length < 2){ return; }
      const target = group[0];
      const targetBody = target.querySelector('tbody');
      if(!targetBody){ return; }
      group.slice(1).forEach(zone => {
        const body = zone.querySelector('tbody');
        if(body){
          Array.from(body.children).forEach(row => targetBody.appendChild(row));
        }
        zone.remove();
      });
    });
  }

  function getZoneSplitData(zone){
    const title = zone.querySelector('.zoneTitle');
    const table = zone.querySelector('table.crTable');
    const tbody = table?.querySelector('tbody');
    const rows = tbody ? Array.from(tbody.children) : [];
    const rowHeights = rows.map(row => row.getBoundingClientRect().height || row.offsetHeight || 0);
    const tableRect = table?.getBoundingClientRect().height || table?.offsetHeight || 0;
    const rowsSum = rowHeights.reduce((sum, h) => sum + h, 0);
    const tableOverhead = Math.max(0, tableRect - rowsSum);
    const titleHeight = title?.getBoundingClientRect().height || title?.offsetHeight || 0;
    return {rows, rowHeights, tableOverhead, titleHeight};
  }

  function cloneZoneShell(zone){
    const clone = zone.cloneNode(true);
    const tbody = clone.querySelector('tbody');
    if(tbody){ tbody.innerHTML = ''; }
    return clone;
  }

  function buildZoneChunk(zone, data, startIndex, maxHeight){
    const {rows, rowHeights, tableOverhead, titleHeight} = data;
    const total = rows.length;
    let height = titleHeight + tableOverhead;
    let endIndex = startIndex;
    while(endIndex < total){
      const rowHeight = rowHeights[endIndex] || 0;
      if(endIndex > startIndex && height + rowHeight > maxHeight){ break; }
      height += rowHeight;
      endIndex += 1;
      if(endIndex === startIndex + 1 && height > maxHeight){ break; }
    }
    const keepSessionHeaderWithNext = !document.body.classList.contains('constraint-off-keepSessionHeaderWithNext');
    if(keepSessionHeaderWithNext && endIndex < total){
      while(endIndex > startIndex + 1 && rows[endIndex - 1]?.classList.contains('sessionSubRow')){
        endIndex -= 1;
      }
    }
    if(endIndex === startIndex && rows[startIndex]?.classList.contains('sessionSubRow') && startIndex + 1 < total){
      endIndex = Math.min(startIndex + 2, total);
    }
    height = titleHeight + tableOverhead;
    for(let i=startIndex;i<endIndex;i++){
      height += rowHeights[i] || 0;
    }
    const chunk = cloneZoneShell(zone);
    const tbody = chunk.querySelector('tbody');
    for(let i=startIndex;i<endIndex;i++){
      tbody.appendChild(rows[i]);
    }
    return {chunk, nextIndex: endIndex, height};
  }

  function updatePageNumbers(){
    const pages = Array.from(document.querySelectorAll('.wrap .page'));
    const total = pages.length;
    pages.forEach((page, idx) => {
      const num = page.querySelector('.pageNum');
      if(!num) return;
      num.textContent = idx === 0 ? '' : `${idx + 1}/${total}`;
    });
  }

  function paginate(){
    const container = document.querySelector('.reportPages');
    const firstPage = container?.querySelector('.page--report');
    if(!container || !firstPage) return;
    const blocksContainer = firstPage.querySelector('.reportBlocks');
    if(!blocksContainer) return;
    mergeZoneBlocks(container);
    const blocks = Array.from(container.querySelectorAll('.reportBlock')).map(block => ({
      node: block,
      height: block.getBoundingClientRect().height || block.offsetHeight || 0,
      splitData: block.classList.contains('zoneBlock') ? getZoneSplitData(block) : null,
    }));

    blocks.forEach(({node}) => node.remove());
    clearExtraPages(container);

    let currentPage = firstPage;
    let currentBlocks = blocksContainer;
    let available = calcAvailable(currentPage, true);
    let used = 0;
    const template = document.getElementById('report-page-template');

    blocks.forEach(({node, height, splitData}) => {
      if(splitData && splitData.rows.length){
        let rowIndex = 0;
        while(rowIndex < splitData.rows.length){
          const remaining = available - used;
          if(remaining <= splitData.titleHeight + splitData.tableOverhead && template && used > 0){
            const clone = template.content.firstElementChild.cloneNode(true);
            container.appendChild(clone);
            currentPage = clone;
            currentBlocks = clone.querySelector('.reportBlocks');
            available = calcAvailable(currentPage, false);
            used = 0;
          }
          const maxHeight = Math.max(available - used, splitData.titleHeight + splitData.tableOverhead);
          const {chunk, nextIndex, height: chunkHeight} = buildZoneChunk(node, splitData, rowIndex, maxHeight);
          if(used > 0 && used + chunkHeight > available && template){
            const clone = template.content.firstElementChild.cloneNode(true);
            container.appendChild(clone);
            currentPage = clone;
            currentBlocks = clone.querySelector('.reportBlocks');
            available = calcAvailable(currentPage, false);
            used = 0;
          }
          currentBlocks.appendChild(chunk);
          const actualHeight = chunk.getBoundingClientRect().height || chunkHeight;
          used += actualHeight;
          rowIndex = nextIndex;
        }
        return;
      }
      if(used > 0 && used + height > available && template){
        const clone = template.content.firstElementChild.cloneNode(true);
        container.appendChild(clone);
        currentPage = clone;
        currentBlocks = clone.querySelector('.reportBlocks');
        available = calcAvailable(currentPage, false);
        used = 0;
      }
      currentBlocks.appendChild(node);
      const actualHeight = node.getBoundingClientRect().height || height;
      used += actualHeight;
    });

    updatePageNumbers();
  }

  window.repaginateReport = paginate;
  window.refreshPagination = function(){
    if(!window.repaginateReport){ return; }
    requestAnimationFrame(() => window.repaginateReport());
  };
  window.addEventListener('load', () => {
    requestAnimationFrame(paginate);
  });
  window.addEventListener('resize', () => {
    clearTimeout(window.__repaginateTimer);
    window.__repaginateTimer = setTimeout(paginate, 200);
  });
})();
"""

ROW_CONTROL_JS = r"""
(function(){
  const hiddenSet = new Set();

  function rowById(id){ return document.querySelector(`tr[data-row-id="${id}"]`); }

  function syncSessionHeaders(){
    document.querySelectorAll('.crTable tbody').forEach(tbody => {
      const rows = Array.from(tbody.querySelectorAll('tr'));
      for(let i=0;i<rows.length;i++){
        const r = rows[i];
        if(!r.classList.contains('sessionSubRow')) continue;
        let hasVisible = false;
        for(let j=i+1;j<rows.length;j++){
          const n = rows[j];
          if(n.classList.contains('sessionSubRow')) break;
          if(n.classList.contains('rowItem') && !n.classList.contains('rowHidden')){ hasVisible = true; break; }
        }
        r.classList.toggle('rowHidden', !hasVisible);
      }
    });
  }

  function syncZoneVisibility(){
    document.querySelectorAll('.zoneBlock').forEach(zone => {
      const visibleItems = zone.querySelectorAll('tr.rowItem:not(.rowHidden)');
      zone.classList.toggle('rowHidden', visibleItems.length === 0);
    });
  }

  function refreshHiddenSelect(){
    const sel = document.getElementById('hiddenRowsSelect');
    if(!sel) return;
    const current = sel.value || "";
    sel.innerHTML = '<option value="">Lignes masquées…</option>';
    Array.from(hiddenSet).sort().forEach(id => {
      const row = rowById(id);
      const title = row ? (row.querySelector('.commentText')?.textContent || id) : id;
      const opt = document.createElement('option');
      opt.value = id;
      opt.textContent = title.trim().slice(0, 90);
      sel.appendChild(opt);
    });
    if(current && hiddenSet.has(current)){ sel.value = current; }
  }

  function setRowVisibility(id, visible){
    const row = rowById(id);
    if(!row) return;
    row.classList.toggle('noPrintRow', !visible);
    row.classList.toggle('rowHidden', !visible);
    if(visible){ hiddenSet.delete(id); }
    else{ hiddenSet.add(id); }
    const cb = row.querySelector('.rowToggle');
    if(cb){ cb.checked = visible; }
    refreshHiddenSelect();
    syncSessionHeaders();
    syncZoneVisibility();
    if(window.repaginateReport){ window.repaginateReport(); }
  }

  document.addEventListener('change', (e) => {
    const cb = e.target.closest('.rowToggle');
    if(!cb) return;
    const target = cb.dataset.target || "";
    if(!target) return;
    setRowVisibility(target, !!cb.checked);
  });

  window.restoreSelectedRow = function(){
    const sel = document.getElementById('hiddenRowsSelect');
    if(!sel || !sel.value) return;
    setRowVisibility(sel.value, true);
  };

  window.restoreAllHiddenRows = function(){
    Array.from(hiddenSet).forEach(id => setRowVisibility(id, true));
    syncSessionHeaders();
    syncZoneVisibility();
    if(window.repaginateReport){ window.repaginateReport(); }
  };

  syncSessionHeaders();
  syncZoneVisibility();
})();
"""



# -------------------------
# HOME (selector)
# -------------------------
def render_home(project: Optional[str] = None, print_mode: bool = False) -> str:
    """
    Page d'accueil : choix projet + réunion + boutons Générer / Imprimable.
    (Important) Toute la JS doit rester dans la string HTML -> sinon SyntaxError Python.
    """
    m = get_meetings().copy()
    m[M_COL_PROJECT_TITLE] = m[M_COL_PROJECT_TITLE].fillna("").astype(str).str.strip()
    m = m.loc[m[M_COL_PROJECT_TITLE] != ""].copy()

    projects = sorted(m[M_COL_PROJECT_TITLE].unique().tolist(), key=lambda x: x.lower())
    if project:
        m = m.loc[m[M_COL_PROJECT_TITLE] == project].copy()

    m["__date__"] = m[M_COL_DATE].apply(_parse_date_any)
    m = m.sort_values("__date__", ascending=False)

    project_opts = "".join(
        f'<option value="{_escape(p)}" {"selected" if p==project else ""}>{_escape(p)}</option>'
        for p in projects
    )

    meeting_opts = ""
    for _, r in m.iterrows():
        mid = str(r.get(M_COL_ID, "")).strip()
        d = _parse_date_any(r.get(M_COL_DATE))
        d_txt = _fmt_date(d) or _escape(r.get(M_COL_DATE_DISPLAY, "")) or _escape(r.get(M_COL_DATE, ""))
        proj = project or str(r.get(M_COL_PROJECT_TITLE, "")).strip()
        meeting_opts += f'<option value="{_escape(mid)}">{_escape(d_txt)} — {_escape(proj)}</option>'

    tempo_logo = _logo_data_url(LOGO_TEMPO_PATH)
    logo_html = f"<img src='{tempo_logo}' alt='TEMPO' class='homeLogo' />" if tempo_logo else "<div class='homeLogoText'>TEMPO</div>"
    return f"""
<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>TEMPO • CR Synthèse</title>
<style>
:root{{--text:#0b1220;--muted:#475569;--border:#e2e8f0;--soft:#f8fafc;--shadow:0 10px 30px rgba(2,6,23,.06);--accent:#0f172a;}}
*{{box-sizing:border-box}}
body{{margin:0;background:#fff;color:var(--text);font:14px/1.45 system-ui,-apple-system,Segoe UI,Roboto,Arial;}}
.wrap{{max-width:1100px;margin:0 auto;padding:26px;}}
.card{{background:#fff;border:1px solid var(--border);border-radius:16px;box-shadow:var(--shadow);padding:16px;}}
.brandline{{display:flex;gap:16px;align-items:center;margin-bottom:12px}}
.homeLogo{{height:44px;width:auto;display:block}}
.homeLogoText{{font-weight:1000;letter-spacing:.18em;font-size:20px}}
.tag{{color:var(--muted);font-weight:800}}
.grid{{display:grid;grid-template-columns:1fr 1fr;gap:14px}}
@media(max-width:780px){{.grid{{grid-template-columns:1fr}}}}
label{{display:block;font-weight:900;margin:0 0 6px}}
select{{width:100%;padding:12px 12px;border-radius:12px;border:1px solid var(--border);background:#fff;font-weight:700}}
.btn{{display:inline-flex;align-items:center;justify-content:center;gap:10px;padding:11px 14px;border-radius:12px;border:1px solid var(--border);background:var(--accent);color:#fff;font-weight:950;cursor:pointer;text-decoration:none}}
.btn.secondary{{background:#fff;color:var(--text);font-weight:900}}
.hint{{color:var(--muted);margin-top:10px;font-weight:700}}
</style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <div class="brandline">
        {logo_html}
        <div>
          <div style="font-weight:1000">Compte-rendu • Réunion de synthèse</div>
          <div class="tag">Application TEMPO</div>
        </div>
      </div>

      <div class="grid">
        <div>
          <label>Projet</label>
          <select id="project" onchange="onProjectChange()">
            <option value="">— Choisir —</option>
            {project_opts}
          </select>
        </div>
        <div>
          <label>Réunion</label>
          <select id="meeting">
            {meeting_opts if meeting_opts else '<option value="">— Sélectionne un projet —</option>'}
          </select>
        </div>
      </div>

      <div style="display:flex;gap:10px;margin-top:14px;flex-wrap:wrap">
        <button class="btn" type="button" onclick="openCR()">Ouvrir le compte-rendu</button>
      </div>

    </div>
  </div>

<script>
function onProjectChange(){{
  const p = document.getElementById('project').value || "";
  const url = p ? `/?project=${{encodeURIComponent(p)}}` : "/";
  window.location.href = url;
}}

function openCR(){{
  const meetingEl = document.getElementById('meeting');
  const projectEl = document.getElementById('project');
  if(!meetingEl){{ alert("Champ réunion introuvable"); return; }}
  const mid = meetingEl.value || "";
  if(!mid){{ alert("Choisis une réunion."); return; }}
  const p = projectEl ? (projectEl.value || "") : "";
  const url = `/cr?meeting_id=${{encodeURIComponent(mid)}}&project=${{encodeURIComponent(p)}}&print=1`;
  window.location.href = url;
}}
</script>

</body>
</html>
"""


def render_missing_data_page(err: MissingDataError) -> str:
    hint = f"Définis la variable d'environnement {err.env_var} pour pointer vers le fichier CSV."
    return f"""
<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Erreur de données — TEMPO</title>
  <style>
    :root{{--text:#0b1220;--muted:#475569;--border:#e2e8f0;--soft:#f8fafc;--shadow:0 10px 30px rgba(2,6,23,.06);--accent:#0f172a;}}
    *{{box-sizing:border-box}}
    body{{margin:0;background:#fff;color:var(--text);font:14px/1.45 system-ui,-apple-system,Segoe UI,Roboto,Arial;}}
    .wrap{{max-width:900px;margin:0 auto;padding:26px;}}
    .card{{background:#fff;border:1px solid var(--border);border-radius:16px;box-shadow:var(--shadow);padding:16px;}}
    .title{{font-weight:1000;font-size:20px;margin:0 0 10px 0;}}
    .muted{{color:var(--muted);font-weight:700;}}
    .mono{{font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,"Liberation Mono","Courier New",monospace;}}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <div class="title">Fichier CSV introuvable</div>
      <div class="muted">Impossible de charger la source de données requise.</div>
      <div style="margin-top:12px">
        <div><strong>Source :</strong> {_escape(err.label)}</div>
        <div><strong>Chemin :</strong> <span class="mono">{_escape(err.path)}</span></div>
      </div>
      <div style="margin-top:12px" class="muted">{_escape(hint)}</div>
    </div>
  </div>
</body>
</html>
"""


# -------------------------
# CR RENDER
# -------------------------
def render_cr(
    meeting_id: str,
    project: str = "",
    print_mode: bool = False,
    pinned_memos: str = "",
    range_start: str = "",
    range_end: str = "",
) -> str:
    mrow = meeting_row(meeting_id)
    meeting_entries = entries_for_meeting(meeting_id)

    project = (project or str(mrow.get(M_COL_PROJECT_TITLE, ""))).strip()
    meet_date = _parse_date_any(mrow.get(M_COL_DATE))
    ref_date = meet_date or date.today()
    date_txt = _fmt_date(meet_date) or _escape(mrow.get(M_COL_DATE_DISPLAY, "")) or _escape(mrow.get(M_COL_DATE, ""))
    range_start_date = _parse_date_any(range_start) if range_start else None
    range_end_date = _parse_date_any(range_end) if range_end else None
    if range_start_date is not None and range_end_date is None:
        range_end_date = ref_date
    range_active = range_start_date is not None or range_end_date is not None
    range_start_value = range_start_date.isoformat() if range_start_date else ""
    range_end_value = range_end_date.isoformat() if range_end_date else ""
    range_ref_date = range_end_date or ref_date if range_active else ref_date
    if range_active:
        project_entries = get_entries().copy()
        project_entries = project_entries.loc[
            project_entries[E_COL_PROJECT_TITLE].fillna("").astype(str).str.strip() == project
        ].copy()
        project_entries = _filter_entries_by_created_range(project_entries, range_start_date, range_end_date)
        edf = pd.concat([project_entries, meeting_entries], ignore_index=False)
        if E_COL_ID in edf.columns:
            edf["__id__"] = _series(edf, E_COL_ID, "").fillna("").astype(str).str.strip()
            edf = edf.loc[~edf["__id__"].duplicated(keep="first")].copy()
        else:
            edf = edf.drop_duplicates()
    else:
        edf = meeting_entries

    pinned_set = {p.strip() for p in str(pinned_memos or "").split(",") if p.strip()}

    # Project header info (Projects.csv)
    pinfo = project_info_by_title(project)
    proj_img = pinfo.get("image", "")
    proj_desc = pinfo.get("desc", "")
    proj_tl = " ".join([x for x in [pinfo.get("start", ""), pinfo.get("end", "")] if x]).strip()
    proj_status = pinfo.get("status", "")

    # Exclude duplicates in "À suivre": tasks already listed in CURRENT meeting only
    current_meeting_entry_ids = set(
        _series(meeting_entries, E_COL_ID, "").fillna("").astype(str).str.strip().tolist()
    )

    att, miss = compute_presence_lists(mrow)
    stats = kpis(mrow, edf, ref_date=range_ref_date)

    # Project-wide reminders / follow-ups
    rem_df = reminders_for_project(
        project_title=project,
        ref_date=range_ref_date,
        max_level=8,
        start_date=range_start_date,
        end_date=range_end_date,
    )
    fol_df = followups_for_project(
        project_title=project,
        ref_date=range_ref_date,
        exclude_entry_ids=current_meeting_entry_ids,
        start_date=range_start_date,
        end_date=range_end_date,
    )

    closed_recent_df = pd.DataFrame()
    project_history = get_entries().copy()
    project_history = project_history.loc[
        project_history[E_COL_PROJECT_TITLE].fillna("").astype(str).str.strip() == project
    ].copy()
    if not project_history.empty:
        edf2 = project_history.copy()
        edf2["__is_task__"] = _series(edf2, E_COL_IS_TASK, False).apply(_bool_true)
        edf2["__completed__"] = _series(edf2, E_COL_COMPLETED, False).apply(_bool_true)
        edf2["__deadline__"] = _series(edf2, E_COL_DEADLINE, None).apply(_parse_date_any)
        edf2["__done__"] = _series(edf2, E_COL_COMPLETED_END, None).apply(_parse_date_any)
        edf2.loc[edf2["__done__"].notna(), "__completed__"] = True
        edf2 = edf2.loc[(edf2["__is_task__"] == True) & (edf2["__completed__"] == True)].copy()
        edf2 = edf2.loc[edf2["__done__"].notna()].copy()
        days_since_done = pd.to_datetime(ref_date) - pd.to_datetime(edf2["__done__"])
        edf2 = edf2.loc[(days_since_done.dt.days >= 0) & (days_since_done.dt.days <= 14)].copy()
        edf2["__reminder__"] = edf2.apply(lambda r: reminder_level_at_done(r.get("__deadline__"), r.get("__done__")), axis=1)
        edf2 = _explode_areas(edf2)
        closed_recent_df = edf2

    closed_recent_ids: set[str] = set()
    if not closed_recent_df.empty:
        closed_recent_ids = set(_series(closed_recent_df, E_COL_ID, "").fillna("").astype(str).str.strip())
        closed_recent_ids.discard("")

    rem_company = reminders_by_company(rem_df)[:12]
    areas = group_meeting_by_area(edf)

    # ensure zones that exist only in reminders/follow-ups are also shown
    extra_zones = (
        set(rem_df["__area_list__"].astype(str).tolist())
        | set(fol_df["__area_list__"].astype(str).tolist())
        | set(closed_recent_df["__area_list__"].astype(str).tolist())
    )
    zone_names = [a for a, _ in areas]
    for z in sorted(extra_zones):
        if z not in zone_names:
            areas.append((z, edf.iloc[0:0].copy()))
    areas.sort(key=lambda x: (0 if x[0].lower() == "général" else 1, x[0].lower()))

    # Meeting labels for grouping rows by séance (notes/mémos/tâches)
    meetings_df = get_meetings().copy()
    if not meetings_df.empty:
        meetings_df = meetings_df.loc[
            meetings_df[M_COL_PROJECT_TITLE].fillna("").astype(str).str.strip() == project
        ].copy()
        meetings_df["__mid__"] = _series(meetings_df, M_COL_ID, "").fillna("").astype(str).str.strip()
        meetings_df["__mdate__"] = _series(meetings_df, M_COL_DATE, None).apply(_parse_date_any)
    meeting_date_by_id: Dict[str, Optional[date]] = {}
    if not meetings_df.empty:
        for _, mr in meetings_df.iterrows():
            mid = str(mr.get("__mid__", "")).strip()
            if not mid:
                continue
            mdate = mr.get("__mdate__")
            meeting_date_by_id[mid] = mdate
    meeting_index, _meeting_total = _meeting_sequence_for_project(meetings_df, meeting_id)
    cr_number_default = f"{meeting_index:02d}"

    # Pinned memos across history (editor helper)
    pinned_df = pd.DataFrame()
    img_col_pinned = None
    if pinned_set:
        pe = get_entries().copy()
        pe = pe.loc[pe[E_COL_PROJECT_TITLE].fillna("").astype(str).str.strip() == project].copy()
        pe["__id__"] = _series(pe, E_COL_ID, "").fillna("").astype(str).str.strip()
        pe = pe.loc[pe["__id__"].isin(pinned_set)].copy()
        pe["__is_task__"] = _series(pe, E_COL_IS_TASK, False).apply(_bool_true)
        pe = pe.loc[pe["__is_task__"] == False].copy()
        pe = _explode_areas(pe)
        pinned_df = pe
        img_col_pinned = detect_images_column(pinned_df)

    # -------------------------
    # Presence table (no KPI block)
    # -------------------------
    kpi_table_html = ""
    reminders_kpi_html = ""

    def render_presence_rows(items: List[Dict], label: str) -> str:
        if not items:
            return f"<tr><td>{_escape(label)} (0)</td><td class='muted'>—</td></tr>"
        rows = []
        for it in items:
            name = _escape(it.get("name", ""))
            logo = (it.get("logo", "") or "").strip()
            logo_html = f"<img class='coLogo' src='{_escape(logo)}' alt='' loading='lazy' />" if logo.startswith("http") else ""
            rows.append(f"<li class='presenceLine'>{logo_html}<span>{name}</span></li>")
        return f"<tr><td>{_escape(label)} ({len(items)})</td><td><ul class='presenceList'>{''.join(rows)}</ul></td></tr>"

    presence_html = f"""
      <div class="presenceWrap">
        <table class="annexTable coverTable presenceTable">
          <thead>
            <tr><th>Type</th><th>Entreprises</th></tr>
          </thead>
          <tbody>
            {render_presence_rows(att, "Présentes")}
            {render_presence_rows(miss, "Absentes / Excusées")}
          </tbody>
        </table>
      </div>
    """

    actions_html = f"""
      <div class="actions noPrint">
        <button class="btn" type="button" onclick="window.print()">Imprimer / PDF</button>
        <button class="btn secondary editCompact" type="button" onclick="window.refreshPagination && window.refreshPagination()">Recalculer la mise en page</button>
        <button class="btn secondary editCompact" id="btnQualityCheck" type="button">Qualité du texte</button>
        <button class="btn secondary editCompact" id="btnAnalysis" type="button">Analyse</button>
        <button class="btn secondary editCompact" id="btnRange" type="button" onclick="toggleRangePanel()">Choisir une période</button>
        <button class="btn secondary editCompact" id="btnConstraints" type="button">Contraintes HTML / impression</button>
        <button class="btn secondary editCompact" id="btnPrintPreview" type="button">Aperçu impression : OFF</button>
        <select id="hiddenRowsSelect" class="hiddenRowsSelect" title="Lignes masquées">
          <option value="">Lignes masquées…</option>
        </select>
        <button class="btn secondary editCompact" type="button" onclick="restoreSelectedRow()">Réafficher la ligne</button>
        <button class="btn secondary editCompact" type="button" onclick="restoreAllHiddenRows()">Réafficher tout</button>
        <a class="btn secondary" href="/">Changer de réunion</a>
      </div>
      <div class="rangePanel noPrint" id="rangePanel" style="display:{'flex' if range_active else 'none'}">
        <div class="rangeFields">
          <div class="rangeField">
            <label for="rangeStart">Du</label>
            <input type="date" id="rangeStart" value="{_escape(range_start_value)}" />
          </div>
          <div class="rangeField">
            <label for="rangeEnd">Au</label>
            <input type="date" id="rangeEnd" value="{_escape(range_end_value)}" />
          </div>
        </div>
        <div class="rangeActions">
          <button class="btn secondary" type="button" onclick="toggleRangePanel()">Fermer</button>
          <button class="btn secondary" type="button" onclick="clearRange()">Réinitialiser</button>
          <button class="btn" type="button" onclick="applyRange()">Appliquer</button>
        </div>
      </div>
      <div class="constraintsPanel noPrint" id="constraintsPanel" style="display:none">
        <div class="panelTitle">Détection des contraintes de mise en page</div>
        <div class="muted small">Désactive une contrainte pour voir immédiatement son effet sur l'affichage HTML et/ou l'impression.</div>
        <div class="constraintList">
          <label><input type="checkbox" data-constraint="fixedA4" checked /> Gabarit A4 fixe (largeur 210mm)</label>
          <label><input type="checkbox" data-constraint="fixedPageHeight" checked /> Hauteur de page forcée (297mm)</label>
          <label><input type="checkbox" data-constraint="pageBreaks" checked /> Sauts de page forcés entre sections</label>
          <label><input type="checkbox" data-constraint="bodyOffset" checked /> Décalage du body (panneau d'actions à gauche)</label>
          <label><input type="checkbox" data-constraint="pagePadding" checked /> Padding interne de la page</label>
          <label><input type="checkbox" data-constraint="footerReserve" checked /> Réserver l'espace avant footer (anti-chevauchement)</label>
          <label class="constraintSubControl">Niveau de réserve footer
            <input type="range" min="-100" max="150" step="5" value="100" id="footerReserveFactor" />
            <span id="footerReserveFactorValue">100 %</span>
          </label>
          <label><input type="checkbox" data-constraint="tableFixed" checked /> Colonnes de tableau en layout fixe</label>
          <label><input type="checkbox" data-constraint="printHideUi" checked /> Masquer les outils UI à l'impression</label>
          <label><input type="checkbox" data-constraint="printStickyHeader" checked /> Header sticky en impression</label>
          <label><input type="checkbox" data-constraint="printCompactRows" checked /> Compactage des lignes pour imprimer</label>
          <label><input type="checkbox" data-constraint="printAvoidSplitRows" checked /> Empêcher la coupure de lignes/blocs</label>
          <label><input type="checkbox" data-constraint="keepSessionHeaderWithNext" checked /> Ne pas laisser « En séance du » seul en bas de page</label>
          <label><input type="checkbox" data-constraint="printAutoOptimize" checked /> Optimisation auto avant impression</label>
          <label><input type="checkbox" data-constraint="topScale" checked /> Mise à l'échelle du bandeau haut</label>
        </div>
      </div>
    """

    # Card renderer for tasks outside the meeting (rappels / à-suivre) — NO BADGES
    def render_task_card_from_row(r, tag: str, extra_class: str, img_col: Optional[str]) -> str:
        title = _format_entry_text_html(r.get(E_COL_TITLE, ""))
        company = _escape(r.get(E_COL_COMPANY_TASK, ""))
        owner = _escape(r.get(E_COL_OWNER, ""))
        deadline = _fmt_date(_parse_date_any(r.get(E_COL_DEADLINE)))
        done = ""
        if _bool_true(r.get(E_COL_COMPLETED)):
            done = _fmt_date(_parse_date_any(r.get(E_COL_COMPLETED_END)))

        concerne = " • ".join([x for x in [company, owner] if x])
        status_txt = _escape(r.get(E_COL_STATUS, "")) or ("Terminé" if _bool_true(r.get(E_COL_COMPLETED)) else "Non terminé")

        img_urls = parse_image_urls_any(r.get(img_col)) if img_col else []
        images_html = render_images_gallery(img_urls, print_mode=print_mode)
        comment_html = render_task_comment(r)

        return f"""
          <div class="topic {extra_class}">
            <div class="topicTop">
              <div class="topicTitle">{title}</div>
              <div class="topicRight">
                <div class="rRow"><div class="rLab">Type</div><div class="rVal">Tâche</div></div>
                <div class="rRow"><div class="rLab">Tag</div><div class="rVal">{_escape(tag) or "—"}</div></div>
                <div class="rRow"><div class="rLab">Statut</div><div class="rVal">{status_txt}</div></div>
              </div>
            </div>

            <div class="meta4">
              <div><div class="metaLabel">Pour le</div><div class="metaVal">{deadline or "—"}</div></div>
              <div><div class="metaLabel">Fait le</div><div class="metaVal">{done or "—"}</div></div>
              <div><div class="metaLabel">Concerne</div><div class="metaVal">{concerne or "—"}</div></div>
              <div><div class="metaLabel">Lot</div><div class="metaVal">{_lot_abbrev_list(r.get(E_COL_PACKAGES, "")) or "—"}</div></div>
            </div>

            {images_html}
            {comment_html}
          </div>
        """

    # Pre-detect image column for each dataset
    img_col_meeting = detect_images_column(edf)
    img_col_memo = detect_memo_images_column(edf)
    img_col_rem = detect_images_column(rem_df)
    img_col_fol = detect_images_column(fol_df)

    # -------------------------
    # PDF TABLE RENDER (NO CARDS)
    # -------------------------
    def render_task_row_tr(
        r,
        tag_text: str,
        img_col: Optional[str] = None,
        is_meeting: bool = False,
        reminder_closed: bool = False,
        completed_recent: bool = False,
        row_id: str = "",
    ) -> str:
        title = _format_entry_text_html(r.get(E_COL_TITLE, ""))
        company = _escape(r.get(E_COL_COMPANY_TASK, ""))
        packages = _escape(r.get(E_COL_PACKAGES, ""))
        concerne_display = _concerne_trigram(company)

        created = _fmt_date(_parse_date_any(r.get(E_COL_CREATED)))
        deadline = _fmt_date(_parse_date_any(r.get(E_COL_DEADLINE)))

        done = ""
        if _bool_true(r.get(E_COL_COMPLETED)):
            done = _fmt_date(_parse_date_any(r.get(E_COL_COMPLETED_END)))

        is_task = _bool_true(r.get(E_COL_IS_TASK))
        deadline_display = deadline or "—" if is_task else "/"
        done_display = done or "—" if is_task else "/"
        lot_display = _lot_abbrev_list(packages) or "—"
        if not is_task and _has_multiple_companies(company):
            concerne_display = "PE"
        else:
            concerne_display = concerne_display or "PE"

        memo_img_col = img_col_memo if (not is_task and img_col_memo) else img_col
        img_urls = parse_image_urls_any(r.get(memo_img_col)) if memo_img_col else []
        thumbs = ""
        if img_urls:
            thumbs_imgs = "".join(
                f"<span class='thumbAWrap' data-thumb><a class='thumbA' href='{_escape(u)}' target='_blank' rel='noopener'><img class='thumb' src='{_escape(u)}' alt='' /></a><button type='button' class='thumbRemove noPrint' title='Supprimer'>×</button><span class='thumbHandle' title='Déplacer / redimensionner'></span></span>"
                for u in img_urls[:6]
            )
            thumbs = f"<div class='thumbs' data-gallery>{thumbs_imgs}</div>"

        row_cls = "rowItem rowMeeting" if is_meeting else "rowItem"
        if completed_recent:
            row_cls += " rowDoneRecent"

        tag_display = _escape(tag_text).replace(" ", "&nbsp;")
        tag_class = "tagReminderGreen" if tag_text.lower().startswith("rappel") and reminder_closed else "tagReminder"
        tag_html = (
            f"<span class='{tag_class}'>{tag_display}</span>"
            if tag_text.lower().startswith("rappel")
            else tag_display
        )

        safe_row_id = _escape(row_id) or _escape(str(r.get(E_COL_ID, "")))
        toggle_html = f"<input type='checkbox' class='rowToggle noPrint' data-target='{safe_row_id}' checked />"
        return f"""
          <tr class="{row_cls} compactRow" data-row-id="{safe_row_id}" data-entry-type="{"task" if is_task else "memo"}">
            <td class="colType">{toggle_html}<div>{tag_html or "—"}</div></td>
            <td class="colComment">
              <div class="rowImageTools noPrint"><button type="button" class="btnAddImage">+ Image</button><input type="file" class="imageInput" accept="image/*" multiple hidden /></div>
              <div class="commentText">{title}</div>
              {thumbs}
              {render_entry_comment(r)}
            </td>
            <td class="colDate">{created or "—"}</td>
            <td class="colDate">{deadline_display}</td>
            <td class="colDate">{done_display}</td>
            <td class="colLot editableCell" contenteditable="true">{lot_display}</td>
            <td class="colWho editableCell" contenteditable="true">{concerne_display}</td>
          </tr>
        """

    def render_session_subheader_tr(session_label: str, is_current_session: bool = False) -> str:
        return f"""
          <tr class="sessionSubRow{' sessionSubRowCurrent' if is_current_session else ''}">
            <td class="colType">—</td>
            <td class="colComment" colspan="6"><strong>{_escape(session_label)}</strong></td>
          </tr>
        """

    def _meeting_sort_and_label(r) -> Tuple[Optional[date], str]:
        mid = str(r.get(E_COL_MEETING_ID, "")).strip()
        created_d = _parse_date_any(r.get(E_COL_CREATED))
        if mid and mid in meeting_date_by_id and meeting_date_by_id[mid] is not None:
            d = meeting_date_by_id[mid]
        else:
            d = created_d
        if d:
            return d, f"En séance du {d.strftime('%d/%m/%Y')} :"
        return None, "Hors séance :"

    def render_zone_table(area_name: str, rows_html: str) -> str:
        if not rows_html.strip():
            return ""
        zt = _escape(area_name)
        return f"""
        <div class="zoneBlock reportBlock" data-zone-id="{zt}">
          <div class="zoneTitle">
            <span>{zt}</span>
            <div class="zoneTools noPrint">
              <button class="zoneBtn" type="button" data-action="move-up">↑</button>
              <button class="zoneBtn" type="button" data-action="move-down">↓</button>
              <button class="zoneBtn" type="button" data-action="highlight">Surligner</button>
                                                        <button class="btnAddMemo" type="button" data-area="{zt}">+ Ajouter mémo</button>
            </div>
          </div>
          <table class="crTable">
            <colgroup>
              <col style="width:var(--col-type)" />
              <col style="width:var(--col-comment)" />
              <col style="width:var(--col-date)" />
              <col style="width:var(--col-date)" />
              <col style="width:var(--col-date)" />
              <col style="width:var(--col-lot)" />
              <col style="width:var(--col-who)" />
            </colgroup>
            <thead>
              <tr>
                <th class="colType">Type <span class="colGrip" data-col="type"></span></th>
                <th class="colComment">Commentaires et observations <span class="colGrip" data-col="comment"></span></th>
                <th class="colDate">Écrit le <span class="colGrip" data-col="date"></span></th>
                <th class="colDate">Pour le <span class="colGrip" data-col="date2"></span></th>
                <th class="colDate">Fait le <span class="colGrip" data-col="date3"></span></th>
                <th class="colLot">Lot <span class="colGrip" data-col="lot"></span></th>
                <th class="colWho">Concerne <span class="colGrip" data-col="who"></span></th>
              </tr>
            </thead>
            <tbody>
              {rows_html}
            </tbody>
          </table>
        </div>
        """

    # Build per-zone blocks
    zones_html_parts: List[str] = []

    current_session_label = (
        f"En séance du {(meet_date or ref_date).strftime('%d/%m/%Y')} :" if (meet_date or ref_date) else ""
    )

    def _entry_id_value(r) -> str:
        return str(r.get(E_COL_ID, "")).strip()

    def _is_completed_recent_row(r) -> bool:
        rid = _entry_id_value(r)
        return bool(rid and rid in closed_recent_ids)

    for area_name, g in areas:
        grouped_rows: List[Tuple[Optional[date], str, str]] = []
        seen_entry_ids: set[str] = set()

        rem_zone = rem_df.loc[rem_df["__area_list__"].astype(str) == str(area_name)].copy()
        if not rem_zone.empty:
            for idx, r in rem_zone.iterrows():
                rid = _entry_id_value(r)
                row_html = render_task_row_tr(
                    r,
                    f"Rappel {int(r.get('__reminder__') or 1)}",
                    img_col=img_col_rem,
                    is_meeting=False,
                    completed_recent=_is_completed_recent_row(r),
                    row_id=f"rem-{area_name}-{idx}",
                )
                sort_d, label = _meeting_sort_and_label(r)
                grouped_rows.append((sort_d, label, row_html))
                if rid:
                    seen_entry_ids.add(rid)

        fol_zone = fol_df.loc[fol_df["__area_list__"].astype(str) == str(area_name)].copy()
        if not fol_zone.empty:
            for idx, r in fol_zone.iterrows():
                rid = _entry_id_value(r)
                row_html = render_task_row_tr(
                    r,
                    "Tâche",
                    img_col=img_col_fol,
                    is_meeting=False,
                    completed_recent=_is_completed_recent_row(r),
                    row_id=f"fol-{area_name}-{idx}",
                )
                sort_d, label = _meeting_sort_and_label(r)
                grouped_rows.append((sort_d, label, row_html))
                if rid:
                    seen_entry_ids.add(rid)

        if pinned_set and (not pinned_df.empty):
            pin_zone = pinned_df.loc[pinned_df["__area_list__"].astype(str) == str(area_name)].copy()
            if not pin_zone.empty:
                for idx, r in pin_zone.iterrows():
                    rid = _entry_id_value(r)
                    row_html = render_task_row_tr(
                        r,
                        "Mémo",
                        img_col=img_col_pinned,
                        is_meeting=False,
                        row_id=f"pin-{area_name}-{idx}",
                    )
                    sort_d, label = _meeting_sort_and_label(r)
                    grouped_rows.append((sort_d, label, row_html))
                    if rid:
                        seen_entry_ids.add(rid)

        if not g.empty:
            g_view = g.copy().sort_values(by=E_COL_CREATED, na_position="last")
            for idx, r in g_view.iterrows():
                rid = _entry_id_value(r)
                tag = "Tâche" if _bool_true(r.get(E_COL_IS_TASK)) else "Mémo"
                is_meeting_entry = str(r.get(E_COL_MEETING_ID, "")).strip() == str(meeting_id)
                row_html = render_task_row_tr(
                    r,
                    tag,
                    img_col=img_col_meeting,
                    is_meeting=is_meeting_entry,
                    completed_recent=_is_completed_recent_row(r),
                    row_id=f"meet-{area_name}-{idx}",
                )
                sort_d, label = _meeting_sort_and_label(r)
                grouped_rows.append((sort_d, label, row_html))
                if rid:
                    seen_entry_ids.add(rid)

        closed_zone = (
            closed_recent_df.loc[closed_recent_df["__area_list__"].astype(str) == str(area_name)].copy()
            if not closed_recent_df.empty
            else pd.DataFrame()
        )
        if not closed_zone.empty:
            for idx, r in closed_zone.iterrows():
                rid = _entry_id_value(r)
                if rid and rid in seen_entry_ids:
                    continue
                lvl = r.get("__reminder__")
                tag = f"Rappel {int(lvl)}" if pd.notna(lvl) else "Tâche"
                row_html = render_task_row_tr(
                    r,
                    tag,
                    img_col=img_col_meeting,
                    is_meeting=False,
                    reminder_closed=True,
                    completed_recent=True,
                    row_id=f"closed-{area_name}-{idx}",
                )
                sort_d, label = _meeting_sort_and_label(r)
                grouped_rows.append((sort_d, label, row_html))
                if rid:
                    seen_entry_ids.add(rid)

        grouped_rows.sort(key=lambda item: (item[0] is None, item[0] or date.max, item[1]))
        rows_parts: List[str] = []
        current_label = None
        for _, label, row_html in grouped_rows:
            if label != current_label:
                rows_parts.append(render_session_subheader_tr(label, is_current_session=(label == current_session_label)))
                current_label = label
            rows_parts.append(row_html)

        zone_table_html = render_zone_table(area_name, "\n".join(rows_parts))
        if zone_table_html:
            zones_html_parts.append(zone_table_html)

    zones_html = "".join(zones_html_parts)
    report_note_html = ""

    # -------------------------
    # CSS
    # -------------------------
    css = f"""
:root{{
  --bg:#ffffff;
  --text:#0b1220;
  --muted:#475569;
  --border:#e2e8f0;
  --soft:#f8fafc;
  --shadow:0 10px 30px rgba(2,6,23,.06);
  --accent:#0f172a;
  --blueSoft:#eff6ff;
  --blueBorder:#bfdbfe;
  --col-type:7%;
  --col-comment:53%;
  --col-date:8%;
  --col-lot:8%;
  --col-who:8%;
  --a4-width:210mm;
  --a4-padding-x:6mm;
  --kpi-cols:4;
  --top-scale:1;
  --footer-reserve-factor:1;
}}
*{{box-sizing:border-box}}
html,body{{margin:0;padding:0;background:var(--bg);color:var(--text);font:14px/1.45 system-ui,-apple-system,Segoe UI,Roboto,Arial;-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
body{{padding:14px 14px 14px 280px;}}
.wrap{{display:flex;flex-direction:column;gap:12px;align-items:center;}}
.page{{width:210mm;height:297mm;min-height:297mm;position:relative;background:#fff;overflow:visible;break-after:page;page-break-after:always;}}
.page:last-child{{break-after:auto;page-break-after:auto;}}
.pageContent{{padding:10mm 8mm 34mm 8mm;}}
.page--cover .pageContent{{padding-top:0;}}
.muted{{color:var(--muted)}}
.small{{font-size:12px}}
.noPrint{{}}
@media print{{ .noPrint{{display:none!important}} }}
@media print{{body{{padding:0;background:#fff}} .page{{margin:0;box-shadow:none}}}}
body.printOptimized .reportBlocks{{gap:0!important}}
body.printOptimized .zoneBlock{{margin:0!important}}
body.printOptimized .crTable th, body.printOptimized .crTable td{{padding:4px 5px!important;line-height:1.16!important}}
body.printOptimized .reportHeader{{margin-bottom:4px!important}}
body.printOptimized .thumb{{height:64px!important;max-width:110px!important}}
body.printPreviewMode .rowToggle,
body.printPreviewMode .rowImageTools,
body.printPreviewMode .thumbRemove,
body.printPreviewMode .btnAddMemo,
body.printPreviewMode .colGrip{{display:none!important}}
body.printPreviewMode .editableCell{{background:transparent!important;box-shadow:none!important}}
body.printPreviewMode .editableCell:focus{{box-shadow:none!important}}
body.printPreviewMode .noPrintRow{{display:none!important}}
@media screen{{body{{background:#e5e7eb;}} .page{{box-shadow:0 14px 30px rgba(15,23,42,.16)}}}}
.topPage{{transform:scale(var(--top-scale));transform-origin:top left}}
@media print{{.topPage{{margin:0;}}}}
.reportTables{{margin-top:0}}
.coverHero{{position:relative;overflow:hidden;background:#fff;min-height:420px}}
.coverHeroImg{{position:relative;min-height:430px;background-size:cover;background-position:center}}
.coverHeroFade{{position:absolute;inset:0;background:linear-gradient(180deg,rgba(255,255,255,.08),rgba(255,255,255,0));}}
.coverHeroCurve{{position:absolute;left:50%;bottom:-95px;width:135%;height:190px;transform:translateX(-50%);background:#fff;border-radius:50% 50% 0 0 / 100% 100% 0 0;z-index:2}}
.coverHeroLogoWrap{{position:absolute;left:50%;bottom:18px;transform:translateX(-50%);z-index:4;background:#fff;padding:10px 18px;border-radius:8px;box-shadow:0 6px 18px rgba(2,6,23,.12)}}
.coverHeroLogo{{height:110px;width:auto;display:block}}
.coverNoteCenter{{text-align:center;padding:10px 16px 12px 16px;font-weight:900;display:flex;flex-direction:column;align-items:center;gap:10px}}
.coverAppNote{{margin-top:8px;font-family:"Arial Nova Cond Light","Arial Narrow",Arial,sans-serif;font-size:14px;line-height:1.45;color:#f97316;font-style:italic;font-weight:600;max-width:640px}}
.coverUrl{{margin-top:6px;font-weight:900;color:#f97316;text-decoration:underline;text-underline-offset:3px}}
.coverUrl::after{{content:" ↗";font-weight:900}}
.coverProjectTitle{{font-family:"Arial Nova Cond Light","Arial Narrow",Arial,sans-serif;font-size:22px;line-height:1.2;color:#f59e0b;font-weight:700;letter-spacing:.5px;text-transform:uppercase}}
.coverCrTitle{{margin-top:10px;font-family:"Arial Nova Cond Light","Arial Narrow",Arial,sans-serif;font-size:22px;line-height:1.2;color:#0f3a40;font-weight:700}}
.coverCrMeta{{margin-top:8px;font-family:"Arial Nova Cond Light","Arial Narrow",Arial,sans-serif;font-size:22px;line-height:1.2;color:#0f3a40;font-weight:700}}
.editInline{{display:inline-block;min-width:40px;padding:0 4px;border-bottom:2px dashed #cbd5e1;outline:none}}
@media print{{.editInline{{border-bottom:none}}}}
.nextMeetingBox{{margin:18px auto 0 auto;max-width:78%;border:2px solid #111;padding:12px 10px;font-weight:1000}}
.nextMeetingLine1{{font-family:"Arial Nova Cond Light","Arial Narrow",Arial,sans-serif;font-size:18px}}
.nextMeetingLine2{{font-family:"Arial Nova Cond Light","Arial Narrow",Arial,sans-serif;font-size:18px;color:#ef4444;margin-top:5px}}
.nextMeetingLine3{{font-family:"Arial Nova Cond Light","Arial Narrow",Arial,sans-serif;font-size:18px;color:#111;margin-top:4px;outline:none}}
@media print{{.coverHeroImg{{min-height:390px}} .coverProjectTitle{{font-size:44px}} .coverCrTitle{{font-size:33px}} .coverCrMeta{{font-size:36px}} .nextMeetingLine1{{font-size:18px}} .nextMeetingLine2{{font-size:32px}} .nextMeetingLine3{{font-size:27px}}}}

/* PROJECT BANNER */
.banner{{
  border:1px solid var(--border);
  border-radius:18px;
  overflow:hidden;
  background:linear-gradient(180deg,#fff, var(--soft));
}}
.bannerImg{{position:relative;min-height:260px;background-size:cover;background-position:center;}}
.bannerOverlay{{position:absolute;inset:0;background:linear-gradient(90deg, rgba(2,6,23,.78), rgba(2,6,23,.10));}}
.bannerContent{{position:relative;padding:18px;color:#fff;max-width:900px;}}
.bannerKicker{{font-weight:800;opacity:.9}}
.bannerTitle{{font-size:26px;font-weight:1000;letter-spacing:.2px;margin-top:6px}}
.bannerMeta{{margin-top:10px;display:flex;flex-wrap:wrap;gap:10px}}
.bannerChip{{background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.18);padding:7px 10px;border-radius:999px;font-weight:700;}}
.bannerDesc{{margin-top:10px;opacity:.95}}
@media print{{.bannerImg{{min-height:300px}} .bannerTitle{{font-size:22px}} .bannerContent{{padding:14px}}}}

/* BANNER LOGO */
.bannerLogoWrap{{display:flex;justify-content:flex-start;margin-bottom:8px}}
.bannerLogo{{height:72px;width:auto;display:block}}

/* KPI */
.card{{background:#fff;border:1px solid var(--border);border-radius:16px;box-shadow:var(--shadow);padding:16px;margin-top:14px;}}
.kpis{{display:grid;grid-template-columns:repeat(var(--kpi-cols),1fr);gap:10px;margin-top:12px}}
.kpi{{border:1px solid var(--border);border-radius:14px;background:#fff;padding:10px}}
.kpi_t{{color:var(--muted);font-weight:700;font-size:11px}}
.kpi_v{{font-weight:1000;font-size:20px;margin-top:6px}}
.topGrip{{height:8px;width:120px;background:#e2e8f0;border-radius:999px;margin:8px auto 0;cursor:ns-resize}}
@media (max-width: 980px){{.kpis{{grid-template-columns:repeat(3,1fr)}}}}
@media print{{.kpis{{grid-template-columns:repeat(4,1fr);gap:6px}} .kpi{{padding:6px}} .kpi_v{{font-size:16px}}}}

/* Sections */
.section{{margin-top:18px}}
.sectionTitle{{
  display:flex;align-items:center;gap:10px;
  padding:14px 14px;border:1px solid var(--border);border-radius:16px;
  background:linear-gradient(180deg,#fff, var(--soft));
  font-weight:1000;font-size:16px;letter-spacing:.2px;
  border-left:6px solid #0f172a;
}}
.zoneTitle{{
  display:flex;align-items:center;gap:10px;
  padding:6px 10px;border:1px solid var(--border);border-bottom:none;
  background:#f59e0b;color:#ffffff;font-weight:900;font-size:11px;text-transform:uppercase;
}}
.zoneTitle button{{margin-left:auto}}
.zoneTools{{display:flex;align-items:center;gap:6px;margin-left:auto}}
.zoneBtn{{border:1px solid #ffffff;background:#fff;border-radius:8px;padding:4px 8px;font-weight:800;cursor:pointer}}
.zoneBlock.highlight{{box-shadow:0 0 0 2px #f59e0b inset; background:linear-gradient(180deg,#fff7ed,#fff)}}
.zoneBlock.pageBreakBefore{{page-break-before:always}}
.u-page-break{{break-before:page;page-break-before:always;}}
.u-avoid-break{{break-inside:avoid;page-break-inside:avoid;}}

.presGrid{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:10px}}
@media (max-width: 780px){{.presGrid{{grid-template-columns:1fr}}}}
.subcard{{border:1px solid var(--border);border-radius:14px;background:#fff;padding:12px}}
.subhead{{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:10px;font-weight:900}}
.chips{{display:flex;flex-wrap:wrap;gap:8px}}
.chip{{padding:7px 10px;border-radius:999px;border:1px solid var(--border);background:#fff;font-weight:700;display:inline-flex;align-items:center;gap:8px;}}
.coLogo{{width:18px;height:18px;border-radius:6px;object-fit:cover;display:block}}

/* Topics */
.topics{{display:flex;flex-direction:column;gap:12px;margin-top:10px}}
.topic{{border:1px solid var(--border);border-radius:14px;background:#fff;padding:12px}}
.topicTop{{display:grid;grid-template-columns:1fr 210px;gap:12px;align-items:start;}}
.topicTitle{{font-weight:600;font-size:15px;line-height:1.25}}
.topicRight{{display:flex;flex-direction:column;gap:8px;align-items:stretch;}}
.rRow{{display:flex;justify-content:space-between;gap:10px;border:1px solid var(--border);border-radius:12px;padding:6px 8px;background:#fff}}
.rLab{{color:var(--muted);font-weight:800;font-size:11px}}
.rVal{{font-weight:900;font-size:12px;text-align:right;max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}

.meta4{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-top:10px}}
@media (max-width: 900px){{.meta4{{grid-template-columns:repeat(2,1fr)}}}}
.metaLabel{{color:var(--muted);font-weight:700;font-size:11px}}
.metaVal{{font-weight:700}}
.topicComment{{margin-top:10px;border-top:1px dashed var(--border);padding-top:10px}}

.imgRow{{display:flex;gap:10px;flex-wrap:wrap;margin-top:10px}}
.imgThumb{{display:block;width:320px;height:200px;border-radius:12px;overflow:hidden;border:1px solid var(--border);background:#fff}}
.imgThumb img{{width:100%;height:100%;object-fit:cover;display:block}}
.imgThumb{{position:relative;resize:both;overflow:auto}}
.imgGrip{{position:absolute;right:6px;bottom:6px;width:14px;height:14px;border:2px solid rgba(15,23,42,.45);border-top:none;border-left:none;pointer-events:none}}

.actions{{position:fixed;top:14px;left:14px;z-index:9999;display:flex;flex-direction:column;gap:8px;width:248px;padding:10px;border:1px solid var(--border);border-radius:12px;background:rgba(255,255,255,.97);box-shadow:0 8px 24px rgba(2,6,23,.12)}}
.actions .btn,.actions .hiddenRowsSelect{{width:100%}}
.btn{{display:inline-flex;align-items:center;justify-content:center;gap:10px;padding:11px 14px;border-radius:12px;border:1px solid var(--border);background:var(--accent);color:#fff;font-weight:950;cursor:pointer;text-decoration:none}}
.btn.secondary{{background:#fff;color:var(--text);font-weight:900}}
#btnPrintPreview.active{{background:#0f172a;color:#fff;border-color:#0f172a}}
.rangePanel{{position:fixed;top:14px;left:14px;z-index:10001;width:248px;border:1px solid var(--border);border-radius:14px;padding:12px;background:#fff;display:flex;flex-direction:column;gap:10px;box-shadow:0 8px 24px rgba(2,6,23,.12);max-height:calc(100vh - 32px);overflow:auto}}
.constraintsPanel{{position:fixed;top:14px;left:276px;z-index:10001;width:420px;border:1px solid var(--border);border-radius:14px;padding:12px;background:#fff;display:flex;flex-direction:column;gap:10px;box-shadow:0 8px 24px rgba(2,6,23,.12);max-height:calc(100vh - 32px);overflow:auto}}
.panelTitle{{font-weight:900;font-size:13px}}
.constraintList{{display:grid;grid-template-columns:1fr;gap:6px}}
.constraintList label{{display:flex;align-items:flex-start;gap:8px;font-size:12px;line-height:1.25}}
.constraintList label.constraintSubControl{{display:grid;grid-template-columns:130px 1fr auto;align-items:center;gap:8px;margin-left:22px}}
.constraintList label.constraintSubControl input[type="range"]{{width:100%}}
.rangeFields{{display:flex;gap:12px;flex-wrap:wrap}}
.rangeField{{display:flex;flex-direction:column;gap:6px;min-width:180px}}
.rangeField label{{font-weight:900;font-size:12px}}
.rangeField input{{padding:8px 10px;border-radius:10px;border:1px solid var(--border);font-weight:700}}
.rangeActions{{display:flex;gap:10px;flex-wrap:wrap}}
.hiddenRowsSelect{{padding:9px 10px;border:1px solid var(--border);border-radius:10px;font-weight:700;background:#fff;min-width:220px}}
@media print{{.actions{{margin:8px 0}} .btn{{padding:8px 10px;font-size:12px}}}}

/* Bleu = sujets réunion */
.newItem{{border-color: var(--blueBorder);background: linear-gradient(180deg, #ffffff, var(--blueSoft));box-shadow: 0 0 0 2px rgba(59,130,246,.05);}}
.reminderItem{{border-left:4px solid #ef4444;}}
.followItem{{border-left:4px solid #f59e0b;}}

/* KPI list */
.kpiList{{display:flex;flex-direction:column;gap:8px}}
.kpiRow{{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:8px 10px;border:1px solid var(--border);border-radius:12px;background:#fff}}
.kpiCo{{display:inline-flex;align-items:center;gap:10px;font-weight:900}}
.kpiCount{{font-weight:1000}}

/* PRINT TABLE */
@page {{ size: A4 portrait; margin: 0; }}
body.constraint-off-fixedA4 .page{{width:auto!important}}
body.constraint-off-fixedPageHeight .page{{height:auto!important;min-height:auto!important}}
body.constraint-off-pageBreaks .page,body.constraint-off-pageBreaks .page:last-child{{break-after:auto!important;page-break-after:auto!important}}
body.constraint-off-bodyOffset{{padding:14px!important}}
body.constraint-off-pagePadding .pageContent{{padding:0!important}}
body.constraint-off-footerReserve .pageContent{{padding-bottom:8mm!important}}
body.constraint-off-tableFixed .crTable{{table-layout:auto!important}}
body.constraint-off-printStickyHeader .printHeaderFixed{{position:static!important;top:auto!important}}
body.constraint-off-printCompactRows.printOptimized .crTable th,body.constraint-off-printCompactRows.printOptimized .crTable td{{padding:7px 8px!important;line-height:1.3!important}}
body.constraint-off-printCompactRows.printOptimized .thumb{{height:80px!important;max-width:140px!important}}
body.constraint-off-topScale .topPage{{transform:none!important}}
@media print{{body.constraint-off-printHideUi .actions,body.constraint-off-printHideUi .rangePanel,body.constraint-off-printHideUi .constraintsPanel{{display:flex!important}}}}
@media print{{body.constraint-off-printAvoidSplitRows .sessionSubRow,body.constraint-off-printAvoidSplitRows .zoneTitle{{break-inside:auto!important;page-break-inside:auto!important;break-after:auto!important;page-break-after:auto!important}}}}

.zoneBlock{{margin:0}}
.zoneBlock + .zoneBlock{{margin-top:0}}
.reportBlocks{{display:flex;flex-direction:column;gap:0}}
.reportBlock{{break-inside:auto;page-break-inside:auto}}
.reportNote{{margin-top:12px}}
.crTable{{width:100%;border-collapse:collapse;table-layout:fixed;border:1px solid var(--border);margin-top:-1px;}}
.crTable thead{{display:table-header-group}}
.crTable tfoot{{display:table-footer-group}}
.crTable th, .crTable td{{border:1px solid var(--border);padding:6px 7px;vertical-align:top;page-break-inside:auto;break-inside:auto;}}
.crTable tr{{page-break-inside:auto;break-inside:auto;}}
.annexTable tr{{page-break-inside:auto;break-inside:auto;}}
.crTable th{{background:#1f4e4f;color:#fff;text-align:center;font-weight:900;font-size:11px;line-height:1.2;white-space:nowrap}}
.crTable td{{font-size:11px;line-height:1.24;word-break:normal;overflow-wrap:break-word;hyphens:none}}
.crTable td.colDate, .crTable th.colDate{{padding:6px 4px}}

.sessionSubRow td{{background:#ffffff;}}
.sessionSubRow td.colType{{color:#94a3b8;font-weight:700;}}
.sessionSubRow td.colComment{{font-size:12px;color:#111827;font-weight:900;text-decoration:none;}}
.sessionSubRowCurrent td.colComment{{color:#1d4ed8;text-decoration:underline;text-underline-offset:2px;}}
.colType{{text-align:center;font-weight:1000;white-space:nowrap;position:relative}}
.colComment{{white-space:normal;position:relative}}
.rowImageTools{{display:flex;justify-content:flex-end;margin-bottom:4px}}
.btnAddImage{{border:1px solid #d1d5db;background:#fff;border-radius:8px;padding:2px 8px;font-size:11px;font-weight:800;cursor:pointer}}
.btnAddImage:hover{{background:#f8fafc}}
.colDate{{text-align:center;font-variant-numeric: tabular-nums;white-space:nowrap;position:relative}}
.colLot{{text-align:center;white-space:nowrap;position:relative}}
.colWho{{text-align:center;white-space:nowrap;position:relative}}
.rowToggle{{width:14px;height:14px;accent-color:#0f172a;cursor:pointer}}
.editableCell{{background:#fff7ed;outline:none}}
.editableCell:focus{{box-shadow:inset 0 0 0 2px #fb923c}}
.noPrintRow{{opacity:.4}}
.rowDoneRecent td{{background:none!important}}
.crTable tr.rowDoneRecent td.colType{{box-shadow:inset 4px 0 0 #16a34a;}}
.crTable tr.rowDoneRecent td.colType div{{color:#15803d;font-weight:900;}}
.rowHidden{{display:none!important}}
.colGrip{{position:absolute;top:0;right:-6px;width:12px;height:100%;cursor:col-resize}}
.colGrip::after{{content:"";position:absolute;top:3px;bottom:3px;left:5px;width:2px;background:#cbd5f5;border-radius:2px;opacity:.7}}

@media print{{ .rowToggle{{display:none}} .noPrintRow{{display:none}} .editableCell{{background:transparent}} .rowImageTools{{display:none!important}} .thumbRemove{{display:none!important}} }}
@media print{{ .sessionSubRow{{break-inside:avoid;page-break-inside:avoid}} .zoneTitle{{break-after:avoid-page;page-break-after:avoid}} }}


.crTable tr.rowMeeting td{{background:#eef8ff;}}
.crTable tr.rowMeeting td.colType{{box-shadow:inset 4px 0 0 #2563eb;}}

.thumbs{{margin-top:6px;display:flex;flex-wrap:wrap;gap:8px;align-items:flex-start}}
.thumb{{width:160px;height:auto;max-width:100%;border:1px solid var(--border);border-radius:8px;display:block;object-fit:cover;background:#fff}}
.entryComment{{margin-top:8px;padding-left:12px;border-left:3px solid #e2e8f0}}
.tagReminderGreen{{color:#16a34a;font-weight:900}}
.thumbA{{display:inline-flex;cursor:grab}}
.commentText{{font-weight:400;line-height:1.24;white-space:normal}}
.tagReminder{{color:#b91c1c;font-weight:900}}
.thumbAWrap{{position:relative;display:inline-flex;touch-action:none;max-width:100%;align-items:flex-start}}
.thumbAWrap.dragging{{opacity:.7;z-index:5}}
.thumbAWrap.resizing{{outline:2px solid #60a5fa;outline-offset:1px}}
.thumbHandle{{position:absolute;right:4px;bottom:4px;width:14px;height:14px;border:2px solid rgba(15,23,42,.45);border-top:none;border-left:none;cursor:nwse-resize;background:rgba(255,255,255,.7)}}
.thumbRemove{{position:absolute;top:2px;right:2px;width:18px;height:18px;border:none;border-radius:999px;background:rgba(15,23,42,.72);color:#fff;font-weight:900;line-height:18px;padding:0;cursor:pointer}}
.thumbRemove:hover{{background:#dc2626}}
.colComment br + br{{display:none}}
.compactRow .colComment{{line-height:1.22}}
.compactRow .colComment .entryComment{{margin-top:6px}}

@media print{{
  .page{{height:auto;min-height:0}}
  .pageContent{{padding:8mm 7mm 30mm 7mm}}
  .crTable th, .crTable td{{padding:5px 6px}}
  .zoneTitle{{padding:5px 7px}}
  .reportHeader{{margin-bottom:6px}}
  .thumb{{height:72px;max-width:130px}}
  .thumbHandle{{display:none}}
  .thumbRemove{{display:none}}
  .btnAddImage{{display:none}}
}}
.annexTable{{width:100%;border-collapse:collapse;font-size:12px;table-layout:fixed;border:1px solid var(--border)}}
.annexTable thead{{display:table-header-group}}
.annexTable th,.annexTable td{{border-bottom:1px solid var(--border);padding:8px 6px;text-align:left;vertical-align:top}}
.annexTable td:first-child{{width:90px;color:#2563eb;font-weight:900}}
.annexTable td:last-child{{text-align:right}}
.annexTable td:last-child .annexLink{{display:inline-block;text-align:right}}
.annexTable th{{font-weight:900;background:#1f4e4f;color:#fff}}
.annexTable .annexLink{{color:#f97316;font-weight:800;text-decoration:underline;text-underline-offset:3px;cursor:pointer}}
.annexTable .annexLink::after{{content:" ↗";font-weight:900;color:#f97316}}
.annexTable tr:last-child td{{border-bottom:none}}
.coverTable{{margin:10px 0 12px 0}}
.coverTable td:first-child{{width:260px;color:#0b1220;font-weight:900}}
.coverTable td.kpiNum{{text-align:right;font-weight:1000}}
.coverTable .chips{{display:flex;flex-wrap:wrap;gap:8px}}
.coverTable .chip{{display:inline-flex;align-items:center;gap:8px;border:1px solid var(--border);border-radius:999px;padding:6px 10px;font-weight:800;background:#fff}}
.coverNote{{margin-top:12px;border:1px solid var(--border);border-radius:14px;padding:12px;background:#fff;line-height:1.5}}
.coverNoteTitle{{font-weight:1000;margin-bottom:6px}}
.reportHeader{{font-family:"Arial Nova Cond Light","Arial Narrow",Arial,sans-serif;font-size:11px;font-weight:400;color:#0b1220;text-align:center;margin:0 0 10px 0;}}
@media print{{.printHeaderFixed{{position:sticky;top:0;background:#fff;padding:1mm 0;z-index:20;}}}}
.reportHeader .accent{{color:#f59e0b;font-weight:900}}
.presenceTable .presenceList{{margin:0;padding-left:0;list-style:none;display:flex;flex-direction:column;gap:6px}}
.presenceTable .presenceLine{{display:flex;align-items:center;gap:8px;font-weight:700}}
.docFooter{{position:absolute;left:0;right:0;bottom:0;height:24mm;display:grid;grid-template-columns:120px 1fr 120px;align-items:center;gap:10px;padding:3mm 10mm;border-top:1px solid #dbe5f0;background:#fff;overflow:hidden;width:100%;box-sizing:border-box}}
.docFooter::before{{content:"";position:absolute;left:0;bottom:0;width:170px;height:42px;background:#123f45;clip-path:polygon(0 100%,100% 100%,0 0)}}
.docFooter::after{{content:"";position:absolute;right:0;bottom:0;width:260px;height:70px;background:#123f45;clip-path:polygon(100% 0,100% 100%,0 100%)}}
.footLeft,.footCenter,.footRight{{position:relative;z-index:2}}
.footLeft{{justify-self:start}}
.footCenter{{text-align:center;justify-self:center}}
.footRight{{justify-self:end;width:120px;display:flex;justify-content:flex-end}}
.pageNum{{font-family:'Arial Nova Cond Light','Arial Narrow',Arial,sans-serif;font-size:13px;font-weight:700;color:rgba(255,255,255,.82);line-height:1;letter-spacing:.2px;padding-right:14px;padding-bottom:6px}}
.tempoLegal{{font-family:"Arial Nova Cond Light","Arial Narrow",Arial,sans-serif;font-size:10px;line-height:1.3;color:#6b7280;font-weight:600}}
.footImg{{display:block;max-height:32px;width:auto}}
.footMark{{max-height:48px}}
.footRythme{{max-height:28px;margin:6px auto 0 auto}}
.footTempo{{max-height:28px;margin-left:auto}}
@media print{{body{{padding:0}} .actions,.rangePanel,.constraintsPanel{{display:none!important}} .page{{width:210mm;min-height:297mm;margin:0;box-shadow:none;break-after:page;page-break-after:always;}} .page:last-child{{break-after:auto;page-break-after:auto;}}}}

{EDITOR_MEMO_MODAL_CSS}
{QUALITY_MODAL_CSS}
{ANALYSIS_MODAL_CSS}
"""

    # Banner / cover HTML
    bg_style = (
        f"background-image:url('{_escape(proj_img)}');"
        if proj_img.startswith("http")
        else "background:linear-gradient(90deg,#cfe8ff,#ffffff);"
    )

    tempo_logo = _logo_data_url(LOGO_TEMPO_PATH)
    logo_rythme = _logo_data_url(LOGO_RYTHME_PATH)
    logo_tmark = _logo_data_url(LOGO_T_MARK_PATH)
    cover_html = ""

    next_meeting_date = (meet_date or ref_date) + timedelta(days=7)
    next_meeting_date_txt = next_meeting_date.strftime("%d/%m/%Y")
    cr_date_txt = (meet_date or ref_date).strftime("%d/%m/%Y")

    cover_hero_html = f"""
      <div class='coverHero'>
        <div class='coverHeroImg' style="{bg_style}">
          <div class='coverHeroLogoWrap'>{("<img class='coverHeroLogo' src='" + tempo_logo + "' alt='TEMPO' />") if tempo_logo else ""}</div>
          <div class='coverHeroFade'></div>
        </div>
        <div class='coverHeroCurve'></div>
      </div>
    """

    cover_note_html = f"""
      <div class='coverNoteCenter'>
        <div class='coverProjectTitle' contenteditable='true'>{_escape(project)}</div>
        <div class='coverCrTitle' contenteditable='true'>CR REUNION DE SYNTHESE TECHNIQUE</div>
        <div class='coverCrMeta'>
          N°<span contenteditable='true' class='editInline' data-sync='cr-number'>{_escape(cr_number_default)}</span>
          du <strong>{_escape(cr_date_txt)}</strong>
        </div>
        <div class='nextMeetingBox'>
          <div class='nextMeetingLine1'>La prochaine réunion de synthèse est fixée au</div>
          <div class='nextMeetingLine2'>
            <span contenteditable='true' class='editInline'>{_escape(next_meeting_date_txt)}</span>
            à
            <span contenteditable='true' class='editInline'>14h00</span>
          </div>
          <div contenteditable='true' class='nextMeetingLine3'>BASE VIE — adresse à compléter</div>
        </div>
        <div class='coverAppNote'>
          Téléchargez gratuitement l’application de gestion de projet METRONOME. L’application développée par TEMPO
          dédiée à la gestion de projet. Celle-ci vous permettra de retrouver l’intégralité des réunions de synthèse, comptes rendu,
          planning et suivi des tâches depuis votre smartphone ou votre ordinateur.
        </div>
        <a class='coverUrl' href='https://app.atelier-tempo.fr' target='_blank' rel='noopener'>app.atelier-tempo.fr</a>
      </div>
    """

    report_header_html = f"""
      <div class='reportHeader printHeaderFixed'>
        {_escape(project)} <span class='accent'>— Compte Rendu</span> n°<span contenteditable='true' class='editInline' data-sync='cr-number'>{_escape(cr_number_default)}</span> — Réunion de Synthèse du {_escape(cr_date_txt)}
      </div>
    """

    top_html = f"""
      <div class="topPage">
        {cover_hero_html}
        {cover_note_html}
      </div>
    """
    annexes_html = ""
    try:
        docs = get_documents().copy()
        if not docs.empty:
            meeting_col = next((c for c in docs.columns if "Meeting/ID" in str(c)), None)
            project_col = next((c for c in docs.columns if "Project/Title" in str(c)), None)
            title_col = next((c for c in docs.columns if "Title" in str(c)), None)
            url_col = next((c for c in docs.columns if "URL" in str(c)), None)
            if not url_col:
                url_col = next((c for c in docs.columns if "Link" in str(c)), None)
            if meeting_col:
                docs = docs.loc[docs[meeting_col].astype(str) == str(meeting_id)].copy()
            elif project_col:
                docs = docs.loc[docs[project_col].fillna("").astype(str).str.strip() == project].copy()
            items = []
            for _, r in docs.iterrows():
                title = _escape(r.get(title_col, "") if title_col else r.get("Title", ""))
                url = _escape(r.get(url_col, "") if url_col else "")
                if title or url:
                    link = (
                        f"<a class='annexLink' href='{url}' target='_blank' rel='noopener'>{title or url}</a>"
                        if url
                        else "—"
                    )
                    label = f"{len(items) + 1}."
                    items.append(
                        f"""
              <tr>
                <td>{label} Annexe</td>
                <td>{link}</td>
              </tr>
                        """
                    )
            if items:
                annexes_html = f"""
      <div class="section reportBlock">
        <table class="annexTable">
          <thead>
            <tr>
              <th>Document</th>
              <th>Lien</th>
            </tr>
          </thead>
          <tbody>
            {''.join(items)}
          </tbody>
        </table>
      </div>
                """
    except MissingDataError:
        annexes_html = ""

    return f"""
<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>CR Synthèse — {_escape(project)} — {_escape(date_txt)}</title>
<style>
{css}
</style>
</head>
<body class="{'pdf' if print_mode else ''}">
  {actions_html}
  <div class="wrap">
    <section class="page page--cover">
      <div class="pageContent">
        {cover_html}
        {top_html}
      </div>
      <div class="docFooter">
        <div class="footLeft">{"<img class='footImg footMark' src='" + logo_tmark + "' alt='' />" if logo_tmark else ""}</div>
        <div class="footCenter"><div style="font-family:'Arial Nova Cond Light','Arial Narrow',Arial,sans-serif;font-size:12px;font-weight:700;color:#111">TEMPO</div><div class="tempoLegal">35, rue Beaubourg, 75003 Paris<br/>SAS au capital de 1 000 Euros - RCS Créteil N° 892 046 301 - APE 7112 B</div>{("<img class='footImg footRythme' src='" + logo_rythme + "' alt='' />") if logo_rythme else ""}</div>
        <div class="footRight"><span class="pageNum"></span></div>
      </div>
    </section>

    <div class="reportPages">
      <section class="page page--report">
        <div class="pageContent">
          <div class="reportTables">
            {report_header_html}
            {presence_html}
            <div class="reportBlocks">
              {zones_html}
              {annexes_html}
              {report_note_html}
            </div>
          </div>
        </div>
        <div class="docFooter">
          <div class="footLeft">{"<img class='footImg footMark' src='" + logo_tmark + "' alt='' />" if logo_tmark else ""}</div>
          <div class="footCenter"><div style="font-family:'Arial Nova Cond Light','Arial Narrow',Arial,sans-serif;font-size:12px;font-weight:700;color:#111">TEMPO</div><div class="tempoLegal">35, rue Beaubourg, 75003 Paris<br/>SAS au capital de 1 000 Euros - RCS Créteil N° 892 046 301 - APE 7112 B</div>{("<img class='footImg footRythme' src='" + logo_rythme + "' alt='' />") if logo_rythme else ""}</div>
          <div class="footRight"><span class="pageNum"></span></div>
        </div>
      </section>
    </div>
  </div>

  <template id="report-page-template">
    <section class="page page--report">
      <div class="pageContent">
        <div class="reportTables">
          {report_header_html}
          <div class="reportBlocks"></div>
        </div>
      </div>
      <div class="docFooter">
        <div class="footLeft">{"<img class='footImg footMark' src='" + logo_tmark + "' alt='' />" if logo_tmark else ""}</div>
        <div class="footCenter"><div style="font-family:'Arial Nova Cond Light','Arial Narrow',Arial,sans-serif;font-size:12px;font-weight:700;color:#111">TEMPO</div><div class="tempoLegal">35, rue Beaubourg, 75003 Paris<br/>SAS au capital de 1 000 Euros - RCS Créteil N° 892 046 301 - APE 7112 B</div>{("<img class='footImg footRythme' src='" + logo_rythme + "' alt='' />") if logo_rythme else ""}</div>
        <div class="footRight"><span class="pageNum"></span></div>
      </div>
    </section>
  </template>

{EDITOR_MEMO_MODAL_HTML}
{QUALITY_MODAL_HTML}
{ANALYSIS_MODAL_HTML}
<script>{EDITOR_MEMO_MODAL_JS}</script>
<script>{QUALITY_MODAL_JS}</script>
<script>{ANALYSIS_MODAL_JS}</script>
<script>{SYNC_EDITABLE_JS}</script>
<script>{RANGE_PICKER_JS}</script>
<script>{PRINT_PREVIEW_TOGGLE_JS}</script>
<script>{CONSTRAINT_TOGGLES_JS}</script>
<script>{LAYOUT_CONTROLS_JS}</script>
<script>{DRAGGABLE_IMAGES_JS}</script>
<script>{PAGINATION_JS}</script>
<script>{PRINT_OPTIMIZE_JS}</script>
<script>{ROW_CONTROL_JS}</script>
<script>{RESIZE_COLUMNS_JS}</script>
<script>{RESIZE_TOP_JS}</script>
</body>
</html>
"""


# -------------------------
# ROUTES
# -------------------------
@app.get("/", response_class=HTMLResponse)
def home(project: Optional[str] = Query(default=None)):
    try:
        return HTMLResponse(render_home(project=project))
    except MissingDataError as err:
        return HTMLResponse(render_missing_data_page(err), status_code=503)


@app.get("/cr", response_class=HTMLResponse)
def cr(
    meeting_id: str = Query(...),
    project: str = Query(default=""),
    print: int = Query(default=0),
    pinned_memos: str = Query(default=""),
    range_start: str = Query(default=""),
    range_end: str = Query(default=""),
):
    try:
        return HTMLResponse(
            render_cr(
                meeting_id=meeting_id,
                project=project,
                print_mode=bool(print),
                pinned_memos=pinned_memos,
                range_start=range_start,
                range_end=range_end,
            )
        )
    except MissingDataError as err:
        return HTMLResponse(render_missing_data_page(err), status_code=503)


@app.get("/health", response_class=JSONResponse)
def health():
    data = {}
    for k, p in [
        ("entries", ENTRIES_PATH),
        ("meetings", MEETINGS_PATH),
        ("companies", COMPANIES_PATH),
        ("projects", PROJECTS_PATH),
    ]:
        try:
            ok = os.path.exists(p)
            mt = _mtime(p)
            data[k] = {"path": p, "exists": ok, "mtime": mt}
        except Exception as e:
            data[k] = {"path": p, "exists": False, "error": str(e)}
    return {"ok": True, "files": data}


@app.get("/api/memos", response_class=JSONResponse)
def api_memos(project: str = Query("", alias="project"), area: str = Query("", alias="area")):
    """
    Return list of MEMOS for a given project and area.
    IMPORTANT: the memo must belong to the zones of THE SAME PROJECT.
    """
    try:
        e = get_entries().copy()
        e = e.loc[e[E_COL_PROJECT_TITLE].fillna("").astype(str).str.strip() == str(project).strip()].copy()
        e = _explode_areas(e)
        if area:
            e = e.loc[e["__area_list__"].astype(str) == str(area)].copy()

        e["__is_task__"] = _series(e, E_COL_IS_TASK, False).apply(_bool_true)
        e = e.loc[e["__is_task__"] == False].copy()

        e["__created__"] = _series(e, E_COL_CREATED, None).apply(_parse_date_any)
        e = e.sort_values(by=["__created__"], ascending=[False])

        items = []
        for _, r in e.iterrows():
            rid = str(r.get(E_COL_ID, "")).strip()
            if not rid:
                continue
            items.append(
                {
                    "id": rid,
                    "title": str(r.get(E_COL_TITLE, "") or "").strip(),
                    "created": _fmt_date(_parse_date_any(r.get(E_COL_CREATED))),
                    "company": str(r.get(E_COL_COMPANY_TASK, "") or "").strip(),
                    "owner": str(r.get(E_COL_OWNER, "") or "").strip(),
                }
            )
        return {"items": items}
    except MissingDataError as err:
        return JSONResponse(
            {"error": str(err), "label": err.label, "path": err.path, "env_var": err.env_var},
            status_code=503,
        )
    except Exception as ex:
        return JSONResponse({"error": str(ex)}, status_code=500)


def _quality_payload(
    text: str,
    language: str = "fr",
    ignore_terms: Optional[set[str]] = None,
) -> Dict[str, object]:
    cleaned_text = re.sub(r"\bnan\b", "", text, flags=re.IGNORECASE).strip()
    if not cleaned_text:
        return {"score": 100, "total": 0, "issues": []}
    ignore_terms = {t.lower() for t in (ignore_terms or set()) if t}
    url = "https://api.languagetool.org/v2/check"
    data = urllib.parse.urlencode({"language": language, "text": cleaned_text}).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    with urllib.request.urlopen(req, timeout=10) as resp:
        payload = json.loads(resp.read().decode("utf-8"))
    matches = payload.get("matches", [])
    words = max(1, len(re.findall(r"\w+", cleaned_text)))
    errors = 0
    score = max(0, int(100 - (errors / words) * 100))
    issues = []
    for m in matches:
        offset = m.get("offset")
        length = m.get("length")
        match_text = (
            cleaned_text[offset : offset + length] if offset is not None and length is not None else ""
        )
        match_text_stripped = match_text.strip()
        if not match_text_stripped:
            continue
        match_lower = match_text_stripped.lower()
        if match_lower == "nan" or match_lower in ignore_terms:
            continue
        if match_text_stripped.isupper() and len(match_text_stripped) > 2:
            continue
        if match_text_stripped.istitle() and len(match_text_stripped) > 2:
            continue
        context = m.get("context", {}) or {}
        repl = ", ".join([r.get("value", "") for r in m.get("replacements", []) if r.get("value")])
        category = (m.get("rule", {}) or {}).get("category", {}) or {}
        errors += 1
        issues.append(
            {
                "message": m.get("message", ""),
                "context": context.get("text", ""),
                "context_offset": context.get("offset"),
                "context_length": context.get("length"),
                "replacements": repl,
                "category": category.get("name", ""),
                "offset": offset,
                "length": length,
                "text": cleaned_text,
            }
        )
    score = max(0, int(100 - (errors / words) * 100))
    return {"score": score, "total": errors, "issues": issues}


@app.get("/api/quality", response_class=JSONResponse)
def api_quality(
    meeting_id: str = Query(...),
    project: str = Query(default=""),
):
    try:
        mrow = meeting_row(meeting_id)
        edf = entries_for_meeting(meeting_id)
        project = (project or str(mrow.get(M_COL_PROJECT_TITLE, ""))).strip()
        ref_date = _parse_date_any(mrow.get(M_COL_DATE)) or date.today()
        rem_df = reminders_for_project(project_title=project, ref_date=ref_date, max_level=8)
        fol_df = followups_for_project(project_title=project, ref_date=ref_date, exclude_entry_ids=set())

        def _items(df: pd.DataFrame, ignore_terms: set[str]) -> List[Dict[str, str]]:
            if df.empty:
                return []
            df = _explode_areas(df.copy())
            out = []
            for _, r in df.iterrows():
                title = str(r.get(E_COL_TITLE, "") or "").strip()
                comment = str(r.get(E_COL_TASK_COMMENT_TEXT, "") or "").strip()
                text = " ".join([t for t in [title, comment] if t]).strip()
                text = re.sub(r"\bnan\b", "", text, flags=re.IGNORECASE).strip()
                if not text:
                    continue
                area = str(r.get("__area_list__", "Général"))
                ignore_terms.add(area.lower())
                ignore_terms.update(_split_words(area))
                out.append({"area": area, "text": text})
            return out
        ignore_terms: set[str] = set()
        if project:
            ignore_terms.add(project.lower())
            ignore_terms.update(_split_words(project))
        company_terms = pd.concat(
            [
                edf.get(E_COL_COMPANY_TASK, pd.Series(dtype=str)),
                rem_df.get(E_COL_COMPANY_TASK, pd.Series(dtype=str)),
                fol_df.get(E_COL_COMPANY_TASK, pd.Series(dtype=str)),
            ],
            ignore_index=True,
        )
        owner_terms = pd.concat(
            [
                edf.get(E_COL_OWNER, pd.Series(dtype=str)),
                rem_df.get(E_COL_OWNER, pd.Series(dtype=str)),
                fol_df.get(E_COL_OWNER, pd.Series(dtype=str)),
            ],
            ignore_index=True,
        )
        for val in pd.concat([company_terms, owner_terms], ignore_index=True).dropna().astype(str):
            ignore_terms.add(val.lower())
            ignore_terms.update(_split_words(val))
        items = _items(edf, ignore_terms) + _items(rem_df, ignore_terms) + _items(fol_df, ignore_terms)
        issues_by_area: Dict[str, List[Dict[str, object]]] = {}
        total_errors = 0
        total_words = 0
        for it in items:
            payload = _quality_payload(it["text"], language="fr", ignore_terms=ignore_terms)
            total_errors += int(payload.get("total", 0))
            cleaned = re.sub(r"\bnan\b", "", it["text"], flags=re.IGNORECASE)
            total_words += max(1, len(re.findall(r"\w+", cleaned)))
            if payload.get("issues"):
                issues_by_area.setdefault(it["area"], []).extend(payload["issues"])
        score = max(0, int(100 - (total_errors / max(1, total_words)) * 100))
        return {"score": score, "total": total_errors, "issues_by_area": issues_by_area}
    except MissingDataError as err:
        return JSONResponse(
            {"error": str(err), "label": err.label, "path": err.path, "env_var": err.env_var},
            status_code=503,
        )
    except Exception as ex:
        return JSONResponse({"error": str(ex)}, status_code=500)


@app.get("/api/analysis", response_class=JSONResponse)
def api_analysis(
    meeting_id: str = Query(...),
    project: str = Query(default=""),
):
    try:
        mrow = meeting_row(meeting_id)
        edf = entries_for_meeting(meeting_id)
        project = (project or str(mrow.get(M_COL_PROJECT_TITLE, ""))).strip()
        ref_date = _parse_date_any(mrow.get(M_COL_DATE)) or date.today()
        rem_df = reminders_for_project(project_title=project, ref_date=ref_date, max_level=8)
        fol_df = followups_for_project(project_title=project, ref_date=ref_date, exclude_entry_ids=set())

        is_task = _series(edf, E_COL_IS_TASK, False).apply(_bool_true)
        tasks = edf[is_task].copy()
        completed = _series(tasks, E_COL_COMPLETED, False).apply(_bool_true)
        open_tasks = tasks[~completed]

        points = []
        risks = []
        follow_ups = []
        late_tasks = len(rem_df)
        followups = len(fol_df)

        if late_tasks:
            points.append(f"{late_tasks} rappel(s) en retard à la date de séance.")
            risks.append("Retards critiques à prioriser avant la prochaine réunion.")
        if len(open_tasks):
            points.append(f"{len(open_tasks)} tâche(s) ouverte(s) dans la séance.")
        if followups:
            follow_ups.append(f"{followups} tâche(s) à suivre sur le projet.")

        if not points:
            points.append("Aucun point bloquant détecté dans la séance.")

        least_responsive = reminders_by_company(rem_df)[:5]
        followups_by_area = {}
        if not fol_df.empty:
            for area, g in fol_df.groupby("__area_list__", dropna=False):
                titles = g.get(E_COL_TITLE, pd.Series([], dtype=str)).fillna("").astype(str).tolist()
                followups_by_area[str(area)] = [t for t in titles if t.strip()][:6]

        return {
            "kpis": {"late_tasks": late_tasks, "open_tasks": int(len(open_tasks)), "followups": followups},
            "points": points,
            "risks": risks,
            "follow_ups": follow_ups,
            "least_responsive": least_responsive,
            "followups_by_area": followups_by_area,
        }
    except MissingDataError as err:
        return JSONResponse(
            {"error": str(err), "label": err.label, "path": err.path, "env_var": err.env_var},
            status_code=503,
        )
    except Exception as ex:
        return JSONResponse({"error": str(ex)}, status_code=500)
