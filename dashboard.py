"""
Motor Test Bench Dashboard — v9
Google Sheets sync + Streamlit Cloud deployment.
Run locally:  py -m streamlit run dashboard.py
Run on cloud: push to GitHub, deploy on share.streamlit.io
"""

import streamlit as st
import pandas as pd
import json
import os
import struct
import tempfile
import time
from datetime import datetime, timedelta
from pathlib import Path

# ─── CONFIG ────────────────────────────────────────────────────────────────────
# Use __file__ so all paths are absolute and work regardless of how the app
# is launched (VS Code, Windows shortcut, batch file, etc.)
_BASE_DIR        = os.path.dirname(os.path.abspath(__file__))
DATA_FILE        = os.path.join(_BASE_DIR, "bench_data.json")
CREDENTIALS_FILE = os.path.join(_BASE_DIR, "credentials.json")
SUPPORTED_EXTS   = {".edvm", ".csv", ".wnq"}
RPM_THRESHOLD    = 100
DRIVE_FOLDER_ID  = "1Jl-3A2dO7_m4Q4G_hRRMiE4XabcxOPbg"

# ─── GOOGLE SHEETS LAYER ───────────────────────────────────────────────────────
#
# HOW IT WORKS
# ─────────────
# All dashboard data (sessions, components, rig config, events, brands) is stored
# as a single JSON string in cell A1 of a Google Sheet called "bench_data".
# Every save writes to that cell. Every load reads from it.
# This means ANY computer — including Streamlit Cloud — always sees live data.
#
# SECURITY MODEL
# ──────────────
# A "Service Account" is a robot Google account that belongs only to this app.
# You share ONLY the specific spreadsheet with it — nothing else in your Drive
# is accessible. The credentials are stored in Streamlit Secrets, never in code.
#
# HOW TO SET UP (one time, ~20 minutes)
# ──────────────────────────────────────
# STEP 1 — Create a Google Sheet
#   a. Go to sheets.google.com → create a new sheet
#   b. Name it anything, e.g. "AIR TestBench Data"
#   c. Copy the Spreadsheet ID from the URL:
#      https://docs.google.com/spreadsheets/d/  ← THIS PART →  /edit
#      It looks like: 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs
#
# STEP 2 — Create a Service Account
#   a. Go to console.cloud.google.com → sign in with company account
#   b. Create a new project (or reuse your existing TestBench project)
#   c. Search "Google Sheets API" → Enable it
#   d. Go to APIs & Services → Credentials → + Create Credentials → Service Account
#   e. Name: testbench-app  → Create and Continue → Done
#   f. Click the service account → Keys tab → Add Key → JSON → Create
#   g. A .json file downloads. Keep it safe — it's the app's password.
#
# STEP 3 — Share the sheet with the service account
#   a. Open the downloaded .json file in Notepad
#   b. Find "client_email" — copy that address (ends in .iam.gserviceaccount.com)
#   c. Open your Google Sheet → Share → paste that email → Editor → Send
#
# STEP 4A — Running LOCALLY
#   a. Rename the downloaded .json to "credentials.json"
#   b. Put it in the same folder as dashboard.py
#   c. Set SHEET_ID below to your spreadsheet ID
#   SHEET_ID = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs"  ← put yours here
#
# STEP 4B — Running on STREAMLIT CLOUD (for management to view)
#   a. Push dashboard.py to a GitHub repo (public or private)
#   b. Go to share.streamlit.io → New app → connect your repo
#   c. In Streamlit Cloud → App settings → Secrets → paste this:
#
#      [gsheets]
#      spreadsheet_id = "your-sheet-id-here"
#      [gsheets.credentials]
#      type = "service_account"
#      project_id = "..."         ← copy from your .json file
#      private_key_id = "..."     ← copy from your .json file
#      private_key = "..."        ← copy from your .json file (the long -----BEGIN RSA...)
#      client_email = "..."       ← copy from your .json file
#      client_id = "..."          ← copy from your .json file
#      token_uri = "https://oauth2.googleapis.com/token"
#
#   d. The app will read secrets automatically — no credentials.json needed
#
# ─── SET YOUR SHEET ID HERE ───────────────────────────────────────────────────
SHEET_ID = "1btpyqWJaLl9YeelLbuvlJJilLBzCwWidSs_fGxQ-W8k"
#              e.g. SHEET_ID = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs"
# ──────────────────────────────────────────────────────────────────────────────

def _get_sheets_client():
    """
    Returns an authenticated gspread client.
    Tries Streamlit Secrets first (for cloud), then credentials.json (for local).
    """
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        return None

    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

    # ── Try Streamlit Secrets (cloud deployment) ──────────────────────────────
    try:
        sec = st.secrets.get("gsheets", {})
        cred_info = dict(sec.get("credentials", {}))
        if cred_info:
            # Fix private_key newlines (Streamlit strips them)
            if "private_key" in cred_info:
                cred_info["private_key"] = cred_info["private_key"].replace("\\n", "\n")
            creds = Credentials.from_service_account_info(cred_info, scopes=SCOPES)
            return gspread.authorize(creds)
    except Exception:
        pass

    # ── Try local credentials.json ────────────────────────────────────────────
    try:
        if os.path.exists(CREDENTIALS_FILE):
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
            return gspread.authorize(creds)
    except Exception:
        pass

    return None


def _get_sheet_id():
    """Get sheet ID from Secrets (cloud) or SHEET_ID constant (local)."""
    try:
        sid = st.secrets.get("gsheets", {}).get("spreadsheet_id", "")
        if sid:
            return sid
    except Exception:
        pass
    return SHEET_ID


def sheets_load():
    """Load data from Google Sheets cell A1. Returns None if unavailable."""
    try:
        gc = _get_sheets_client()
        sid = _get_sheet_id()
        if not gc or not sid:
            return None
        sh    = gc.open_by_key(sid)
        ws    = sh.worksheet("bench_data")
        raw   = ws.acell("A1").value
        if raw:
            return json.loads(raw)
        return {}
    except Exception:
        return None


def sheets_save(data):
    """Write entire data dict as JSON to Google Sheets cell A1."""
    try:
        gc = _get_sheets_client()
        sid = _get_sheet_id()
        if not gc or not sid:
            return False
        sh  = gc.open_by_key(sid)
        # Create worksheet if it doesn't exist
        try:
            ws = sh.worksheet("bench_data")
        except Exception:
            ws = sh.add_worksheet("bench_data", rows=1, cols=1)
        ws.update("A1", [[json.dumps(data)]])
        return True
    except Exception:
        return False

DEFAULT_BRANDS = {
    "nidec": "Motor", "reb30": "Motor", "reb60": "Motor",
    "mgm": "ESC", "elmo": "ESC",
    "hilex": "Prop", "helix": "Prop", "e-prop": "Prop", "eprop": "Prop",
}

# ─── DATA HELPERS ──────────────────────────────────────────────────────────────

DEFAULTS = {
    "components": [], "sessions": [], "imported_files": [],
    "last_sync": None, "brands": DEFAULT_BRANDS.copy(),
    "rig_config": {},           # legacy single config
    "rig_side_a": {},           # Side A component config
    "rig_side_b": {},           # Side B component config
    "active_side": "Side A",    # "Side A" | "Side B" | "Both"
    "events": [],
    "asana": {"token": "", "project_gid": "1212056850873170", "section_filter": ""},
}

def _ensure_keys(d):
    for k, v in DEFAULTS.items():
        if k not in d:
            d[k] = v if not isinstance(v, dict) else v.copy()
    return d

def load_data():
    """
    Load order:
      1. Google Sheets — live, works on any computer
      2. Local bench_data.json — fallback if Sheets unreachable
    """
    sheets_data = sheets_load()
    if sheets_data is not None:
        d = _ensure_keys(sheets_data)
        # Keep a local backup
        try:
            with open(DATA_FILE, "w") as f:
                json.dump(d, f, indent=2)
        except Exception:
            pass
        return d
    # Fallback: local file
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r") as f:
                return _ensure_keys(json.load(f))
        except Exception:
            pass
    return _ensure_keys({})


def save_data(data):
    """
    Save to Google Sheets (primary) AND local file (backup).
    """
    sheets_save(data)
    try:
        with open(DATA_FILE, "w") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass

def get_comp_hours(comp_id, sessions):
    return sum(s["hours"] for s in sessions if comp_id in s.get("components", []))

def get_comp_sessions(comp_id, sessions):
    return [s for s in sessions if comp_id in s.get("components", [])]

def comp_status(hours, limit):
    if not limit:
        return "ok", "No limit set"
    pct = hours / limit
    if pct >= 1.0:  return "danger",  "Needs maintenance"
    if pct >= 0.75: return "warning", "Due soon"
    return "ok", "OK"

def fmt_dur(hours):
    if hours is None or (isinstance(hours, float) and hours != hours):
        return "—"
    total_sec = int(round(float(hours) * 3600))
    hh = total_sec // 3600
    mm = (total_sec % 3600) // 60
    ss = total_sec % 60
    return f"{hh:02d}:{mm:02d}:{ss:02d}"

def fmt_label(hours):
    total_min = int(hours * 60)
    if total_min < 1:   return "< 1m"
    if total_min < 60:  return f"{total_min}m"
    return f"{total_min // 60}h {total_min % 60}m"

# ─── TIME RANGE HELPERS ────────────────────────────────────────────────────────

def hours_in_range(sessions, start_dt, end_dt):
    total = 0.0
    for s in sessions:
        try:
            ts = datetime.fromisoformat(s.get("timestamp", ""))
            if start_dt <= ts < end_dt:
                total += s.get("hours", 0)
        except Exception:
            pass
    return total

def sessions_in_range(sessions, start_dt, end_dt):
    result = []
    for s in sessions:
        try:
            ts = datetime.fromisoformat(s.get("timestamp", ""))
            if start_dt <= ts < end_dt:
                result.append(s)
        except Exception:
            pass
    return result

def today_range():
    now = datetime.now()
    start = now.replace(hour=0, minute=0, second=0, microsecond=0)
    return start, start + timedelta(days=1)

def week_range():
    now = datetime.now()
    # Sun-Fri week: find last Sunday
    days_since_sunday = (now.weekday() + 1) % 7
    start = (now - timedelta(days=days_since_sunday)).replace(
        hour=0, minute=0, second=0, microsecond=0)
    end = start + timedelta(days=6)  # Saturday is end
    return start, end

def month_range():
    now = datetime.now()
    start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    if now.month == 12:
        end = start.replace(year=now.year + 1, month=1)
    else:
        end = start.replace(month=now.month + 1)
    return start, end

# ─── MANUFACTURER HELPERS ──────────────────────────────────────────────────────

def get_manufacturer(comp_name):
    """Extract brand name from component name like 'Nidec_FX124' → 'Nidec'"""
    return comp_name.split("_")[0] if "_" in comp_name else comp_name

def build_manufacturer_chart_data(data):
    """Build chart rows grouped by manufacturer + type"""
    rows = []
    for comp in data["components"]:
        hours    = get_comp_hours(comp["id"], data["sessions"])
        sessions = len(get_comp_sessions(comp["id"], data["sessions"]))
        mfr      = get_manufacturer(comp["name"])
        rows.append({
            "manufacturer": mfr,
            "type":         comp["type"],
            "name":         comp["name"],
            "hours":        hours,
            "sessions":     sessions,
            "id":           comp["id"],
            "limit":        comp.get("limit"),
        })
    return pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["manufacturer","type","name","hours","sessions","id","limit"])

# ─── RIG SIDE HELPERS ─────────────────────────────────────────────────────────

def get_side_component_ids(data, side):
    """
    Return list of component IDs currently installed on a given side.
    side: "Side A" | "Side B" | "Both"
    Looks up component names from rig config and matches them to known components.
    Supports ESC, ESC2 (dual ESC per side).
    """
    cfg_fields = ["Motor", "ESC", "ESC2", "Prop", "Prop2", "Battery"]
    side_ids   = []

    def resolve(cfg):
        ids = []
        for field in cfg_fields:
            name = cfg.get(field, "").strip()
            if not name:
                continue
            for c in data["components"]:
                if c["name"].lower() == name.lower():
                    ids.append(c["id"])
                    break
        return ids

    if side in ("Side A", "Both"):
        side_ids += resolve(data.get("rig_side_a", {}))
    if side in ("Side B", "Both"):
        side_ids += resolve(data.get("rig_side_b", {}))
    return list(dict.fromkeys(side_ids))  # deduplicate, preserve order

# ─── RPM-BASED RUN TIME ────────────────────────────────────────────────────────

def rpm_active_seconds_from_array(rpm_array, time_array, threshold=RPM_THRESHOLD):
    import numpy as np
    rpm   = np.array(rpm_array, dtype=float)
    times = np.array(time_array, dtype=float)
    if len(rpm) < 2:
        return 0.0
    active = rpm >= threshold
    dt = np.diff(times)
    return float(np.sum(dt[active[:-1] & active[1:]]))

def rpm_duration_from_edvm(filepath):
    try:
        import scipy.io as sio
        mat       = sio.loadmat(filepath)
        time_data = mat["Signal01"][0][0][0].flatten()
        for sig in ["Signal02", "Signal03"]:
            try:
                c = mat[sig][0][0][0].flatten()
                if c.max() > RPM_THRESHOLD:
                    return rpm_active_seconds_from_array(c, time_data)
            except Exception:
                continue
        return float(time_data[-1] - time_data[0])
    except Exception:
        return None

def rpm_duration_from_csv(filepath):
    try:
        sep, enc = ";", "latin-1"
        for e in ["latin-1", "utf-8", "cp1252"]:
            try:
                df = pd.read_csv(filepath, sep=";", encoding=e, nrows=3)
                if len(df.columns) > 1:
                    enc, sep = e, ";"; break
                df = pd.read_csv(filepath, sep=",", encoding=e, nrows=3)
                if len(df.columns) > 1:
                    enc, sep = e, ","; break
            except Exception:
                continue
        df      = pd.read_csv(filepath, sep=sep, encoding=enc)
        rpm_col = next((c for c in df.columns if "revolution" in c.lower() or "rpm" in c.lower()), None)
        if rpm_col is None:
            return None
        esc_col  = next((c for c in df.columns if "ESC time" in c), None)
        time_col = next((c for c in df.columns if c.strip().lower() in ["time", "time (sec)"]), None)
        if esc_col:
            times = df[esc_col].values.astype(float)
        elif time_col:
            t = pd.to_datetime(df[time_col], format="%H:%M:%S.%f", errors="coerce")
            times = (t - t.min()).dt.total_seconds().values
        else:
            times = [i * 0.3 for i in range(len(df))]
        rpm_vals = df[rpm_col].values.astype(float)
        if rpm_vals.max() < 500 and rpm_vals.max() > 0:
            rpm_vals = rpm_vals * 100
        return rpm_active_seconds_from_array(rpm_vals, times)
    except Exception:
        return None

# ─── FILE PARSERS ───────────────────────────────────────────────────────────────

def get_or_create_component(data, brand, serial, comp_type):
    name = f"{brand}_{serial}"
    for c in data["components"]:
        if c["name"].lower() == name.lower() and c["type"] == comp_type:
            return c["id"]
    new_id = f"{comp_type}_{brand}_{serial}".lower().replace(" ", "_")
    if new_id not in {c["id"] for c in data["components"]}:
        data["components"].append({"id": new_id, "type": comp_type, "name": name, "limit": None})
    return new_id

def parse_filename_v6(fname, data):
    stem, parts = Path(fname).stem, Path(fname).stem.replace("-", "_").split("_")
    brands     = data.get("brands", DEFAULT_BRANDS)
    date_str, start_idx, comp_ids = None, 0, []
    for i in range(len(parts) - 2):
        try:
            dt = datetime.strptime(f"{parts[i]}_{parts[i+1]}_{parts[i+2]}", "%Y_%m_%d")
            date_str, start_idx = dt.strftime("%Y-%m-%d"), i + 3
            break
        except ValueError:
            continue
    i = start_idx
    while i < len(parts) - 1:
        comp_type = brands.get(parts[i].lower())
        if comp_type:
            cid = get_or_create_component(data, parts[i], parts[i+1], comp_type)
            if cid not in comp_ids:
                comp_ids.append(cid)
            i += 2
        else:
            i += 1
    return date_str, comp_ids

def parse_edvm(filepath):
    try:
        import scipy.io as sio
        mat      = sio.loadmat(filepath)
        rpm_secs = rpm_duration_from_edvm(filepath)
        if rpm_secs and rpm_secs > 0:
            duration_sec = rpm_secs
        else:
            try:
                time_data    = mat["Signal01"][0][0][0].flatten()
                duration_sec = float(time_data[-1] - time_data[0])
            except Exception:
                duration_sec = 0.0
            if duration_sec <= 0:
                try:
                    duration_sec = float(mat["Metadata"][0][0][0][0][0][0][3][0][0])
                except Exception:
                    duration_sec = 0.0
        ts = datetime.now().isoformat()
        try:
            ts = datetime.strptime(str(mat["Metadata"][0][0][0][0][0][0][6][0]).strip()[:19],
                                   "%Y-%m-%d %H:%M:%S").isoformat()
        except Exception:
            pass
        return {"hours": duration_sec / 3600, "timestamp": ts, "ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}

def parse_csv(filepath):
    try:
        rpm_secs = rpm_duration_from_csv(filepath)
        sep, enc = ";", "latin-1"
        for e in ["latin-1", "utf-8", "cp1252"]:
            try:
                df = pd.read_csv(filepath, sep=";", encoding=e, nrows=3)
                if len(df.columns) > 1:
                    enc, sep = e, ";"; break
                df = pd.read_csv(filepath, sep=",", encoding=e, nrows=3)
                if len(df.columns) > 1:
                    enc, sep = e, ","; break
            except Exception:
                continue
        df = pd.read_csv(filepath, sep=sep, encoding=enc)
        if rpm_secs and rpm_secs > 0:
            duration_sec = rpm_secs
        else:
            esc_col  = next((c for c in df.columns if "ESC time" in c), None)
            time_col = next((c for c in df.columns if c.strip().lower() in ["time","time (sec)"]), None)
            if esc_col:
                duration_sec = float(df[esc_col].max() - df[esc_col].min())
            elif time_col:
                t = pd.to_datetime(df[time_col], format="%H:%M:%S.%f", errors="coerce")
                duration_sec = (t.max()-t.min()).total_seconds() if t.notna().sum()>1 else len(df)*0.3
            else:
                duration_sec = len(df) * 0.3
        fname = Path(filepath).stem
        ts    = None
        parts = fname.replace("-","_").split("_")
        if len(parts) >= 3:
            try:
                ts = datetime.strptime("_".join(parts[:3]), "%Y_%m_%d").isoformat()
            except Exception:
                pass
        if not ts:
            ts = datetime.fromtimestamp(os.path.getmtime(filepath)).isoformat()
        return {"hours": duration_sec / 3600, "timestamp": ts, "ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}

def parse_wnq(filepath):
    try:
        with open(filepath, "rb") as f:
            raw = f.read(8)
            if len(raw) < 8:
                raise ValueError("File too small")
            dow = struct.unpack_from("<H", raw, 0)[0]
            nch = struct.unpack_from("<H", raw, 2)[0]
            srd = struct.unpack_from("<H", raw, 4)[0]
            msr = struct.unpack_from("<H", raw, 6)[0]
            if nch < 1 or nch > 256:
                raise ValueError(f"Bad channels: {nch}")
            sr = (msr/srd/nch) if srd>0 and msr>0 else 240.0/nch
            f.seek(0,2); fs = f.tell()
            ds = fs - dow*2
            if dow*2 >= fs:
                raise ValueError("Bad offset")
            dur = (ds//2)//nch / sr if sr>0 else 0.0
        return {"hours": dur/3600,
                "timestamp": datetime.fromtimestamp(os.path.getmtime(filepath)).isoformat(),
                "ok": True}
    except Exception as e:
        return {"ok": False, "error": str(e)}

PARSERS = {".edvm": parse_edvm, ".csv": parse_csv, ".wnq": parse_wnq}

def parse_file(fp):
    ext = Path(fp).suffix.lower()
    return PARSERS[ext](fp) if ext in PARSERS else {"ok": False, "error": f"Unsupported: {ext}"}

def build_session(stem, full_path, result, fname, data):
    date_str, comp_ids = parse_filename_v6(fname, data)
    ts = result["timestamp"]
    if date_str:
        ts = date_str + "T00:00:00"
    return {"id": f"import_{stem}_{datetime.now().timestamp():.0f}",
            "file_name": stem, "full_path": full_path, "timestamp": ts,
            "hours": round(result["hours"], 6), "components": comp_ids, "notes": ""}

def import_local_files(data, watch_folder=None):
    if watch_folder is None:
        watch_folder = os.path.join(_BASE_DIR, "test_data")
    wp = Path(watch_folder)
    if not wp.exists():
        return 0, []
    imp  = set(data.get("imported_files", []))
    mnam = {s.get("file_name","") for s in data["sessions"]}
    cnt, errs = 0, []
    for fp in sorted(wp.iterdir()):
        if fp.suffix.lower() not in PARSERS:
            continue
        if fp.name in imp or fp.stem in mnam:
            if fp.stem in mnam:
                data["imported_files"].append(fp.name)
            continue
        r = parse_file(str(fp))
        if not r["ok"]:
            errs.append(f"{fp.name}: {r['error']}"); continue
        data["sessions"].append(build_session(fp.stem, fp.name, r, fp.name, data))
        data["imported_files"].append(fp.name)
        cnt += 1
    save_data(data)
    return cnt, errs

# ─── GOOGLE DRIVE ──────────────────────────────────────────────────────────────

def get_drive_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    return build("drive","v3", credentials=service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE, scopes=["https://www.googleapis.com/auth/drive.readonly"]))

def list_drive_files_recursive(service, folder_id, path=""):
    results, pt = [], None
    while True:
        r = service.files().list(q=f"'{folder_id}' in parents and trashed=false",
            fields="nextPageToken, files(id,name,mimeType,modifiedTime)",
            pageToken=pt, pageSize=200).execute()
        for item in r.get("files",[]):
            ip = f"{path}/{item['name']}" if path else item["name"]
            if item["mimeType"]=="application/vnd.google-apps.folder":
                results.extend(list_drive_files_recursive(service,item["id"],ip))
            elif Path(item["name"]).suffix.lower() in SUPPORTED_EXTS:
                results.append({"id":item["id"],"name":item["name"],"full_path":ip,
                                 "modifiedTime":item.get("modifiedTime","")})
        pt = r.get("nextPageToken")
        if not pt: break
    return results

def download_drive_file(service, file_id, dest):
    from googleapiclient.http import MediaIoBaseDownload
    import io
    req = service.files().get_media(fileId=file_id)
    fh  = io.FileIO(dest,"wb")
    dl  = MediaIoBaseDownload(fh, req)
    done = False
    while not done: _, done = dl.next_chunk()
    fh.close()

def sync_from_drive(data):
    if not os.path.exists(CREDENTIALS_FILE):
        return 0,0,["credentials.json not found"]
    try: service = get_drive_service()
    except Exception as e: return 0,0,[f"Drive error: {e}"]
    files = list_drive_files_recursive(service, DRIVE_FOLDER_ID)
    imp   = set(data.get("imported_files",[]))
    mnam  = {s.get("file_name","") for s in data["sessions"]}
    nc = sk = 0; errs = []
    with tempfile.TemporaryDirectory() as td:
        for f in files:
            fn,st = f["name"],Path(f["name"]).stem
            if fn in imp: sk+=1; continue
            if st in mnam: data["imported_files"].append(fn); sk+=1; continue
            tp = os.path.join(td, fn)
            try: download_drive_file(service, f["id"], tp)
            except Exception as e: errs.append(f"{fn}: {e}"); continue
            r = parse_file(tp)
            if not r["ok"]: errs.append(f"{fn}: {r['error']}"); data["imported_files"].append(fn); continue
            data["sessions"].append(build_session(st, f["full_path"], r, fn, data))
            data["imported_files"].append(fn); nc+=1
    data["last_sync"] = datetime.now().isoformat()
    save_data(data)
    return nc, sk, errs

# ─── ASANA ─────────────────────────────────────────────────────────────────────

def fetch_asana_sections(token, project_gid):
    """Fetch all sections in an Asana project."""
    try:
        import urllib.request
        url = f"https://app.asana.com/api/1.0/projects/{project_gid}/sections?opt_fields=name,gid"
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            return json.loads(resp.read())["data"]
    except Exception:
        return []

def fetch_asana_tasks(token, project_gid, section_gid=""):
    """Fetch tasks from an Asana project, optionally filtered to a section."""
    try:
        import urllib.request
        if section_gid:
            url = (f"https://app.asana.com/api/1.0/tasks"
                   f"?section={section_gid}"
                   f"&opt_fields=name,due_on,completed,completed_at,notes,memberships.section.name"
                   f"&limit=100")
        else:
            url = (f"https://app.asana.com/api/1.0/tasks"
                   f"?project={project_gid}"
                   f"&opt_fields=name,due_on,completed,completed_at,notes,memberships.section.name"
                   f"&limit=100")
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            return json.loads(resp.read())["data"]
    except Exception:
        return []

# ─── PAGE SETUP ────────────────────────────────────────────────────────────────

st.set_page_config(page_title="AIR — Test Bench", layout="wide")

st.markdown("""
<style>
html,body,[data-testid="stAppViewContainer"],[data-testid="stApp"]{
    background-color:#f9f9f7!important;
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Helvetica,Arial,sans-serif!important;
}
[data-testid="stHeader"]{background-color:#f9f9f7!important;}
/* sidebar nav */
.stButton>button{border-radius:0!important;font-weight:400;font-size:13px;letter-spacing:.04em;
    border:1px solid #222!important;background:transparent!important;color:#111!important;
    transition:all .15s;text-decoration:none!important;}
.stButton>button:hover{background:#111!important;color:#fff!important;}
.stButton>button[kind="primary"]{background:#111!important;color:#fff!important;border:1px solid #111!important;}
.stButton>button[kind="primary"]:hover{background:#333!important;}
.stButton>button span{text-decoration:none!important;border-bottom:none!important;}
input,select,textarea{border-radius:0!important;border:1px solid #ddd!important;}
h1,h2,h3,h4,h5{font-weight:500!important;letter-spacing:-.01em;color:#111!important;}
.badge-ok     {background:#e8f5e2;color:#2d6a1f;padding:2px 10px;border-radius:2px;font-size:11px;font-weight:500;letter-spacing:.04em;}
.badge-warning{background:#fff3e0;color:#7a4f00;padding:2px 10px;border-radius:2px;font-size:11px;font-weight:500;letter-spacing:.04em;}
.badge-danger {background:#fdecea;color:#9b2020;padding:2px 10px;border-radius:2px;font-size:11px;font-weight:500;letter-spacing:.04em;}
.kpi-card{background:#fff;border:1px solid #e0e0e0;border-radius:2px;padding:18px 20px;text-align:center;}
.kpi-value{font-size:24px;font-weight:600;color:#111;letter-spacing:-.02em;}
.kpi-label{font-size:11px;color:#888;margin-top:4px;letter-spacing:.06em;text-transform:uppercase;}
.ev-info    {border-left:3px solid #378ADD;padding:6px 10px;background:#E6F1FB;border-radius:0 4px 4px 0;margin:4px 0;}
.ev-warning {border-left:3px solid #EF9F27;padding:6px 10px;background:#FAEEDA;border-radius:0 4px 4px 0;margin:4px 0;}
.ev-critical{border-left:3px solid #E24B4A;padding:6px 10px;background:#FCEBEB;border-radius:0 4px 4px 0;margin:4px 0;}
hr{border-color:#e8e8e6!important;}
[data-testid="stSidebar"]{background:#fafafa!important;border-right:1px solid #e0e0e0!important;}
[data-testid="stSidebar"] .stButton>button{
    border:none!important;background:transparent!important;color:#333!important;
    text-align:left!important;padding:8px 12px!important;border-radius:2px!important;
    font-size:13px!important;font-weight:400!important;letter-spacing:.01em!important;
    width:100%!important;justify-content:flex-start!important;
}
[data-testid="stSidebar"] .stButton>button:hover{background:#f0f0f0!important;color:#111!important;}
</style>
""", unsafe_allow_html=True)

# ── Session state ──────────────────────────────────────────────────────────────
for k, v in [
    ("data",                  None),
    ("timer_running",         False),
    ("timer_start",           None),
    ("timer_elapsed",         0.0),
    ("drill_down_mfr",        None),
    ("drill_down_type",       None),
    ("confirm_delete_id",     None),
    ("confirm_delete_comp",   None),
    ("editing_comp_id",       None),
    ("goto_components",       False),
]:
    if k not in st.session_state:
        st.session_state[k] = v

if st.session_state.data is None:
    st.session_state.data = load_data()

data = st.session_state.data

# ── Sidebar navigation ────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="padding:18px 0 20px;">
      <svg width="60" height="24" viewBox="0 0 160 50" xmlns="http://www.w3.org/2000/svg">
        <text x="4" y="38" font-family="-apple-system,BlinkMacSystemFont,Segoe UI,Helvetica,Arial,sans-serif"
          font-size="42" font-weight="600" letter-spacing="-1" fill="#111">air</text>
      </svg>
      <div style="font-size:10px;color:#999;letter-spacing:.14em;text-transform:uppercase;margin-top:2px;">Motor Test Bench</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<hr style='margin:0 0 12px;border-color:#e0e0e0;'>", unsafe_allow_html=True)

    NAV_PAGES = [
        "Overview",
        "Dashboard",
        "Log session",
        "Log book",
        "Components",
        "Rig configuration",
        "Log events",
        "Settings",
    ]
    if "nav_page" not in st.session_state:
        st.session_state.nav_page = "Overview"

    for page in NAV_PAGES:
        is_active = st.session_state.nav_page == page
        style = ("background:#111;color:#fff;" if is_active
                 else "background:transparent;color:#333;")
        if st.button(page, key=f"nav_{page}", use_container_width=True):
            st.session_state.nav_page = page
            st.rerun()

active_page = st.session_state.nav_page

# ── Page header ───────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="display:flex;align-items:center;gap:20px;padding:8px 0 14px;
            border-bottom:1px solid #222;margin-bottom:1.5rem;">
  <span style="font-size:20px;font-weight:500;color:#111;letter-spacing:-.01em;">{active_page}</span>
</div>
""", unsafe_allow_html=True)

# ── Helper: only render a section if it's the active page ─────────────────────
def page(name):
    return active_page == name

# ══════════════════════════════════════════════════════════════════════════════
# ── OVERVIEW ─────────────────────────────────────────────────────────────────
if page("Overview"):
    now        = datetime.now()
    t_start, t_end = today_range()
    w_start, w_end = week_range()
    m_start, m_end = month_range()

    daily_sessions   = sessions_in_range(data["sessions"], t_start, t_end)
    weekly_sessions  = sessions_in_range(data["sessions"], w_start, w_end)
    monthly_sessions = sessions_in_range(data["sessions"], m_start, m_end)
    all_sessions     = data["sessions"]

    daily_h   = sum(s["hours"] for s in daily_sessions)
    weekly_h  = sum(s["hours"] for s in weekly_sessions)
    monthly_h = sum(s["hours"] for s in monthly_sessions)
    total_h   = sum(s["hours"] for s in all_sessions)

    alert_comps = [c for c in data["components"]
                   if c.get("limit") and get_comp_hours(c["id"], data["sessions"]) >= c["limit"] * 0.75]

    # ── Row 1: Time KPI cards ─────────────────────────────────────────────────
    k1, k2, k3, k4, k5 = st.columns(5)
    for col, label, val in [
        (k1, "TODAY",      fmt_dur(daily_h)),
        (k2, "THIS WEEK",  fmt_dur(weekly_h)),
        (k3, "THIS MONTH", fmt_dur(monthly_h)),
        (k4, "OVERALL",    fmt_dur(total_h)),
        (k5, "ALERTS",     str(len(alert_comps)) + (" ⚠️" if alert_comps else " ✓")),
    ]:
        bg  = "#FCEBEB" if label == "ALERTS" and alert_comps else "#fff"
        clr = "#9b2020" if label == "ALERTS" and alert_comps else "#111"
        col.markdown(f"""<div class="kpi-card" style="background:{bg};">
          <div class="kpi-value" style="color:{clr};">{val}</div>
          <div class="kpi-label">{label}</div>
        </div>""", unsafe_allow_html=True)

    # ── Row 2: counts + today's actual run notes ─────────────────────────────
    r1, r2, r3, r4, r5 = st.columns(5)

    # Today — list each session's note text (or "Manual run" if no note)
    if daily_sessions:
        lines_html = ""
        for s in daily_sessions:
            note = (s.get("notes") or "").strip()
            dur  = fmt_dur(s["hours"])
            label = note if note else "Manual run"
            lines_html += (f"<div style='font-size:11px;color:#555;margin:1px 0;'>"
                           f"<b>{dur}</b> — {label}</div>")
    else:
        lines_html = "<div style='font-size:11px;color:#aaa;'>No runs today</div>"
    r1.markdown(f"""<div style="background:#fff;border:1px solid #ebebeb;padding:10px 14px;min-height:60px;">
      <div style="font-size:11px;font-weight:600;color:#888;margin-bottom:4px;">TODAY'S RUNS</div>
      {lines_html}
    </div>""", unsafe_allow_html=True)

    # Week / Month / Overall — count
    for col, sess, lbl in [
        (r2, weekly_sessions,  "THIS WEEK"),
        (r3, monthly_sessions, "THIS MONTH"),
        (r4, all_sessions,     "OVERALL"),
    ]:
        col.markdown(f"""<div style="background:#fff;border:1px solid #ebebeb;padding:10px 14px;min-height:60px;">
          <div style="font-size:11px;font-weight:600;color:#888;margin-bottom:4px;">RUNS</div>
          <div style="font-size:22px;font-weight:700;color:#111;">{len(sess)}</div>
        </div>""", unsafe_allow_html=True)

    # Alerts
    if alert_comps:
        if r5.button("View alerts →", key="ov_alert_btn"):
            st.session_state.goto_components = True
    else:
        r5.markdown("""<div style="background:#fff;border:1px solid #ebebeb;padding:10px 14px;min-height:60px;">
          <div style="font-size:11px;font-weight:600;color:#888;margin-bottom:4px;">STATUS</div>
          <div style="font-size:13px;color:#639922;font-weight:600;">All clear ✓</div>
        </div>""", unsafe_allow_html=True)

    # ── Component lifetime KPIs ───────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    motor_h = sum(get_comp_hours(c["id"], all_sessions)
                  for c in data["components"] if c["type"] == "Motor")
    esc_h   = sum(get_comp_hours(c["id"], all_sessions)
                  for c in data["components"] if c["type"] == "ESC")
    prop_h  = sum(get_comp_hours(c["id"], all_sessions)
                  for c in data["components"] if c["type"] == "Prop")
    lk1, lk2, lk3 = st.columns(3)
    for col, lbl, val in [(lk1, "ALL MOTORS", motor_h),
                          (lk2, "ALL ESCs",   esc_h),
                          (lk3, "ALL PROPS",  prop_h)]:
        col.markdown(f"""<div class="kpi-card">
          <div class="kpi-value">{fmt_dur(val)}</div>
          <div class="kpi-label">{lbl} — TOTAL HOURS</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    left_col, right_col = st.columns([3, 2])

    # ── Left: Line chart by month ─────────────────────────────────────────────
    with left_col:
        st.markdown("#### Run time over time")

        import plotly.graph_objects as go
        from collections import defaultdict

        def hours_to_hhmm(h):
            total_sec = int(round(h * 3600))
            hh = total_sec // 3600
            mm = (total_sec % 3600) // 60
            return f"{hh:02d}:{mm:02d}"

        # Aggregate by month (YYYY-MM)
        month_totals = defaultdict(float)
        for s in all_sessions:
            ts = s.get("timestamp", "")[:7]  # "YYYY-MM"
            if ts and len(ts) == 7:
                month_totals[ts] += s["hours"]

        if month_totals:
            sorted_months = sorted(month_totals.keys())
            # Pretty x labels: "Apr 2026"
            def month_label(ym):
                try:
                    dt = datetime.strptime(ym, "%Y-%m")
                    return dt.strftime("%b %Y")
                except:
                    return ym
            x_labels = [month_label(m) for m in sorted_months]
            y_vals   = [month_totals[m] for m in sorted_months]

            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=x_labels,
                y=y_vals,
                mode="lines+markers",
                line=dict(color="#111111", width=2.5),
                marker=dict(size=6, color="#111111"),
                hovertemplate="%{x}<br><b>%{customdata}</b><extra></extra>",
                customdata=[hours_to_hhmm(h) for h in y_vals],
                fill="tozeroy",
                fillcolor="rgba(17,17,17,0.05)",
            ))
            # Y-axis ticks: show clean HH:MM labels
            max_h = max(y_vals) if y_vals else 1
            step  = max(max_h / 5, 1/60)
            tick_vals = [round(i * step, 4) for i in range(6)]
            tick_text = [hours_to_hhmm(v) for v in tick_vals]

            fig.update_layout(
                height=280,
                margin=dict(l=10, r=10, t=10, b=10),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="#f9f9f7",
                xaxis=dict(
                    showgrid=False,
                    tickfont=dict(size=11),
                    tickangle=0,
                    fixedrange=True,
                ),
                yaxis=dict(
                    showgrid=True,
                    gridcolor="#e8e8e6",
                    tickfont=dict(size=11),
                    tickvals=tick_vals,
                    ticktext=tick_text,
                    title=None,
                    fixedrange=True,
                    rangemode="tozero",
                ),
                showlegend=False,
            )
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
        else:
            st.info("No session data to chart yet.")

    # ── Right: Rig config + Event log ─────────────────────────────────────────
    with right_col:
        # Rig config — dual A/B
        active = data.get("active_side","Side A")
        active_color = "#e85d00" if active=="Both" else "#111"
        st.markdown(f"""
        <div style="margin-bottom:8px;">
          <span style="font-size:15px;font-weight:600;">Current rig configuration</span>
          &nbsp;&nbsp;
          <span style="background:#111;color:#fff;padding:2px 10px;border-radius:2px;
                       font-size:12px;font-weight:500;">{active} active</span>
        </div>""", unsafe_allow_html=True)

        cfg_a = data.get("rig_side_a",{})
        cfg_b = data.get("rig_side_b",{})

        def render_cfg_display(cfg):
            display_fields = [
                ("Motor",   cfg.get("Motor",""),   None),
                ("ESC",   cfg.get("ESC",""),     cfg.get("ESC_role","")),
                ("ESC",   cfg.get("ESC2",""),    cfg.get("ESC2_role","")),
                ("Prop",    cfg.get("Prop",""),    None),
                ("Prop",    cfg.get("Prop2",""),   None),
                ("Battery", cfg.get("Battery",""), None),
            ]
            for label, val, role in display_fields:
                if not val:
                    continue
                role_tag = (f" <span style='font-size:10px;background:#eee;padding:1px 5px;"
                            f"border-radius:2px;'>{role}</span>") if role else ""
                st.markdown(
                    f"<div style='font-size:12px;margin:2px 0;'><b>{label}:</b> {val}{role_tag}</div>",
                    unsafe_allow_html=True)
            for f in ["Test objective","Operator","Notes"]:
                v = cfg.get(f,"")
                if v:
                    st.markdown(
                        f"<div style='font-size:12px;margin:2px 0;color:#666;'><b>{f}:</b> {v}</div>",
                        unsafe_allow_html=True)

        if not cfg_a and not cfg_b:
            st.caption("No configuration set. Add it in the Input tab.")
        else:
            ra, rb = st.columns(2)
            with ra:
                st.markdown("**Side A**")
                render_cfg_display(cfg_a)
            with rb:
                st.markdown("**Side B**")
                render_cfg_display(cfg_b)

        st.markdown("<br>", unsafe_allow_html=True)

        # Event log
        st.markdown("#### Event log")
        events = sorted(data.get("events", []),
                        key=lambda x: x.get("timestamp",""), reverse=True)
        if not events:
            st.caption("No events logged. Add them in the Input tab.")
        else:
            for ev in events[:8]:
                sev   = ev.get("severity", "info")
                ts    = ev.get("timestamp","")[:16].replace("T"," ")
                txt   = ev.get("text","")
                st.markdown(f"""
                <div class="ev-{sev}">
                  <span style="font-size:11px;color:#888;">{ts}</span>
                  <span style="font-size:11px;text-transform:uppercase;font-weight:600;
                               margin:0 6px;color:{'#378ADD' if sev=='info' else '#EF9F27' if sev=='warning' else '#E24B4A'};">{sev}</span>
                  <span style="font-size:13px;">{txt}</span>
                </div>""", unsafe_allow_html=True)

    # ── Asana comparison ──────────────────────────────────────────────────────
    asana_cfg = data.get("asana", {})
    if asana_cfg.get("token") and asana_cfg.get("project_gid"):
        st.divider()
        st.markdown("#### Planned vs actual (Asana)")
        # Load sections for filter dropdown
        if "asana_sections" not in st.session_state:
            st.session_state.asana_sections = []
        col_asana_sec, col_asana_btn = st.columns([3,1])
        with col_asana_btn:
            if st.button("Load sections", key="load_sections"):
                with st.spinner("Loading sections…"):
                    st.session_state.asana_sections = fetch_asana_sections(
                        asana_cfg["token"], asana_cfg["project_gid"])
        sections     = st.session_state.asana_sections
        sec_names    = ["All sections"] + [s["name"] for s in sections]
        saved_filter = asana_cfg.get("section_filter","")
        saved_idx    = next((i+1 for i,s in enumerate(sections) if s["name"]==saved_filter), 0)
        with col_asana_sec:
            sel_sec = st.selectbox("Filter by section", sec_names,
                                   index=saved_idx, key="asana_sec_sel")
        sel_sec_gid = ""
        if sel_sec != "All sections":
            sel_sec_gid = next((s["gid"] for s in sections if s["name"]==sel_sec), "")
            # Save preference
            if sel_sec != saved_filter:
                data["asana"]["section_filter"] = sel_sec
                save_data(data); st.session_state.data = data

        with st.spinner("Loading Asana tasks…"):
            tasks = fetch_asana_tasks(asana_cfg["token"], asana_cfg["project_gid"], sel_sec_gid)
        if tasks:
            rows = []
            for t in tasks:
                due     = t.get("due_on") or "—"
                done    = t.get("completed", False)
                done_at = (t.get("completed_at") or "")[:10] or "—"
                today   = now.strftime("%Y-%m-%d")
                if done:
                    status = "Done"
                elif due != "—" and due < today:
                    status = "Overdue"
                else:
                    status = "Planned"
                section = ""
                for m in t.get("memberships", []):
                    sec = m.get("section")
                    if sec:
                        section = sec.get("name","")
                        break
                rows.append({
                    "Task":      t["name"],
                    "Section":   section,
                    "Due":       due,
                    "Status":    status,
                    "Completed": done_at,
                })
            if rows:
                df_asana = pd.DataFrame(rows)
                # Color-code status
                st.dataframe(df_asana, use_container_width=True, hide_index=True)
                done_count    = sum(1 for r in rows if r["Status"]=="Done")
                overdue_count = sum(1 for r in rows if r["Status"]=="Overdue")
                planned_count = sum(1 for r in rows if r["Status"]=="Planned")
                a1,a2,a3 = st.columns(3)
                a1.metric("Done",    done_count)
                a2.metric("Planned", planned_count)
                a3.metric("Overdue", overdue_count,
                          delta=f"-{overdue_count}" if overdue_count else None,
                          delta_color="inverse")
        else:
            st.info("No tasks found — check your Asana token and project GID in Settings.")

# ══════════════════════════════════════════════════════════════════════════════
# ── DASHBOARD ────────────────────────────────────────────────────────────────
if page("Dashboard"):
    import plotly.graph_objects as go
    from collections import defaultdict

    all_sessions  = data["sessions"]
    all_comps     = data["components"]

    # ── Helper: get brand from component name (first token) ───────────────────
    def brand_of(name):
        return name.split("_")[0] if "_" in name else name

    # ── Helper: build brand→hours for a given component type ─────────────────
    def brand_hours(comp_type):
        bh = defaultdict(float)
        for s in all_sessions:
            for cid in s.get("components", []):
                c = next((x for x in all_comps if x["id"] == cid), None)
                if c and c["type"] == comp_type:
                    bh[brand_of(c["name"])] += s["hours"]
        return dict(bh)

    # ── Row 1: Basic KPIs ─────────────────────────────────────────────────────
    t_start, t_end = today_range()
    w_start, w_end = week_range()
    daily_h  = hours_in_range(all_sessions, t_start, t_end)
    weekly_h = hours_in_range(all_sessions, w_start, w_end)
    total_h  = sum(s["hours"] for s in all_sessions)
    alert_comps = [c for c in all_comps
                   if c.get("limit") and get_comp_hours(c["id"], all_sessions) >= c["limit"]*0.75]

    k1, k2, k3, k4 = st.columns(4)
    for col, label, val in [
        (k1, "Total run time", fmt_dur(total_h)),
        (k2, "Today",          fmt_dur(daily_h)),
        (k3, "This week",      fmt_dur(weekly_h)),
        (k4, "Alerts",         str(len(alert_comps))+(" ⚠️" if alert_comps else " ✓")),
    ]:
        bg  = "#FCEBEB" if label=="Alerts" and alert_comps else "#fff"
        clr = "#9b2020" if label=="Alerts" and alert_comps else "#111"
        col.markdown(f"""<div class="kpi-card" style="background:{bg};">
          <div class="kpi-value" style="color:{clr};">{val}</div>
          <div class="kpi-label">{label}</div>
        </div>""", unsafe_allow_html=True)

    if alert_comps:
        if st.button("View component alerts", key="db_alert_btn"):
            st.session_state.goto_components = True

    st.markdown("<br>", unsafe_allow_html=True)

    # Alert panel
    if st.session_state.goto_components:
        st.subheader("Components needing attention")
        if st.button("Close", key="close_comps"):
            st.session_state.goto_components = False; st.rerun()
        for comp in all_comps:
            hours = get_comp_hours(comp["id"], all_sessions)
            limit = comp.get("limit")
            s_key, s_label = comp_status(hours, limit)
            if s_key == "ok": continue
            pct = min((hours/limit) if limit else 0, 1.0)
            bar_color = "#EF9F27" if s_key=="warning" else "#E24B4A"
            st.markdown(f"""
            <div style="background:#fff8f8;border:1px solid #f5c6c6;border-radius:4px;
                        padding:10px 14px;margin-bottom:6px;">
              <b>{comp['name']}</b> <span style="font-size:12px;color:#555;">{comp['type']}</span>
              &nbsp;<span class="badge-{s_key}">{s_label}</span>
              <div style="background:#eee;border-radius:2px;height:5px;margin-top:6px;">
                <div style="width:{pct*100:.0f}%;background:{bar_color};height:5px;border-radius:2px;"></div>
              </div>
              <small style="color:#888;">{fmt_dur(hours)} / {fmt_dur(limit) if limit else "—"}</small>
            </div>""", unsafe_allow_html=True)
        st.divider()

    CHART_COLORS = ["#111","#444","#777","#aaa","#ccc","#e0e0e0"]

    def small_donut(title, bh_dict, key):
        if not bh_dict:
            st.caption(f"No data — {title}")
            return
        labels = list(bh_dict.keys())
        values = [bh_dict[l] for l in labels]
        hover  = [fmt_dur(v) for v in values]
        fig = go.Figure(go.Pie(
            labels=labels, values=values,
            hole=0.55,
            marker_colors=CHART_COLORS[:len(labels)],
            hovertemplate="%{label}<br>%{customdata}<extra></extra>",
            customdata=hover,
            textinfo="percent",
            textfont_size=11,
        ))
        fig.update_layout(
            title=dict(text=title, font_size=13, x=0),
            height=220, margin=dict(t=36,b=0,l=0,r=0),
            paper_bgcolor="rgba(0,0,0,0)",
            legend=dict(font_size=10, orientation="h", y=-0.1),
            showlegend=True,
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False}, key=key)

    # ── Row A: Brand share donuts ─────────────────────────────────────────────
    st.subheader("Brand share by component")
    da, db, dc = st.columns(3)
    with da: small_donut("Motors",     brand_hours("Motor"), "donut_motor")
    with db: small_donut("ESCs",       brand_hours("ESC"),   "donut_esc")
    with dc: small_donut("Propellers", brand_hours("Prop"),  "donut_prop")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Row B: Top 5 setups + Rig utilisation ────────────────────────────────
    col_top5, col_rig = st.columns([3, 2])

    with col_top5:
        st.subheader("Top 5 setups")
        setup_hours = defaultdict(float)
        for s in all_sessions:
            motor_brand = esc_brand = prop_brand = "—"
            for cid in s.get("components", []):
                c = next((x for x in all_comps if x["id"] == cid), None)
                if not c: continue
                b = brand_of(c["name"])
                if c["type"] == "Motor": motor_brand = b
                elif c["type"] == "ESC": esc_brand   = b
                elif c["type"] == "Prop": prop_brand = b
            key = f"{motor_brand} / {esc_brand} / {prop_brand}"
            setup_hours[key] += s["hours"]

        if setup_hours:
            top5 = sorted(setup_hours.items(), key=lambda x: -x[1])[:5]
            labels = [t[0] for t in top5]
            values = [t[1] for t in top5]
            fig_top5 = go.Figure(go.Bar(
                y=labels, x=values,
                orientation="h",
                marker_color=CHART_COLORS[:len(labels)],
                hovertemplate="%{y}<br>%{customdata}<extra></extra>",
                customdata=[fmt_dur(v) for v in values],
            ))
            fig_top5.update_layout(
                height=240, margin=dict(t=0,b=10,l=0,r=10),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="#f9f9f7",
                xaxis=dict(showgrid=True, gridcolor="#e8e8e6", title=None, tickfont_size=10),
                yaxis=dict(showgrid=False, tickfont_size=10, autorange="reversed"),
            )
            st.plotly_chart(fig_top5, use_container_width=True,
                            config={"displayModeBar":False}, key="top5_chart")
        else:
            st.caption("No session data yet.")

    with col_rig:
        st.subheader("Rig utilisation")
        side_a_h = sum(s["hours"] for s in all_sessions if s.get("side") in ("Side A","Both"))
        side_b_h = sum(s["hours"] for s in all_sessions if s.get("side") in ("Side B","Both"))
        unattrib = sum(s["hours"] for s in all_sessions if not s.get("side"))

        rig_labels = []; rig_vals = []
        if side_a_h: rig_labels.append("Side A"); rig_vals.append(side_a_h)
        if side_b_h: rig_labels.append("Side B"); rig_vals.append(side_b_h)
        if unattrib: rig_labels.append("Unattributed"); rig_vals.append(unattrib)

        if rig_vals:
            fig_rig = go.Figure(go.Pie(
                labels=rig_labels, values=rig_vals,
                hole=0.5,
                marker_colors=["#111","#777","#ccc"],
                hovertemplate="%{label}<br>%{customdata}<extra></extra>",
                customdata=[fmt_dur(v) for v in rig_vals],
                textinfo="percent+label", textfont_size=11,
            ))
            fig_rig.update_layout(
                height=240, margin=dict(t=0,b=0,l=0,r=0),
                paper_bgcolor="rgba(0,0,0,0)",
                showlegend=False,
            )
            st.plotly_chart(fig_rig, use_container_width=True,
                            config={"displayModeBar":False}, key="rig_chart")
            # numeric labels below
            for lbl, val in zip(rig_labels, rig_vals):
                st.markdown(f"<div style='font-size:12px;color:#555;'><b>{lbl}:</b> {fmt_dur(val)}</div>",
                            unsafe_allow_html=True)
        else:
            st.caption("No side data logged yet. Use 'Run side' when logging sessions.")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Row C: Run outcomes + Test types ─────────────────────────────────────
    col_out, col_type = st.columns(2)

    with col_out:
        st.subheader("Run outcomes")
        outcome_hours = defaultdict(float)
        for s in all_sessions:
            outcome = (s.get("outcome") or "Success").strip() or "Success"
            outcome_hours[outcome] += s["hours"]
        if outcome_hours:
            out_labels = list(outcome_hours.keys())
            out_vals   = [outcome_hours[l] for l in out_labels]
            out_colors = {"Success":"#3B6D11","Hardware Failure":"#A32D2D",
                          "User Abort":"#854F0B","Inconclusive":"#777"}
            colors_out = [out_colors.get(l, "#aaa") for l in out_labels]
            fig_out = go.Figure(go.Pie(
                labels=out_labels, values=out_vals,
                hole=0.5,
                marker_colors=colors_out,
                hovertemplate="%{label}<br>%{customdata}<extra></extra>",
                customdata=[fmt_dur(v) for v in out_vals],
                textinfo="percent+label", textfont_size=11,
            ))
            fig_out.update_layout(
                height=240, margin=dict(t=0,b=0,l=0,r=0),
                paper_bgcolor="rgba(0,0,0,0)", showlegend=False,
            )
            st.plotly_chart(fig_out, use_container_width=True,
                            config={"displayModeBar":False}, key="outcome_chart")
        else:
            st.caption("No outcome data. Add outcome when logging sessions.")

    with col_type:
        st.subheader("Test types")
        type_hours = defaultdict(float)
        for s in all_sessions:
            ttype = (s.get("test_type") or "Unspecified").strip() or "Unspecified"
            type_hours[ttype] += s["hours"]
        if type_hours:
            tt_labels = list(type_hours.keys())
            tt_vals   = [type_hours[l] for l in tt_labels]
            fig_tt = go.Figure(go.Pie(
                labels=tt_labels, values=tt_vals,
                hole=0.5,
                marker_colors=CHART_COLORS[:len(tt_labels)],
                hovertemplate="%{label}<br>%{customdata}<extra></extra>",
                customdata=[fmt_dur(v) for v in tt_vals],
                textinfo="percent+label", textfont_size=11,
            ))
            fig_tt.update_layout(
                height=240, margin=dict(t=0,b=0,l=0,r=0),
                paper_bgcolor="rgba(0,0,0,0)", showlegend=False,
            )
            st.plotly_chart(fig_tt, use_container_width=True,
                            config={"displayModeBar":False}, key="testtype_chart")
        else:
            st.caption("No test type data. Select a type when logging sessions.")

# ══════════════════════════════════════════════════════════════════════════════
# ── LOG SESSION ──────────────────────────────────────────────────────────────
if page("Log session"):
    st.subheader("Log a session manually")
    with st.expander("+ Add a new component"):
        qc1,qc2,qc3 = st.columns([1,2,1])
        q_type  = qc1.selectbox("Type",["Motor","ESC","Prop","Battery"],key="q_type")
        q_name  = qc2.text_input("Name / ID",placeholder="e.g. Nidec_FX124",key="q_name")
        q_limit = qc3.number_input("Max hours",min_value=0.0,step=1.0,value=50.0,key="q_limit")
        if st.button("Add component",key="q_add"):
            if not q_name.strip():
                st.error("Please enter a name.")
            else:
                data["components"].append({"id":f"{q_type}_{datetime.now().timestamp():.0f}",
                    "type":q_type,"name":q_name.strip(),
                    "limit":q_limit if q_limit>0 else None})
                save_data(data); st.session_state.data=data
                st.session_state["added_msg"] = f"Added {q_type}: {q_name}"
                st.rerun()

    if "added_msg" in st.session_state:
        st.success(st.session_state.pop("added_msg"))

    # Side selector — smart component attribution
    st.markdown("#### Which side is running?")
    run_side = st.radio("Run side", ["Side A", "Side B", "Both", "Manual selection"],
                        horizontal=True, key="run_side_sel")

    # Show which components will be attributed based on rig config
    if run_side != "Manual selection":
        auto_ids   = get_side_component_ids(data, run_side)
        auto_names = [c["name"] for c in data["components"] if c["id"] in auto_ids]
        if auto_names:
            st.caption(f"Will log hours to: **{', '.join(auto_names)}** (from rig configuration)")
        else:
            st.caption("No components configured for this side yet. Set them up in the Input tab.")
        selected_ids = auto_ids
    else:
        comp_options   = {c["name"]:c["id"] for c in data["components"]}
        selected_names = st.multiselect("Components used",options=list(comp_options.keys()))
        selected_ids   = [comp_options[n] for n in selected_names]

    col_h, col_m = st.columns(2)
    hours_input   = col_h.number_input("Hours",   min_value=0, step=1, value=0)
    minutes_input = col_m.number_input("Minutes", min_value=0, max_value=59, step=1, value=0)

    note_col, type_col, outcome_col = st.columns(3)
    notes_input    = note_col.text_input("Notes (optional)", placeholder="e.g. ATP, endurance…")
    TEST_TYPES     = ["","Endurance","Thrust mapping","ATP","Vibration","Thermal","Acceptance","Other"]
    OUTCOMES       = ["Success","Hardware Failure","User Abort","Inconclusive","Other"]
    test_type_input = type_col.selectbox("Test type", TEST_TYPES, key="manual_test_type")
    outcome_input   = outcome_col.selectbox("Outcome", OUTCOMES, key="manual_outcome")

    file_name_input = st.text_input("Session name (optional)",
        placeholder="e.g. 2024-03-15_Nidec_FX124_MGM_12345_Helix_0001")

    if st.button("Log session", type="primary", key="log_manual"):
        total_h = hours_input + minutes_input/60
        if total_h <= 0:
            st.error("Please enter a duration greater than zero.")
        elif not selected_ids:
            st.error("No components selected or configured for this side.")
        else:
            side_label = run_side if run_side != "Manual selection" else "manual"
            data["sessions"].append({
                "id":        f"manual_{datetime.now().timestamp():.0f}",
                "file_name": file_name_input.strip(),
                "timestamp": datetime.now().isoformat(),
                "hours":     round(total_h, 6),
                "components": selected_ids,
                "side":      side_label,
                "notes":     notes_input.strip(),
                "test_type": test_type_input.strip(),
                "outcome":   outcome_input,
            })
            save_data(data); st.session_state.data = data
            st.success(f"Session logged: {fmt_dur(total_h)} → "
                       f"{', '.join(c['name'] for c in data['components'] if c['id'] in selected_ids)}")

    st.divider()
    st.subheader("Live timer")
    st.caption(f"Run time counted only when RPM > {RPM_THRESHOLD} (for imported files). Manual timer counts all time.")
    tc1,tc2,tc3 = st.columns(3)
    if tc1.button("Start",disabled=st.session_state.timer_running,key="btn_start"):
        st.session_state.timer_running=True; st.session_state.timer_start=datetime.now().timestamp(); st.rerun()
    if tc2.button("Pause",disabled=not st.session_state.timer_running,key="btn_pause"):
        st.session_state.timer_elapsed+=datetime.now().timestamp()-st.session_state.timer_start
        st.session_state.timer_running=False; st.session_state.timer_start=None; st.rerun()
    stop_dis = not st.session_state.timer_running and st.session_state.timer_elapsed==0
    if tc3.button("Stop & log",disabled=stop_dis,key="btn_stop"):
        if st.session_state.timer_running:
            st.session_state.timer_elapsed+=datetime.now().timestamp()-st.session_state.timer_start
        total_h = st.session_state.timer_elapsed/3600
        side_label = run_side if run_side != "Manual selection" else "manual"
        data["sessions"].append({
            "id":        f"timer_{datetime.now().timestamp():.0f}",
            "file_name": "",
            "timestamp": datetime.now().isoformat(),
            "hours":     round(total_h, 6),
            "components": selected_ids,
            "side":      side_label,
            "notes":     notes_input.strip(),
            "test_type": test_type_input.strip(),
            "outcome":   outcome_input,
        })
        save_data(data); st.session_state.data=data
        st.session_state.timer_running=False; st.session_state.timer_start=None
        st.session_state.timer_elapsed=0.0
        st.success(f"Session logged: {fmt_dur(total_h)}"); st.rerun()

    is_r = st.session_state.timer_running
    alr  = st.session_state.timer_elapsed
    st_s = st.session_state.timer_start or 0.0
    tel  = alr+(datetime.now().timestamp()-st_s) if is_r else alr
    hh,mm,ss = int(tel//3600),int((tel%3600)//60),int(tel%60)
    clr = "#e85d00" if is_r else ("#e07b00" if alr>0 else "#888")
    lbl = "Running…" if is_r else ("Paused" if alr>0 else "Ready — press Start")
    st.markdown(f"""
    <div style="font-family:monospace;font-size:52px;font-weight:600;
                color:{clr};letter-spacing:4px;padding:10px 0 2px;">{hh:02d}:{mm:02d}:{ss:02d}</div>
    <div style="font-size:13px;color:#888;margin-bottom:8px;">{lbl}</div>
    """, unsafe_allow_html=True)
    if is_r:
        time.sleep(1); st.rerun()

# ══════════════════════════════════════════════════════════════════════════════

# ── LOG BOOK ─────────────────────────────────────────────────────────────────
if page("Log book"):
    st.subheader("All sessions")
    if st.session_state.confirm_delete_id:
        ds = next((s for s in data["sessions"] if s["id"]==st.session_state.confirm_delete_id),None)
        if ds:
            cn = ", ".join(c["name"] for c in data["components"] if c["id"] in ds.get("components",[])) or "—"
            with st.container():
                st.markdown(f"""
                <div style="background:#fff3cd;border:1.5px solid #f0ad4e;border-radius:8px;
                            padding:16px 20px;margin:8px 0 16px;max-width:600px;">
                  <div style="font-size:16px;font-weight:600;margin-bottom:8px;">Delete this session?</div>
                  <b>Date:</b> {ds.get('timestamp','')[:10]}<br>
                  <b>Components:</b> {cn}<br>
                  <b>Duration:</b> {fmt_dur(ds['hours'])}
                </div>""", unsafe_allow_html=True)
                c1,c2,_ = st.columns([1,1,4])
                if c1.button("Yes, delete",type="primary",key="confirm_yes"):
                    data["sessions"]=[s for s in data["sessions"] if s["id"]!=st.session_state.confirm_delete_id]
                    save_data(data); st.session_state.data=data; st.session_state.confirm_delete_id=None; st.rerun()
                if c2.button("Cancel",key="confirm_no"):
                    st.session_state.confirm_delete_id=None; st.rerun()

    if not data["sessions"]:
        st.info("No sessions logged yet.")
    else:
        ss = sorted(data["sessions"],key=lambda x:x.get("timestamp",""),reverse=True)
        rows_e = []
        for s in ss:
            cn = ", ".join(c["name"] for c in data["components"] if c["id"] in s.get("components",[])) or "—"
            rows_e.append({"Date":s.get("timestamp","")[:10],"Components":cn,
                           "Duration":fmt_dur(s["hours"]),"Hours":round(s["hours"],4),"Notes":s.get("notes","")})
        st.download_button("Export as CSV",data=pd.DataFrame(rows_e).to_csv(index=False).encode(),
            file_name=f"bench_log_{datetime.now().strftime('%Y%m%d')}.csv",mime="text/csv",key="export_log")
        st.divider()
        h1,h2,h3,h4,h5 = st.columns([1.5,2.5,1.2,2.5,1.2])
        h1.markdown("**Date**"); h2.markdown("**Components**")
        h3.markdown("**Duration**"); h4.markdown("**Notes**"); h5.markdown("**Action**")
        st.markdown("<hr style='margin:4px 0 8px 0'>", unsafe_allow_html=True)
        for s in ss:
            cn = ", ".join(c["name"] for c in data["components"] if c["id"] in s.get("components",[])) or "—"
            c1,c2,c3,c4,c5 = st.columns([1.5,2.5,1.2,2.5,1.2])
            c1.write(s.get("timestamp","")[:10]); c2.write(cn)
            c3.write(fmt_dur(s["hours"])); c4.write(s.get("notes","") or "—")
            if c5.button("Delete",key=f"del_s_{s['id']}"):
                st.session_state.confirm_delete_id=s["id"]; st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# ── COMPONENTS ───────────────────────────────────────────────────────────────
if page("Components"):
    if st.session_state.confirm_delete_comp:
        dc = next((c for c in data["components"] if c["id"]==st.session_state.confirm_delete_comp),None)
        if dc:
            _,mc,_ = st.columns([1,2,1])
            with mc:
                hu = get_comp_hours(dc["id"],data["sessions"])
                st.markdown(f"""
                <div style="background:#fff;border:2px solid #f0ad4e;border-radius:12px;
                            padding:24px 28px;margin:16px 0;text-align:center;">
                  <div style="font-size:32px;margin-bottom:8px;">⚠️</div>
                  <div style="font-size:17px;font-weight:600;margin-bottom:10px;">Remove this component?</div>
                  <div style="font-size:13px;color:#555;line-height:1.8;">
                    <b>{dc['name']}</b> ({dc['type']})<br>
                    {fmt_dur(hu)} logged<br>
                    <span style="color:#A32D2D;">Sessions keep their data.</span>
                  </div>
                </div>""", unsafe_allow_html=True)
                r1,r2 = st.columns(2)
                if r1.button("Yes, remove",type="primary",key="confirm_comp_yes",use_container_width=True):
                    data["components"]=[c for c in data["components"] if c["id"]!=st.session_state.confirm_delete_comp]
                    save_data(data); st.session_state.data=data
                    st.session_state.confirm_delete_comp=None; st.rerun()
                if r2.button("Cancel",key="confirm_comp_no",use_container_width=True):
                    st.session_state.confirm_delete_comp=None; st.rerun()
            st.stop()

    st.subheader("Add a component")
    ct, cb, cs, cl = st.columns([1, 1.5, 1.5, 1])
    comp_type  = ct.selectbox("Type", ["Motor","ESC","Prop","Battery"], key="main_comp_type")
    brands     = data.get("brands", DEFAULT_BRANDS.copy())
    type_brands = sorted([b for b,t in brands.items() if t == comp_type])
    comp_brand = cb.selectbox("Brand", [""] + type_brands + ["Other (type below)"], key="main_comp_brand")
    comp_other = cs.text_input("Brand (if Other)" if comp_brand == "Other (type below)" else "Serial number",
                               placeholder="Serial number" if comp_brand != "Other (type below)" else "Brand name",
                               key="main_comp_serial")
    comp_limit = cl.number_input("Max hours", min_value=0.0, step=1.0, value=50.0, key="main_comp_limit")
    if st.button("Add component", type="primary", key="add_comp_main"):
        if comp_brand == "Other (type below)":
            final_brand = comp_other.strip()
            final_serial = ""
        else:
            final_brand = comp_brand
            final_serial = comp_other.strip()
        if not final_brand:
            st.error("Please select or enter a brand.")
        else:
            comp_name = f"{final_brand}_{final_serial}" if final_serial else final_brand
            exists = any(c["name"].lower() == comp_name.lower() for c in data["components"])
            if exists:
                st.warning(f"{comp_name} already exists.")
            else:
                new_id = f"{comp_type.lower()}_{final_brand.lower()}_{final_serial.lower()}" if final_serial else f"{comp_type.lower()}_{final_brand.lower()}_{datetime.now().timestamp():.0f}"
                data["components"].append({
                    "id": new_id, "type": comp_type, "name": comp_name,
                    "limit": comp_limit if comp_limit > 0 else None
                })
                save_data(data); st.session_state.data = data
                st.success(f"Added {comp_type}: {comp_name}")

    st.divider()
    for ctype in ["Motor","ESC","Prop","Battery"]:
        items = [c for c in data["components"] if c["type"]==ctype]
        if not items: continue
        st.subheader(f"{ctype}s")
        for comp in items:
            hours = get_comp_hours(comp["id"],data["sessions"])
            limit = comp.get("limit")
            pct   = min((hours/limit) if limit else 0, 1.0)
            s_key,s_label = comp_status(hours,limit)
            bar_color = "#639922" if s_key=="ok" else "#EF9F27" if s_key=="warning" else "#E24B4A"
            is_editing = st.session_state.editing_comp_id==comp["id"]
            if is_editing:
                e1,e2,e3,e4 = st.columns([2.5,1.5,0.9,0.9])
                new_name = e1.text_input("Name",value=comp["name"],key=f"edit_name_{comp['id']}",label_visibility="collapsed")
                opts = ["Motor","ESC","Prop","Battery"]
                new_type = e2.selectbox("Type",opts,index=opts.index(comp["type"]) if comp["type"] in opts else 0,
                    key=f"edit_type_{comp['id']}",label_visibility="collapsed")
                if e3.button("Save name",key=f"save_name_{comp['id']}"):
                    if new_name.strip():
                        comp["name"]=new_name.strip(); comp["type"]=new_type
                        save_data(data); st.session_state.data=data
                        st.session_state.editing_comp_id=None; st.rerun()
                if e4.button("Cancel",key=f"cancel_edit_{comp['id']}"):
                    st.session_state.editing_comp_id=None; st.rerun()
            else:
                b,l,sv,ed,rm = st.columns([3.5,1.2,0.9,0.8,0.8])
                with b:
                    st.markdown(f"""
                    <div style="margin-bottom:2px;">
                      <b style="font-size:14px;">{comp['name']}</b>
                      <span style="font-size:12px;color:#777;margin-left:8px;">
                        {fmt_dur(hours)} / {fmt_dur(limit) if limit else "no limit"}
                      </span>
                      &nbsp;<span class="badge-{s_key}">{s_label}</span>
                    </div>
                    <div style="background:#e8e8e6;border-radius:2px;height:4px;margin-top:3px;">
                      <div style="width:{pct*100:.0f}%;background:{bar_color};height:4px;border-radius:2px;"></div>
                    </div>""", unsafe_allow_html=True)
                new_limit = l.number_input("Limit",min_value=0.0,step=1.0,
                    value=float(limit) if limit else 0.0,
                    key=f"limit_{comp['id']}",label_visibility="collapsed")
                if sv.button("Save",key=f"save_limit_{comp['id']}"):
                    comp["limit"]=new_limit if new_limit>0 else None
                    save_data(data); st.session_state.data=data; st.success("Saved!"); st.rerun()
                if ed.button("Edit",key=f"edit_{comp['id']}"):
                    st.session_state.editing_comp_id=comp["id"]; st.rerun()
                if rm.button("Remove",key=f"del_{comp['id']}"):
                    st.session_state.confirm_delete_comp=comp["id"]; st.rerun()
            st.markdown("<hr style='margin:4px 0 10px;border:none;border-top:1px solid #f0f0f0;'>",
                        unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# ── RIG CONFIGURATION ────────────────────────────────────────────────────────
if page("Rig configuration"):
    st.subheader("Rig configuration")
    cfg_extra_fields = ["Test objective", "Operator", "Notes"]

    # Active side selector
    st.markdown("**Active setup**")
    active_side = st.radio("Currently running",
        ["Side A", "Side B", "Both"],
        index=["Side A","Side B","Both"].index(data.get("active_side","Side A")),
        horizontal=True, key="active_side_sel")
    if active_side != data.get("active_side"):
        data["active_side"] = active_side
        save_data(data); st.session_state.data = data

    st.divider()

    # ── Inline add-new-component form ─────────────────────────────────────
    with st.expander("+ Add a new component"):
        qc1, qc2 = st.columns(2)
        q_type   = qc1.selectbox("Type", ["Motor","ESC","Prop","Battery"], key="rig_q_type")
        q_limit  = qc2.number_input("Hour limit", min_value=0, value=100, step=10, key="rig_q_limit")
        brands     = data.get("brands", DEFAULT_BRANDS.copy())
        type_brands_rig = sorted([b for b,t in brands.items() if t == q_type])
        qc3, qc4 = st.columns(2)
        q_brand  = qc3.selectbox("Brand", [""] + type_brands_rig + ["Other (type below)"], key="rig_q_brand")
        q_serial = qc4.text_input(
            "Brand (if Other)" if q_brand == "Other (type below)" else "Serial number",
            placeholder="Brand name" if q_brand == "Other (type below)" else "e.g. FX124",
            key="rig_q_serial")
        if st.button("Add component", key="rig_add_comp"):
            if q_brand == "Other (type below)":
                final_brand = q_serial.strip(); final_serial = ""
            else:
                final_brand = q_brand; final_serial = q_serial.strip()
            if not final_brand:
                st.error("Please select or enter a brand.")
            else:
                new_name = f"{final_brand}_{final_serial}" if final_serial else final_brand
                exists   = any(c["name"].lower() == new_name.lower() for c in data["components"])
                if exists:
                    st.warning(f"{new_name} already exists.")
                else:
                    new_id = (f"{q_type.lower()}_{final_brand.lower()}_{final_serial.lower()}"
                              if final_serial else
                              f"{q_type.lower()}_{final_brand.lower()}_{datetime.now().timestamp():.0f}")
                    data["components"].append({
                        "id": new_id, "name": new_name, "type": q_type,
                        "hours": 0.0, "sessions": 0,
                        "limit": float(q_limit) if q_limit else None,
                    })
                    save_data(data); st.session_state.data = data
                    st.success(f"Added {q_type}: {new_name}")
                    st.rerun()

    st.divider()
    side_a_col, side_b_col = st.columns(2)

    ESC_ROLES = ["", "Master", "Slave", "ESC"]

    def render_side(side_key, label, prefix):
        st.markdown(f"**{label}**")
        cfg = data.get(side_key, {})
        upd = {}

        # Motor
        motor_opts = [""] + [c["name"] for c in data["components"] if c["type"]=="Motor"]
        cur = cfg.get("Motor","")
        upd["Motor"] = st.selectbox("Motor", motor_opts,
            index=motor_opts.index(cur) if cur in motor_opts else 0,
            key=f"{prefix}_Motor")

        # ESC
        esc_opts = [""] + [c["name"] for c in data["components"] if c["type"]=="ESC"]
        st.markdown("<span style='font-size:13px;font-weight:600;'>ESC</span>", unsafe_allow_html=True)
        ec1, ec2 = st.columns([3,2])
        cur_esc = cfg.get("ESC","")
        upd["ESC"] = ec1.selectbox("ESC", esc_opts,
            index=esc_opts.index(cur_esc) if cur_esc in esc_opts else 0,
            key=f"{prefix}_ESC", label_visibility="collapsed")
        cur_role = cfg.get("ESC_role", "")
        upd["ESC_role"] = ec2.selectbox("Role", ESC_ROLES,
            index=ESC_ROLES.index(cur_role) if cur_role in ESC_ROLES else 0,
            key=f"{prefix}_ESC_role", label_visibility="collapsed")

        # ESC 2 (optional second ESC)
        show_esc2_key = f"{prefix}_show_esc2"
        if show_esc2_key not in st.session_state:
            st.session_state[show_esc2_key] = bool(cfg.get("ESC2",""))

        if not st.session_state[show_esc2_key]:
            if st.button("+ Add second ESC", key=f"{prefix}_add_esc2"):
                st.session_state[show_esc2_key] = True
                st.rerun()
        else:
            st.markdown("<span style='font-size:13px;font-weight:600;'>ESC</span>", unsafe_allow_html=True)
            ec3, ec4, ec5 = st.columns([3, 2, 1])
            cur_esc2 = cfg.get("ESC2","")
            upd["ESC2"] = ec3.selectbox("ESC2", esc_opts,
                index=esc_opts.index(cur_esc2) if cur_esc2 in esc_opts else 0,
                key=f"{prefix}_ESC2", label_visibility="collapsed")
            cur_role2 = cfg.get("ESC2_role","")
            upd["ESC2_role"] = ec4.selectbox("Role2", ESC_ROLES,
                index=ESC_ROLES.index(cur_role2) if cur_role2 in ESC_ROLES else 0,
                key=f"{prefix}_ESC2_role", label_visibility="collapsed")
            if ec5.button("✕", key=f"{prefix}_rem_esc2"):
                st.session_state[show_esc2_key] = False
                upd["ESC2"] = ""
                upd["ESC2_role"] = ""

        # Prop
        prop_opts = [""] + [c["name"] for c in data["components"] if c["type"]=="Prop"]
        cur_prop = cfg.get("Prop","")
        upd["Prop"] = st.selectbox("Prop", prop_opts,
            index=prop_opts.index(cur_prop) if cur_prop in prop_opts else 0,
            key=f"{prefix}_Prop")

        # Prop 2 (optional second prop)
        show_prop2_key = f"{prefix}_show_prop2"
        if show_prop2_key not in st.session_state:
            st.session_state[show_prop2_key] = bool(cfg.get("Prop2",""))

        if not st.session_state[show_prop2_key]:
            if st.button("+ Add second prop", key=f"{prefix}_add_prop2"):
                st.session_state[show_prop2_key] = True
                st.rerun()
        else:
            pc1, pc2 = st.columns([5, 1])
            cur_prop2 = cfg.get("Prop2","")
            upd["Prop2"] = pc1.selectbox("Prop2", prop_opts,
                index=prop_opts.index(cur_prop2) if cur_prop2 in prop_opts else 0,
                key=f"{prefix}_Prop2", label_visibility="collapsed")
            if pc2.button("✕", key=f"{prefix}_rem_prop2"):
                st.session_state[show_prop2_key] = False
                upd["Prop2"] = ""

        # Battery
        bat_opts = [""] + [c["name"] for c in data["components"] if c["type"]=="Battery"]
        cur_bat = cfg.get("Battery","")
        upd["Battery"] = st.selectbox("Battery", bat_opts,
            index=bat_opts.index(cur_bat) if cur_bat in bat_opts else 0,
            key=f"{prefix}_Battery")

        # Extra fields
        for f in cfg_extra_fields:
            upd[f] = st.text_input(f, value=cfg.get(f,""), key=f"{prefix}_extra_{f}")

        if st.button(f"Save {label}", type="primary", key=f"save_{prefix}"):
            data[side_key] = {k:v for k,v in upd.items() if v}
            save_data(data); st.session_state.data = data
            st.success(f"{label} saved!")

    with side_a_col:
        render_side("rig_side_a", "Side A", "cfg_a")

    with side_b_col:
        render_side("rig_side_b", "Side B", "cfg_b")

# ══════════════════════════════════════════════════════════════════════════════
# ── LOG EVENTS ───────────────────────────────────────────────────────────────
if page("Log events"):
    st.subheader("Log an event")
    st.caption("Record errors, problems, config changes, or any notable event.")

    ev_left, ev_right = st.columns([1, 2])
    with ev_left:
        ev_sev  = st.selectbox("Severity", ["info", "warning", "critical"], key="ev_sev")
        ev_date = st.date_input("Date", value=datetime.now().date(), key="ev_date")
        ev_time = st.time_input("Time", value=datetime.now().time(), key="ev_time")
        ev_text = st.text_area("Event description", key="ev_text",
                               placeholder="e.g. Motor overheated at 120°C during run #5",
                               height=120)
        if st.button("Log event", type="primary", key="log_event"):
            if not ev_text.strip():
                st.error("Please enter a description.")
            else:
                ts = datetime.combine(ev_date, ev_time).isoformat()
                data["events"].append({
                    "id":        f"ev_{datetime.now().timestamp():.0f}",
                    "severity":  ev_sev,
                    "text":      ev_text.strip(),
                    "timestamp": ts,
                })
                save_data(data); st.session_state.data = data
                st.success("Event logged!")

    with ev_right:
        st.markdown("#### Recent events")
        events = sorted(data.get("events", []), key=lambda x: x.get("timestamp",""), reverse=True)
        if not events:
            st.info("No events logged yet.")
        else:
            del_ev_id = None
            for ev in events[:20]:
                sev = ev.get("severity","info")
                ts  = ev.get("timestamp","")[:16].replace("T"," ")
                txt = ev.get("text","")
                ec1, ec2 = st.columns([5, 1])
                with ec1:
                    st.markdown(f"""
                    <div class="ev-{sev}">
                      <span style="font-size:11px;color:#888;">{ts}</span>
                      <span style="font-size:11px;text-transform:uppercase;font-weight:600;margin:0 6px;
                        color:{'#378ADD' if sev=='info' else '#EF9F27' if sev=='warning' else '#E24B4A'};">{sev}</span>
                      <span style="font-size:13px;">{txt}</span>
                    </div>""", unsafe_allow_html=True)
                with ec2:
                    if st.button("Delete", key=f"del_ev_{ev['id']}"):
                        del_ev_id = ev["id"]
            if del_ev_id:
                data["events"] = [e for e in data["events"] if e["id"] != del_ev_id]
                save_data(data); st.session_state.data = data; st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# ── SETTINGS ─────────────────────────────────────────────────────────────────
if page("Settings"):
    st.subheader("Asana integration")
    st.caption("Connect your Asana project to see planned vs actual on the Overview page.")
    asana = data.get("asana", {"token":"","project_gid":"1212056850873170"})
    a1 = st.text_input("Personal Access Token", value=asana.get("token",""),
                        type="password", key="asana_token",
                        help="Asana → My Profile Settings → Apps → Personal Access Tokens → New Token")
    a2 = st.text_input("Project GID", value=asana.get("project_gid","1212056850873170"),
                        key="asana_project",
                        help="Open your project in Asana — copy the number from the URL bar")
    if st.button("Save Asana settings", key="save_asana"):
        data["asana"] = {"token": a1.strip(), "project_gid": a2.strip()}
        save_data(data); st.session_state.data=data; st.success("Asana settings saved!")
    if a1.strip() and a2.strip():
        if st.button("Test connection", key="test_asana"):
            with st.spinner("Connecting to Asana…"):
                tasks = fetch_asana_tasks(a1.strip(), a2.strip())
            if tasks:
                st.success(f"Connected! Found {len(tasks)} task(s) in your project.")
            else:
                st.error("Could not fetch tasks. Check your token and project GID.")
        if st.button("Load sections for filter", key="load_sections_settings"):
            with st.spinner("Loading sections…"):
                secs = fetch_asana_sections(a1.strip(), a2.strip())
            if secs:
                st.success(f"Found {len(secs)} section(s): {', '.join(s['name'] for s in secs)}")
                st.caption("Go to the Overview page to select which section to display.")
            else:
                st.error("Could not load sections.")

    st.divider()
    st.subheader("Cloud sync")

    sid_ok = bool(_get_sheet_id())

    # ── Detailed credential diagnostics ──────────────────────────────────────
    cred_path = CREDENTIALS_FILE
    cred_exists = os.path.exists(cred_path)

    # Try to import libraries
    try:
        import gspread
        from google.oauth2.service_account import Credentials as _Creds
        libs_ok = True
    except ImportError as _ie:
        libs_ok = False
        libs_err = str(_ie)

    # Try to actually build the client and capture any error
    gc_ok = False
    gc_err = ""
    if libs_ok and cred_exists and sid_ok:
        try:
            _scopes = ["https://www.googleapis.com/auth/spreadsheets"]
            _creds  = _Creds.from_service_account_file(cred_path, scopes=_scopes)
            _gc     = gspread.authorize(_creds)
            _gc.open_by_key(_get_sheet_id())   # actually test the connection
            gc_ok = True
        except Exception as _e:
            gc_err = str(_e)

    if gc_ok:
        st.success("Google Sheets sync is active — all saves go to the cloud automatically.")
        c1, c2 = st.columns(2)
        if c1.button("Test connection", key="test_sheets_settings"):
            with st.spinner("Connecting…"):
                result = sheets_load()
            if result is not None:
                st.success("Connected! Google Sheets is reachable and working.")
            else:
                st.error("Could not read from Google Sheets.")
        if c2.button("Pull latest from Sheets now", key="pull_sheets_settings"):
            with st.spinner("Pulling…"):
                latest = sheets_load()
            if latest is not None:
                save_data(_ensure_keys(latest))
                st.session_state.data = load_data()
                st.success("Pulled latest data from Google Sheets!")
                st.rerun()
            else:
                st.error("Could not reach Google Sheets.")
    else:
        st.warning("Not connected — see diagnostics below.")
        # Show each check clearly
        st.markdown("**Diagnostics:**")
        # Library check
        if libs_ok:
            st.markdown("- gspread + google-auth: **installed**")
        else:
            st.error(f"- gspread / google-auth not installed. Run: `py -m pip install gspread google-auth`")
            st.markdown(f"  Error: `{libs_err}`")
        # File check
        st.markdown(f"- Looking for credentials.json at:")
        st.code(cred_path)
        if cred_exists:
            st.markdown("- credentials.json: **found**")
            # Peek inside to confirm it's valid JSON with expected fields
            try:
                with open(cred_path) as _f:
                    _j = json.load(_f)
                has_email = "client_email" in _j
                has_key   = "private_key" in _j
                st.markdown(f"- File is valid JSON: **yes**")
                st.markdown(f"- client_email present: **{'yes — ' + _j.get('client_email','') if has_email else 'NO — wrong file?'}**")
                st.markdown(f"- private_key present: **{'yes' if has_key else 'NO — wrong file?'}**")
            except Exception as _pe:
                st.error(f"- File exists but could not be read: {_pe}")
        else:
            st.error("- credentials.json: **NOT FOUND at that path**")
        # Sheet ID check
        if sid_ok:
            st.markdown(f"- Sheet ID: **{_get_sheet_id()}**")
        else:
            st.error("- Sheet ID: not set")
        # Connection error
        if gc_err:
            st.error(f"- Connection error: {gc_err}")

    st.markdown(f"""
    <div style="background:#fff;border:1px solid #e0e0e0;padding:14px 18px;font-size:13px;margin-top:12px;">
      <div style="font-weight:500;margin-bottom:4px;">Sheet ID</div>
      <div style="color:#555;font-family:monospace;font-size:12px;">{SHEET_ID or "— not set —"}</div>
      <div style="color:#888;font-size:11px;margin-top:6px;">
        Local backup: <span style="font-family:monospace;">{os.path.abspath(DATA_FILE)}</span>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Migrate local data to cloud ───────────────────────────────────────────
    if gc_ok and sid_ok:
        st.divider()
        st.subheader("Migrate local data to cloud")
        st.caption(
            "If the Google Sheet is empty and your local bench_data.json has historical "
            "records, use this to push everything to the cloud in one step."
        )
        local_exists = os.path.exists(DATA_FILE)
        local_sessions = 0
        if local_exists:
            try:
                with open(DATA_FILE, "r") as _f:
                    _local = json.load(_f)
                local_sessions = len(_local.get("sessions", []))
            except Exception:
                pass

        if local_exists and local_sessions > 0:
            st.info(f"Local file contains **{local_sessions}** session(s). "
                    f"The cloud will be overwritten with this data.")
            if st.button("Migrate local data to Google Sheets", key="migrate_to_cloud",
                         type="primary"):
                with st.spinner("Pushing local data to Google Sheets…"):
                    try:
                        with open(DATA_FILE, "r") as _f:
                            local_data = _ensure_keys(json.load(_f))
                        ok = sheets_save(local_data)
                    except Exception as e:
                        ok = False
                        st.error(f"Error reading local file: {e}")
                if ok:
                    st.success(f"Done! {local_sessions} session(s) migrated to Google Sheets. "
                               f"All computers will now see this data.")
                    st.session_state.data = load_data()
                else:
                    st.error("Migration failed. Check your credentials.json and Sheet ID.")
        elif local_exists and local_sessions == 0:
            st.info("Local file exists but contains no sessions — nothing to migrate.")
        else:
            st.info("No local bench_data.json found — nothing to migrate.")

    st.divider()
    st.subheader("Data backup")
    data_path = os.path.abspath(DATA_FILE)

    col_exp, col_imp = st.columns(2)
    with col_exp:
        st.markdown("**Export backup**")
        backup_json = json.dumps(data, indent=2)
        st.download_button(
            "Download bench_data.json",
            data=backup_json,
            file_name=f"bench_data_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
            mime="application/json",
            key="export_backup"
        )
    with col_imp:
        st.markdown("**Import backup**")
        uploaded = st.file_uploader("Upload bench_data.json", type="json", key="import_backup")
        if uploaded:
            try:
                imported = json.loads(uploaded.read())
                save_data(_ensure_keys(imported))
                st.session_state.data = load_data()
                st.success("Data imported successfully!")
                st.rerun()
            except Exception as e:
                st.error(f"Import failed: {e}")

    st.divider()
    st.subheader("File import")
    st.caption("Drop .edvm, .csv, or .wnq files into the test_data folder next to dashboard.py.")
    if st.button("Scan test_data folder", key="local_scan_settings"):
        Path(os.path.join(_BASE_DIR, "test_data")).mkdir(exist_ok=True)
        count, errors = import_local_files(data)
        st.session_state.data = load_data(); data = st.session_state.data
        if count:
            st.success(f"Imported {count} new file(s).")
        else:
            st.info("No new files found.")
        for e in errors:
            st.error(e)

    st.divider()
    st.subheader("Maintenance")
    if st.button("Remove duplicate components", key="dedup_settings"):
        seen, keep, removed = {}, [], 0
        for c in data["components"]:
            k = (c["name"].lower(), c["type"])
            if k not in seen:
                seen[k] = True; keep.append(c)
            else:
                removed += 1
        data["components"] = keep
        save_data(data); st.session_state.data = data
        st.success(f"Done — removed {removed} duplicate(s)." if removed else "No duplicates found.")


