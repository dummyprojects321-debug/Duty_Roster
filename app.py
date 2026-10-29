"""
Duty Roster Command Center  (v2)
--------------------------------
Create, reuse and print daily duty rosters for multi-station enforcement teams.

Workflow:  Employees  ->  Beats (Settings)  ->  Roster Builder  ->  PDF

Run:  streamlit run app.py
"""

import hashlib
import io
import os
import re
import secrets
import sqlite3
import time as _time
from contextlib import contextmanager
from datetime import date, datetime, time, timedelta
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st

# =============================================================================
# CONFIG  (safe to edit)
# =============================================================================
DB_FILE = os.environ.get("DUTY_DB_PATH", "duty_roster.db")
SEED_FILE = "DUTY_ROSTER_Standard.xlsx"          # optional: used once to pre-load data
ORG_LINE = "PERA  |  DERB Gujranwala"            # printed in the PDF header
APP_TITLE = "Duty Roster Command Center"

STATIONS = [
    "GUJRANWALA CITY",
    "GUJRANWALA SADDAR",
    "KAMOKE",
    "NOWSHEHRA VIRKAN",
    "WAZIRABAD",
]
ADMIN = "MASTER_ADMIN"

# First-run passwords only. After first login, change them in  Settings > Account.
DEFAULT_PASSWORDS = {
    "MASTER_ADMIN": "admin123",
    "KAMOKE": "kamoke2026",
    "NOWSHEHRA VIRKAN": "nowshehra2026",
    "GUJRANWALA CITY": "city2026",
    "WAZIRABAD": "wazirabad2026",
    "GUJRANWALA SADDAR": "saddar2026",
}

DEFAULT_SHIFTS = [  # name, start, end
    ("Morning", "09:00", "17:00"),
    ("Mid-Day", "10:00", "20:00"),
    ("Evening", "13:30", "22:00"),
    ("Night", "20:00", "06:00"),
    ("Sunday Duty", "17:00", "22:00"),
]
DEFAULT_DESIGNATIONS = [
    "SENIOR SERGEANT", "SERGEANT", "ENFORCEMENT OFFICER",
    "INVESTIGATION OFFICER", "DRIVER", "MUHARRAR",
]
REMOVAL_REASONS = [
    "Transferred out of district", "Retired", "Resigned",
    "Terminated / dismissed", "Deputation / on long leave", "Other",
]

COLS = [
    "station", "name", "belt_no", "designation", "contact_no",
    "squad_team", "beat_area", "shift", "start_time", "end_time", "remarks",
]
PHONE_RE = re.compile(r"^03\d{2}-\d{7}$")


# =============================================================================
# SMALL HELPERS
# =============================================================================


def _blank(v):
    if v is None:
        return True
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def _s(v):
    return "" if _blank(v) else str(v).strip()


def to_hhmm(v):
    """Anything time-like -> 'HH:MM' (or '')."""
    if _blank(v):
        return ""
    if isinstance(v, (time, datetime)):
        return v.strftime("%H:%M")
    s = str(v).strip()
    for f in ("%H:%M:%S", "%H:%M", "%I:%M %p", "%I:%M:%S %p"):
        try:
            return datetime.strptime(s, f).strftime("%H:%M")
        except ValueError:
            pass
    return ""


def to_time(s):
    s = to_hhmm(s)
    return datetime.strptime(s, "%H:%M").time() if s else None


def fmt12(s):
    s = to_hhmm(s)
    if not s:
        return "—"
    return datetime.strptime(s, "%H:%M").strftime("%I:%M %p").lstrip("0")


def norm_date(v, fallback):
    s = _s(v)
    for f in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, f).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return fallback


def norm_phone(v):
    s = _s(v)
    digits = re.sub(r"\D", "", s)
    if len(digits) == 11 and digits.startswith("03"):
        return f"{digits[:4]}-{digits[4:]}"
    return s


def minutes(t):
    t = to_hhmm(t)
    if not t:
        return None
    h, m = t.split(":")
    return int(h) * 60 + int(m)


def person_key(row):
    b = _s(row["belt_no"]).upper()
    return b if b else _s(row["name"]).lower()


# =============================================================================
# DATABASE
# =============================================================================


@contextmanager
def db():
    conn = sqlite3.connect(DB_FILE)
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def _hash(pw, salt):
    return hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), 120_000).hex()


def set_password(conn, username, pw):
    salt = secrets.token_hex(16)
    conn.execute(
        "INSERT OR REPLACE INTO users (username, salt, pw_hash) VALUES (?,?,?)",
        (username, salt, _hash(pw, salt)),
    )


def verify_password(username, pw):
    with db() as c:
        row = c.execute(
            "SELECT salt, pw_hash FROM users WHERE username=?", (username,)
        ).fetchone()
    return bool(row) and secrets.compare_digest(_hash(pw, row[0]), row[1])


def audit(conn, user, action, detail=""):
    conn.execute(
        "INSERT INTO audit (ts, user, action, detail) VALUES (?,?,?,?)",
        (datetime.now().isoformat(timespec="seconds"), user, action, detail),
    )




# ---- key helpers --------------------------------------------------------------


def emp_key(belt, name):
    b = _s(belt).upper()
    return b if b else _s(name).lower()


def today_iso():
    return date.today().strftime("%Y-%m-%d")


# ---- schema / migration / seed ------------------------------------------------


def _backfill_employees(c):
    """Every roster row must point at an employee. Create employees from legacy rows."""
    known = {}
    for eid, stn, belt, name in c.execute("SELECT id, station, belt_no, name FROM employees"):
        known[(stn, emp_key(belt, name))] = eid
    rows = c.execute(
        "SELECT id, station, name, belt_no, designation, contact_no FROM roster "
        "WHERE emp_id IS NULL ORDER BY roster_date, id").fetchall()
    now = datetime.now().isoformat(timespec="seconds")
    for rid, stn, name, belt, des, ph in rows:
        if not _s(name):
            continue
        k = (stn, emp_key(belt, name))
        if k not in known:
            cur = c.execute(
                "INSERT INTO employees (station,name,belt_no,designation,contact_no,status,created_at) "
                "VALUES (?,?,?,?,?,'active',?)",
                (stn, _s(name), _s(belt).upper(), _s(des).upper(), norm_phone(ph), now))
            known[k] = cur.lastrowid
        c.execute("UPDATE roster SET emp_id=? WHERE id=?", (known[k], rid))


def _seed_lists(c):
    if c.execute("SELECT COUNT(*) FROM lists").fetchone()[0]:
        return
    for kind, col in (("beat", "beat_area"), ("squad", "squad_team")):
        seen = set()
        for stn, val in c.execute(
                f"SELECT station, {col} FROM roster WHERE {col} IS NOT NULL AND {col}<>'' ORDER BY id"):
            if (stn, val.lower()) in seen:
                continue
            seen.add((stn, val.lower()))
            n = sum(1 for k in seen if k[0] == stn)
            c.execute("INSERT INTO lists (kind, station, name, sort) VALUES (?,?,?,?)", (kind, stn, val, n))


@st.cache_resource(show_spinner=False)
def init_db():
    today = today_iso()
    with db() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS roster (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                station TEXT, roster_date TEXT, name TEXT, belt_no TEXT,
                designation TEXT, contact_no TEXT, squad_team TEXT,
                beat_area TEXT, shift TEXT, start_time TEXT, end_time TEXT,
                remarks TEXT DEFAULT '', updated_at TEXT, updated_by TEXT, emp_id INTEGER
            )""")
        have = {r[1] for r in c.execute("PRAGMA table_info(roster)")}
        for col, typ in (("remarks", "TEXT DEFAULT ''"), ("updated_at", "TEXT"),
                         ("updated_by", "TEXT"), ("emp_id", "INTEGER")):
            if col not in have:
                c.execute(f"ALTER TABLE roster ADD COLUMN {col} {typ}")
        c.execute("CREATE INDEX IF NOT EXISTS ix_roster_date ON roster(roster_date, station)")
        c.execute("CREATE INDEX IF NOT EXISTS ix_roster_emp ON roster(emp_id, roster_date)")
        c.execute("""CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT, station TEXT, name TEXT, belt_no TEXT,
            designation TEXT, contact_no TEXT, status TEXT DEFAULT 'active',
            joined_on TEXT, left_on TEXT, left_reason TEXT, created_at TEXT)""")
        c.execute("CREATE INDEX IF NOT EXISTS ix_emp_station ON employees(station, status)")
        c.execute("""CREATE TABLE IF NOT EXISTS transfers (
            id INTEGER PRIMARY KEY AUTOINCREMENT, emp_id INTEGER, name TEXT,
            from_station TEXT, to_station TEXT, effective TEXT, by_user TEXT, ts TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS lists (
            id INTEGER PRIMARY KEY AUTOINCREMENT, kind TEXT, station TEXT, name TEXT, sort INTEGER)""")
        c.execute("""CREATE TABLE IF NOT EXISTS templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT, station TEXT, name TEXT,
            created_by TEXT, created_at TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS template_rows (
            id INTEGER PRIMARY KEY AUTOINCREMENT, template_id INTEGER, emp_id INTEGER, name TEXT,
            squad_team TEXT, beat_area TEXT, shift TEXT, start_time TEXT, end_time TEXT, remarks TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS shifts (
            name TEXT PRIMARY KEY, start_time TEXT, end_time TEXT, sort INTEGER)""")
        c.execute("CREATE TABLE IF NOT EXISTS options (kind TEXT, value TEXT, sort INTEGER)")
        c.execute("""CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY, salt TEXT, pw_hash TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS audit (
            id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, user TEXT, action TEXT, detail TEXT)""")

        if c.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
            for u, p in DEFAULT_PASSWORDS.items():
                set_password(c, u, p)

        # normalise legacy rows (dates -> ISO, times -> HH:MM)
        for rid, d, s_, e_ in c.execute(
                "SELECT id, roster_date, start_time, end_time FROM roster").fetchall():
            nd, ns, ne = norm_date(d, today), to_hhmm(s_), to_hhmm(e_)
            if (nd, ns, ne) != (d, s_, e_):
                c.execute("UPDATE roster SET roster_date=?, start_time=?, end_time=? WHERE id=?",
                          (nd, ns, ne, rid))

        # one-time seed from the Excel template
        empty = (c.execute("SELECT COUNT(*) FROM roster").fetchone()[0] == 0
                 and c.execute("SELECT COUNT(*) FROM employees").fetchone()[0] == 0)
        if empty and os.path.exists(SEED_FILE):
            try:
                df = pd.read_excel(SEED_FILE, sheet_name=0)
                df = df.rename(columns={
                    "Station": "station", "Roster Date": "roster_date", "Name": "name",
                    "Belt No.": "belt_no", "Designation": "designation",
                    "Contact No.": "contact_no", "Squad / Team": "squad_team",
                    "Beat Area": "beat_area", "Shift": "shift",
                    "Start Time": "start_time", "End Time": "end_time"})
                valid = [norm_date(x, "") for x in df["roster_date"]]
                seed_date = next((d for d in valid if d), today)
                for col in COLS:
                    if col not in df:
                        df[col] = ""
                out = pd.DataFrame({k: df[k].map(_s) for k in COLS if k not in ("start_time", "end_time")})
                out["start_time"] = df["start_time"].map(to_hhmm)
                out["end_time"] = df["end_time"].map(to_hhmm)
                out["contact_no"] = out["contact_no"].map(norm_phone)
                out["belt_no"] = out["belt_no"].str.upper()
                out["designation"] = out["designation"].str.upper()
                out["roster_date"] = seed_date
                out["updated_at"] = datetime.now().isoformat(timespec="seconds")
                out["updated_by"] = "seed"
                out.to_sql("roster", c, if_exists="append", index=False)
            except Exception as e:  # noqa: BLE001
                print("Seed error:", e)

        if c.execute("SELECT COUNT(*) FROM shifts").fetchone()[0] == 0:
            known = {n for n, _, _ in DEFAULT_SHIFTS}
            extra = [r[0] for r in c.execute(
                "SELECT DISTINCT shift FROM roster WHERE shift IS NOT NULL AND shift<>''") if r[0] not in known]
            for i, (n, s_, e_) in enumerate(DEFAULT_SHIFTS + [(x, "09:00", "17:00") for x in extra]):
                c.execute("INSERT INTO shifts VALUES (?,?,?,?)", (n, s_, e_, i))
        if c.execute("SELECT COUNT(*) FROM options WHERE kind='designation'").fetchone()[0] == 0:
            for i, d in enumerate(DEFAULT_DESIGNATIONS):
                c.execute("INSERT INTO options VALUES ('designation',?,?)", (d, i))

        _backfill_employees(c)
        _seed_lists(c)
    return True


# ---- simple loaders -----------------------------------------------------------


def load_shifts():
    with db() as c:
        return pd.read_sql_query("SELECT name, start_time, end_time FROM shifts ORDER BY sort", c)


def load_designations(extra=()):
    with db() as c:
        base = [r[0] for r in c.execute("SELECT value FROM options WHERE kind='designation' ORDER BY sort")]
        used = [r[0] for r in c.execute(
            "SELECT DISTINCT designation FROM employees WHERE designation<>'' AND designation IS NOT NULL")]
    out = list(base)
    for d in list(used) + list(extra):
        if d and d not in out:
            out.append(d)
    return out


def sort_df(df, shift_order):
    if df.empty:
        return df
    df = df.copy()
    df["_st"] = df["station"].map({s: i for i, s in enumerate(STATIONS)}).fillna(99)
    df["_sh"] = df["shift"].map(shift_order).fillna(99)
    df = df.sort_values(["_st", "_sh", "start_time", "squad_team", "beat_area", "name", "id"])
    return df.drop(columns=["_st", "_sh"]).reset_index(drop=True)


def load_day(date_iso, stations):
    stations = list(stations)
    ph = ",".join("?" * len(stations))
    with db() as c:
        df = pd.read_sql_query(
            f"SELECT * FROM roster WHERE roster_date=? AND station IN ({ph})", c,
            params=[date_iso, *stations])
    for col in COLS:
        df[col] = df[col].map(_s)
    df["start_time"] = df["start_time"].map(to_hhmm)
    df["end_time"] = df["end_time"].map(to_hhmm)
    df["emp_id"] = pd.to_numeric(df["emp_id"], errors="coerce")
    sh = load_shifts()
    return sort_df(df, {n: i for i, n in enumerate(sh["name"])})


def load_employees(stations, status="Active"):
    stations = list(stations)
    ph = ",".join("?" * len(stations))
    q = f"SELECT * FROM employees WHERE station IN ({ph})"
    if status == "Active":
        q += " AND status='active'"
    elif status == "Inactive":
        q += " AND status<>'active'"
    with db() as c:
        df = pd.read_sql_query(q, c, params=stations)
    for col in ("name", "belt_no", "designation", "contact_no", "left_on", "left_reason", "joined_on"):
        df[col] = df[col].map(_s)
    df["_st"] = df["station"].map({s: i for i, s in enumerate(STATIONS)}).fillna(99)
    df = df.sort_values(["_st", "name"]).drop(columns="_st").reset_index(drop=True)
    return df


def get_list(kind, station):
    with db() as c:
        return [r[0] for r in c.execute(
            "SELECT name FROM lists WHERE kind=? AND station=? ORDER BY sort, name", (kind, station))]


def set_list(kind, station, names, user):
    clean = list(dict.fromkeys(_s(n) for n in names if _s(n)))
    with db() as c:
        c.execute("DELETE FROM lists WHERE kind=? AND station=?", (kind, station))
        for i, n in enumerate(clean):
            c.execute("INSERT INTO lists (kind, station, name, sort) VALUES (?,?,?,?)", (kind, station, n, i))
        audit(c, user, "LIST", f"{station}: {kind} list saved ({len(clean)})")
    return len(clean)


def add_to_list(kind, station, names, user):
    existing = get_list(kind, station)
    low = {x.lower() for x in existing}
    add = []
    for n in names:
        n = _s(n)
        if n and n.lower() not in low:
            add.append(n)
            low.add(n.lower())
    if add:
        set_list(kind, station, existing + add, user)
    return len(add)


# ---- roster operations --------------------------------------------------------


def _insert_roster(c, date_iso, rec, emp_id, user, now):
    c.execute(
        f"INSERT INTO roster (roster_date,{','.join(COLS)},updated_at,updated_by,emp_id) "
        f"VALUES (?,{','.join('?' * len(COLS))},?,?,?)",
        [date_iso, *[rec[k] for k in COLS], now, user, int(emp_id)])


def assign_duty(date_iso, emp_ids, beat, squad, shift, start, end, remarks, user):
    now = datetime.now().isoformat(timespec="seconds")
    added = skipped = 0
    ids = [int(i) for i in emp_ids]
    with db() as c:
        ph = ",".join("?" * len(ids))
        emps = {r[0]: r for r in c.execute(
            f"SELECT id,station,name,belt_no,designation,contact_no FROM employees WHERE id IN ({ph})", ids)}
        existing = {(r[0], r[1], r[2]) for r in c.execute(
            "SELECT emp_id, shift, beat_area FROM roster WHERE roster_date=?", (date_iso,))}
        for eid in ids:
            e = emps.get(eid)
            if not e or (eid, shift, beat) in existing:
                skipped += 1
                continue
            rec = {"station": e[1], "name": e[2], "belt_no": e[3], "designation": e[4], "contact_no": e[5],
                   "squad_team": squad, "beat_area": beat, "shift": shift, "start_time": start,
                   "end_time": end, "remarks": remarks}
            _insert_roster(c, date_iso, rec, eid, user, now)
            added += 1
        if added:
            audit(c, user, "ASSIGN", f"{date_iso}: {added} -> {beat} ({shift})")
    return added, skipped


def copy_rows(src, station, date_iso, mode, user, label):
    """Copy assignment rows (needs emp_id, name, squad_team, beat_area, shift, start/end, remarks)
    into `station` on `date_iso`, refreshing person details and skipping anyone who has left/moved."""
    now = datetime.now().isoformat(timespec="seconds")
    copied, skipped = 0, []
    with db() as c:
        emps = {r[0]: r for r in c.execute(
            "SELECT id,name,belt_no,designation,contact_no FROM employees WHERE station=? AND status='active'",
            (station,))}
        if mode == "Replace existing":
            c.execute("DELETE FROM roster WHERE roster_date=? AND station=?", (date_iso, station))
        existing = {(r[0], r[1], r[2]) for r in c.execute(
            "SELECT emp_id, shift, beat_area FROM roster WHERE roster_date=? AND station=?",
            (date_iso, station))}
        for r in src.itertuples():
            eid = None if _blank(r.emp_id) else int(r.emp_id)
            e = emps.get(eid)
            if e is None:
                skipped.append(_s(r.name) or "unknown")
                continue
            key = (eid, _s(r.shift), _s(r.beat_area))
            if key in existing:
                continue
            existing.add(key)
            rec = {"station": station, "name": e[1], "belt_no": e[2], "designation": e[3],
                   "contact_no": e[4], "squad_team": _s(r.squad_team), "beat_area": _s(r.beat_area),
                   "shift": _s(r.shift), "start_time": to_hhmm(r.start_time),
                   "end_time": to_hhmm(r.end_time), "remarks": _s(r.remarks)}
            _insert_roster(c, date_iso, rec, eid, user, now)
            copied += 1
        audit(c, user, "COPY", f"{station} {date_iso} <- {label} ({mode}): {copied} copied, {len(skipped)} skipped")
    return copied, skipped


EDIT_COLS = ["beat_area", "squad_team", "shift", "start_time", "end_time", "remarks"]


def apply_roster_edits(orig, edited, user, presets):
    o = orig.set_index("id")
    now = datetime.now().isoformat(timespec="seconds")
    stat = {"updated": 0, "removed": 0}
    with db() as c:
        for _, r in edited.iterrows():
            rid = int(r["id"])
            if bool(r.get("Remove")):
                c.execute("DELETE FROM roster WHERE id=?", (rid,))
                stat["removed"] += 1
                continue
            rec = {k: _s(r[k]) for k in ("beat_area", "squad_team", "shift", "remarks")}
            rec["start_time"], rec["end_time"] = to_hhmm(r["start_time"]), to_hhmm(r["end_time"])
            before = {k: _s(o.loc[rid, k]) for k in EDIT_COLS}
            p = presets.get(rec["shift"])
            if p:
                if rec["shift"] != before["shift"] and (rec["start_time"], rec["end_time"]) == (
                        before["start_time"], before["end_time"]):
                    rec["start_time"], rec["end_time"] = p  # shift changed -> adopt its timing
                rec["start_time"] = rec["start_time"] or p[0]
                rec["end_time"] = rec["end_time"] or p[1]
            if before != rec:
                c.execute(
                    f"UPDATE roster SET {','.join(f'{k}=?' for k in EDIT_COLS)}, updated_at=?, updated_by=? WHERE id=?",
                    [*[rec[k] for k in EDIT_COLS], now, user, rid])
                stat["updated"] += 1
        if any(stat.values()):
            audit(c, user, "EDIT", f"roster rows: ~{stat['updated']} -{stat['removed']}")
    return stat


def clear_day(station, date_iso, user):
    with db() as c:
        n = c.execute("DELETE FROM roster WHERE roster_date=? AND station=?", (date_iso, station)).rowcount
        audit(c, user, "CLEAR", f"{station} {date_iso}: {n} rows")
    return n


# ---- templates ----------------------------------------------------------------


def save_template(station, name, day_df, user):
    name = _s(name)
    now = datetime.now().isoformat(timespec="seconds")
    with db() as c:
        for (tid,) in c.execute("SELECT id FROM templates WHERE station=? AND lower(name)=lower(?)",
                                (station, name)).fetchall():
            c.execute("DELETE FROM template_rows WHERE template_id=?", (tid,))
            c.execute("DELETE FROM templates WHERE id=?", (tid,))
        tid = c.execute("INSERT INTO templates (station,name,created_by,created_at) VALUES (?,?,?,?)",
                        (station, name, user, now)).lastrowid
        n = 0
        for r in day_df.itertuples():
            if _blank(r.emp_id):
                continue
            c.execute(
                "INSERT INTO template_rows (template_id,emp_id,name,squad_team,beat_area,shift,start_time,end_time,remarks) "
                "VALUES (?,?,?,?,?,?,?,?,?)",
                (tid, int(r.emp_id), r.name, r.squad_team, r.beat_area, r.shift, r.start_time, r.end_time, r.remarks))
            n += 1
        audit(c, user, "TEMPLATE", f"{station}: saved '{name}' ({n} duties)")
    return n


def list_templates(station):
    with db() as c:
        return pd.read_sql_query(
            "SELECT t.id, t.name, t.created_by, t.created_at, "
            "(SELECT COUNT(*) FROM template_rows r WHERE r.template_id=t.id) AS n "
            "FROM templates t WHERE t.station=? ORDER BY t.name", c, params=[station])


def load_template(tid):
    with db() as c:
        df = pd.read_sql_query(
            "SELECT emp_id, name, squad_team, beat_area, shift, start_time, end_time, remarks "
            "FROM template_rows WHERE template_id=? ORDER BY id", c, params=[int(tid)])
    return df


def delete_template(tid, user):
    with db() as c:
        c.execute("DELETE FROM template_rows WHERE template_id=?", (int(tid),))
        c.execute("DELETE FROM templates WHERE id=?", (int(tid),))
        audit(c, user, "TEMPLATE", f"deleted template {tid}")


# ---- employees ----------------------------------------------------------------


def _find_employee(c, station, belt, name, exclude_id=None):
    k = emp_key(belt, name)
    for eid, b, n, stt in c.execute(
            "SELECT id, belt_no, name, status FROM employees WHERE station=?", (station,)):
        if eid != exclude_id and emp_key(b, n) == k:
            return eid, stt
    return None


def add_employee(station, name, belt, designation, phone, user):
    name, belt = _s(name), _s(belt).upper()
    if not name:
        return False, "Name is required."
    with db() as c:
        hit = _find_employee(c, station, belt, name)
        if hit:
            return False, ("This person already exists in " + station.title() +
                           (" but was removed earlier — reinstate them from the list." if hit[1] != "active"
                            else "."))
        c.execute(
            "INSERT INTO employees (station,name,belt_no,designation,contact_no,status,joined_on,created_at) "
            "VALUES (?,?,?,?,?,'active',?,?)",
            (station, name, belt, _s(designation).upper(), norm_phone(phone), today_iso(),
             datetime.now().isoformat(timespec="seconds")))
        audit(c, user, "EMP_ADD", f"{name} ({belt}) -> {station}")
    return True, f"{name} added to {station.title()}."


def edit_employee(eid, name, belt, designation, phone, user):
    name, belt = _s(name), _s(belt).upper()
    if not name:
        return False, "Name is required."
    with db() as c:
        row = c.execute("SELECT station FROM employees WHERE id=?", (eid,)).fetchone()
        if not row:
            return False, "Employee not found."
        if _find_employee(c, row[0], belt, name, exclude_id=eid):
            return False, "Another employee with the same belt no. / name exists in this station."
        des, ph = _s(designation).upper(), norm_phone(phone)
        c.execute("UPDATE employees SET name=?, belt_no=?, designation=?, contact_no=? WHERE id=?",
                  (name, belt, des, ph, eid))
        c.execute("UPDATE roster SET name=?, belt_no=?, designation=?, contact_no=? "
                  "WHERE emp_id=? AND roster_date>=?", (name, belt, des, ph, eid, today_iso()))
        audit(c, user, "EMP_EDIT", f"{name} ({belt})")
    return True, "Details updated (today's and upcoming rosters refreshed)."


def upcoming_count(eid, from_iso):
    with db() as c:
        return c.execute("SELECT COUNT(*) FROM roster WHERE emp_id=? AND roster_date>=?",
                         (int(eid), from_iso)).fetchone()[0]


def transfer_employee(eid, new_station, effective_iso, remove_old, user):
    with db() as c:
        row = c.execute("SELECT station, name, belt_no FROM employees WHERE id=?", (eid,)).fetchone()
        if not row:
            return False, "Employee not found."
        old, name, belt = row
        if old == new_station:
            return False, "Choose a different station."
        if _find_employee(c, new_station, belt, name):
            return False, f"A person with this belt no. / name already exists in {new_station.title()}."
        removed = 0
        if remove_old:
            removed = c.execute("DELETE FROM roster WHERE emp_id=? AND roster_date>=?",
                                (eid, effective_iso)).rowcount
        c.execute("UPDATE employees SET station=?, status='active', left_on=NULL, left_reason=NULL WHERE id=?",
                  (new_station, eid))
        c.execute("INSERT INTO transfers (emp_id,name,from_station,to_station,effective,by_user,ts) "
                  "VALUES (?,?,?,?,?,?,?)",
                  (eid, name, old, new_station, effective_iso, user, datetime.now().isoformat(timespec="seconds")))
        audit(c, user, "TRANSFER", f"{name}: {old} -> {new_station} (eff {effective_iso}, {removed} duties removed)")
    return True, f"{name} moved to {new_station.title()}. {removed} upcoming duty entr{'y' if removed == 1 else 'ies'} removed from {old.title()}."


def deactivate_employee(eid, reason, left_iso, remove_upcoming, user):
    with db() as c:
        row = c.execute("SELECT name FROM employees WHERE id=?", (eid,)).fetchone()
        if not row:
            return False, "Employee not found."
        removed = 0
        if remove_upcoming:
            removed = c.execute("DELETE FROM roster WHERE emp_id=? AND roster_date>=?", (eid, left_iso)).rowcount
        c.execute("UPDATE employees SET status='inactive', left_on=?, left_reason=? WHERE id=?",
                  (left_iso, reason, eid))
        audit(c, user, "EMP_REMOVE", f"{row[0]}: {reason} (eff {left_iso}, {removed} duties removed)")
    return True, f"{row[0]} removed from the active list ({removed} upcoming duty entries cleared)."


def reinstate_employee(eid, station, user):
    with db() as c:
        row = c.execute("SELECT name, belt_no FROM employees WHERE id=?", (eid,)).fetchone()
        if not row:
            return False, "Employee not found."
        if _find_employee(c, station, row[1], row[0], exclude_id=eid):
            return False, "A matching active employee already exists in that station."
        c.execute("UPDATE employees SET status='active', station=?, left_on=NULL, left_reason=NULL WHERE id=?",
                  (station, eid))
        audit(c, user, "EMP_REINSTATE", f"{row[0]} -> {station}")
    return True, f"{row[0]} reinstated in {station.title()}."


def import_employees(df, station, user):
    cols = {str(c).lower().strip(): c for c in df.columns}

    def pick(*keys):
        for k in keys:
            for lk, orig in cols.items():
                if k in lk:
                    return orig
        return None

    cn, cb, cd, cp = pick("name"), pick("belt", "badge"), pick("designation", "rank"), pick("contact", "phone", "mobile")
    if cn is None:
        return 0, 0, "Could not find a 'Name' column."
    added = skipped = 0
    for _, r in df.iterrows():
        ok, _msg = add_employee(station, r[cn], r[cb] if cb else "", r[cd] if cd else "", r[cp] if cp else "", user)
        added += ok
        skipped += (not ok)
    return added, skipped, ""


def restore_db(data):
    tmp = DB_FILE + ".restore"
    with open(tmp, "wb") as f:
        f.write(data)
    try:
        con = sqlite3.connect(tmp)
        tabs = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        con.close()
    except sqlite3.DatabaseError:
        os.remove(tmp)
        return False, "That file is not a valid database."
    if "roster" not in tabs:
        os.remove(tmp)
        return False, "That database has no roster table."
    os.replace(tmp, DB_FILE)
    init_db.clear()
    return True, "Database restored."




# =============================================================================
# DATA HEALTH CHECKS
# =============================================================================


def health_check(df):
    rows = []
    for _, r in df.iterrows():
        who, stn = r["name"] or "(no name)", r["station"]
        if r["contact_no"] and not PHONE_RE.match(r["contact_no"]):
            rows.append(("🟠 Invalid phone", who, stn, r["contact_no"]))
        if not r["contact_no"]:
            rows.append(("🟡 Missing phone", who, stn, ""))
        if not r["belt_no"]:
            rows.append(("🟡 Missing belt no.", who, stn, ""))
        if not r["beat_area"]:
            rows.append(("🟡 Missing beat area", who, stn, ""))
    phones = df[df["contact_no"] != ""].copy()
    phones["k"] = phones.apply(person_key, axis=1)
    for num, g in phones.groupby("contact_no"):
        if g["k"].nunique() > 1:
            rows.append(("🟠 Duplicate phone", ", ".join(sorted(set(g["name"]))),
                         ", ".join(sorted(set(g["station"]))), num))
    d = df.copy()
    d["k"] = d.apply(person_key, axis=1)
    for _, g in d[d["k"] != ""].groupby("k"):
        spans = []
        for r in g.itertuples():
            s, e = minutes(r.start_time), minutes(r.end_time)
            if s is None or e is None:
                continue
            spans.append((s, e + 1440 if e <= s else e, r))
        for i in range(len(spans)):
            for j in range(i + 1, len(spans)):
                a, b = spans[i], spans[j]
                if a[0] < b[1] and b[0] < a[1]:
                    rows.append((
                        "🔴 Double-booked", a[2].name, a[2].station,
                        f"{a[2].shift} {fmt12(a[2].start_time)}–{fmt12(a[2].end_time)} "
                        f"overlaps {b[2].shift} {fmt12(b[2].start_time)}–{fmt12(b[2].end_time)}"))
    return pd.DataFrame(rows, columns=["Issue", "Person", "Station", "Details"])



# =============================================================================
# REPORTS  (PDF + Excel)
# =============================================================================


def build_pdf(df, date_iso, scope_label, shift_rank, expected=()):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    from reportlab.platypus import (BaseDocTemplate, Frame, PageBreak, PageTemplate,
                                    Paragraph, Spacer, Table, TableStyle)

    NAVY, TEAL, MIST, ZEBRA, LINE = (colors.HexColor(x) for x in
                                     ("#0B2545", "#13A89E", "#DCE8F5", "#F4F7FB", "#CBD5E1"))
    d = datetime.strptime(date_iso, "%Y-%m-%d")
    long_date = d.strftime("%A, %d %B %Y")
    generated = datetime.now().strftime("%d %b %Y, %I:%M %p")

    class NumberedCanvas(canvas.Canvas):
        def __init__(self, *a, **k):
            super().__init__(*a, **k)
            self._pages = []

        def showPage(self):
            self._pages.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            total = len(self._pages)
            for state in self._pages:
                self.__dict__.update(state)
                self._decorate(total)
                super().showPage()
            super().save()

        def _decorate(self, total):
            w, h = landscape(A4)
            self.setFillColor(NAVY)
            self.rect(0, h - 21 * mm, w, 21 * mm, stroke=0, fill=1)
            self.setFillColor(TEAL)
            self.rect(0, h - 22.2 * mm, w, 1.2 * mm, stroke=0, fill=1)
            self.setFillColor(colors.white)
            self.setFont("Helvetica", 8.5)
            self.drawString(12 * mm, h - 8 * mm, ORG_LINE.upper())
            self.setFont("Helvetica-Bold", 17)
            self.drawString(12 * mm, h - 16 * mm, "DAILY DUTY ROSTER")
            self.setFont("Helvetica-Bold", 10.5)
            self.drawRightString(w - 12 * mm, h - 10 * mm, long_date)
            self.setFont("Helvetica", 8.5)
            self.drawRightString(w - 12 * mm, h - 16 * mm, scope_label)
            self.setStrokeColor(LINE)
            self.line(12 * mm, 11 * mm, w - 12 * mm, 11 * mm)
            self.setFillColor(colors.HexColor("#64748B"))
            self.setFont("Helvetica", 7.5)
            self.drawString(12 * mm, 7 * mm, f"Generated {generated}")
            self.drawCentredString(w / 2, 7 * mm, "Duty Roster Command Center")
            self.drawRightString(w - 12 * mm, 7 * mm, f"Page {self._pageNumber} of {total}")

    cell = ParagraphStyle("c", fontName="Helvetica", fontSize=7.8, leading=9.6, textColor=colors.HexColor("#0F172A"))
    cellb = ParagraphStyle("cb", parent=cell, fontName="Helvetica-Bold")
    head = ParagraphStyle("h", parent=cell, fontName="Helvetica-Bold", textColor=colors.white, fontSize=7.8)
    sub = ParagraphStyle("s", parent=cell, fontName="Helvetica-Bold", fontSize=8.2, textColor=NAVY)
    banner_l = ParagraphStyle("bl", fontName="Helvetica-Bold", fontSize=12, leading=15, textColor=NAVY)
    banner_r = ParagraphStyle("br", fontName="Helvetica", fontSize=8.5, leading=11, textColor=NAVY, alignment=2)
    sig = ParagraphStyle("sig", fontName="Helvetica", fontSize=8, alignment=1, textColor=colors.HexColor("#475569"))

    def P(t, st_=cell):
        return Paragraph(escape(_s(t)) or "&nbsp;", st_)

    buf = io.BytesIO()
    doc = BaseDocTemplate(buf, pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm,
                          topMargin=27 * mm, bottomMargin=15 * mm,
                          title=f"Duty Roster {long_date}", author=ORG_LINE)
    doc.addPageTemplates([PageTemplate(id="p", frames=[
        Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="f",
              leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)])])

    widths = [9, 40, 22, 33, 26, 24, 59, 34, 26]  # mm, total 273
    widths = [w * mm for w in widths]
    story = []
    present = set(df["station"])
    stations = [s for s in STATIONS if s in present or s in expected] + sorted(present - set(STATIONS))

    if not stations:
        story.append(Spacer(1, 40 * mm))
        story.append(Paragraph("No duty entries recorded for this date.", banner_l))

    for si, stn in enumerate(stations):
        sdf = df[df["station"] == stn]
        banner = Table([[Paragraph(escape(stn), banner_l),
                         Paragraph(f"{len(sdf)} duty entries &nbsp;|&nbsp; {sdf['name'].str.lower().nunique()} personnel" if len(sdf) else "Roster not submitted", banner_r)]],
                       colWidths=[doc.width * 0.6, doc.width * 0.4])
        banner.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#E6F4F2")),
            ("LINEBEFORE", (0, 0), (0, -1), 3, TEAL),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8)]))
        story += [banner, Spacer(1, 3 * mm)]
        if sdf.empty:
            story += [Spacer(1, 12 * mm), Paragraph("No duty roster has been submitted for this station on this date.", banner_l)]
            if si < len(stations) - 1:
                story.append(PageBreak())
            continue

        rows = [[P(h, head) for h in ("#", "Name", "Belt No.", "Designation", "Contact", "Squad / Team", "Beat Area / Location", "Timing", "Remarks")]]
        style = [
            ("BACKGROUND", (0, 0), (-1, 0), NAVY),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.4, LINE),
            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4)]
        n = 0
        sdf = sdf.assign(_r=sdf["shift"].map(shift_rank).fillna(99)).sort_values(
            ["_r", "start_time", "squad_team", "beat_area", "name"])
        for shift, g in sdf.groupby("_r", sort=True):
            label = g["shift"].iloc[0] or "Unassigned"
            first = g.iloc[0]
            pairs = set(zip(g["start_time"], g["end_time"]))
            span = f"{fmt12(first.start_time)} – {fmt12(first.end_time)}" if len(pairs) == 1 and first.start_time else ""
            rows.append([P(f"{label.upper()}   {span}   ({len(g)})", sub)] + [""] * 8)
            r_i = len(rows) - 1
            style += [("SPAN", (0, r_i), (-1, r_i)), ("BACKGROUND", (0, r_i), (-1, r_i), MIST)]
            for k, r in enumerate(g.itertuples()):
                n += 1
                rows.append([P(n), P(r.name, cellb), P(r.belt_no), P(r.designation), P(r.contact_no),
                             P(r.squad_team), P(r.beat_area), P(f"{fmt12(r.start_time)} – {fmt12(r.end_time)}"),
                             P(r.remarks)])
                if k % 2:
                    style.append(("BACKGROUND", (0, len(rows) - 1), (-1, len(rows) - 1), ZEBRA))
        t = Table(rows, colWidths=widths, repeatRows=1)
        t.setStyle(TableStyle(style))
        story.append(t)
        story.append(Spacer(1, 14 * mm))
        sigs = Table([[Paragraph("_______________________<br/>Prepared by", sig),
                       Paragraph("_______________________<br/>Verified by", sig),
                       Paragraph("_______________________<br/>Approved by", sig)]],
                     colWidths=[doc.width / 3] * 3)
        sigs.keepWithNext = False
        story.append(sigs)
        if si < len(stations) - 1:
            story.append(PageBreak())

    doc.build(story, canvasmaker=NumberedCanvas)
    return buf.getvalue()


def build_excel(df, date_iso):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    cols = {"station": "Station", "name": "Name", "belt_no": "Belt No.", "designation": "Designation",
            "contact_no": "Contact No.", "squad_team": "Squad / Team", "beat_area": "Beat Area",
            "shift": "Shift", "start_time": "Start", "end_time": "End", "remarks": "Remarks"}
    out = df[list(cols)].rename(columns=cols).copy()
    out["Start"] = out["Start"].map(fmt12)
    out["End"] = out["End"].map(fmt12)
    out.insert(0, "Sr", range(1, len(out) + 1))
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        groups = [("All Stations", out)] + [(s[:31], g) for s, g in out.groupby("Station", sort=False)]
        for title, g in groups:
            g = g.copy()
            g["Sr"] = range(1, len(g) + 1)
            g.to_excel(xw, index=False, sheet_name=title)
            ws = xw.sheets[title]
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="0B2545")
                c.alignment = Alignment(vertical="center")
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for i, col in enumerate(g.columns, 1):
                width = max([len(str(col))] + [len(str(v)) for v in g[col]]) + 2
                ws.column_dimensions[get_column_letter(i)].width = min(width, 55)
    return buf.getvalue()


@st.cache_data(show_spinner=False)
def cached_pdf(df, date_iso, scope_label, shift_rank, expected=()):
    return build_pdf(df, date_iso, scope_label, shift_rank, expected)


@st.cache_data(show_spinner=False)
def cached_excel(df, date_iso):
    return build_excel(df, date_iso)



# =============================================================================
# STYLING
# =============================================================================

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"], .stMarkdown, button, input, textarea { font-family: 'Inter', system-ui, sans-serif; }
.block-container { padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1500px; }
header[data-testid="stHeader"] { background: transparent; }

/* hero */
.hero { background: linear-gradient(120deg,#0B2545 0%,#13315C 55%,#0E6E74 100%); color:#fff;
  border-radius: 18px; padding: 22px 28px; display:flex; justify-content:space-between; align-items:center;
  box-shadow: 0 10px 30px rgba(11,37,69,.18); margin-bottom: 18px; }
.hero-t { font-size: 1.55rem; font-weight: 700; letter-spacing:-.01em; }
.hero-s { opacity:.85; font-size:.95rem; margin-top:2px; }
.pill { display:inline-block; background:rgba(255,255,255,.14); border:1px solid rgba(255,255,255,.25);
  padding:5px 12px; border-radius:999px; font-size:.8rem; font-weight:500; margin-left:6px; }

/* KPI cards */
.cards { display:grid; grid-template-columns: repeat(auto-fit,minmax(170px,1fr)); gap:14px; margin-bottom: 18px; }
.kpi { background:#fff; border:1px solid #E3E9F2; border-radius:14px; padding:14px 18px;
  box-shadow: 0 2px 10px rgba(15,23,42,.04); border-top:3px solid #13A89E; }
.kpi.warn { border-top-color:#F2B134; } .kpi.bad { border-top-color:#E5484D; } .kpi.ok { border-top-color:#2FB170; }
.kpi .k { font-size:.72rem; text-transform:uppercase; letter-spacing:.06em; color:#64748B; font-weight:600; }
.kpi .v { font-size:1.8rem; font-weight:700; color:#0B2545; line-height:1.15; }
.kpi .s { font-size:.78rem; color:#64748B; }

/* tabs */
.stTabs [data-baseweb="tab-list"] { gap: 6px; border-bottom: 1px solid #E3E9F2; }
.stTabs [data-baseweb="tab"] { height: 44px; padding: 0 18px; border-radius: 10px 10px 0 0; font-weight:600; }
.stTabs [aria-selected="true"] { color:#0E6E74; }

/* buttons */
.stButton > button, .stDownloadButton > button { border-radius: 10px; font-weight:600; padding: .5rem 1.1rem; }
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
  background: linear-gradient(135deg,#13A89E,#0E7C86); border:0; color:#fff; box-shadow: 0 4px 14px rgba(19,168,158,.35); }

/* sidebar */
[data-testid="stSidebar"] { background: linear-gradient(180deg,#0B2545 0%,#10294D 100%); }
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p, [data-testid="stSidebar"] .side-t,
[data-testid="stSidebar"] .side-s { color:#DCE6F5 !important; }
.side-t { font-size:1.15rem; font-weight:700; color:#fff; margin-bottom:2px; }
.side-s { font-size:.78rem; opacity:.75; }
.user-chip { background:rgba(255,255,255,.08); border:1px solid rgba(255,255,255,.18); color:#fff;
  padding:10px 14px; border-radius:12px; font-size:.85rem; margin:10px 0 14px; }
[data-testid="stSidebar"] .stButton > button { background:rgba(255,255,255,.08); color:#fff;
  border:1px solid rgba(255,255,255,.25); box-shadow:none; }
[data-testid="stSidebar"] .stButton > button p { color:#fff; }
[data-testid="stSidebar"] .stButton > button:hover { background:rgba(255,255,255,.18); }

.saveline { display:flex; align-items:center; gap:10px; font-size:.85rem; color:#64748B; }
.dot { width:9px; height:9px; border-radius:50%; background:#2FB170; display:inline-block; }
.dot.dirty { background:#F2B134; box-shadow:0 0 0 4px rgba(242,177,52,.25); }
.login-card { background:#fff; border:1px solid #E3E9F2; border-radius:18px; padding:28px; box-shadow:0 12px 40px rgba(11,37,69,.10); }
/* sidebar navigation */
[data-testid="stSidebar"] [data-testid="stRadio"] label { background: rgba(255,255,255,.06); border-radius: 10px;
  padding: 8px 12px; margin-bottom: 4px; width: 100%; }
[data-testid="stSidebar"] [data-testid="stRadio"] label p { color:#E8EEF7 !important; font-weight:600; }
[data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) { background: rgba(19,168,158,.38); }
[data-testid="stSidebar"] [data-testid="stRadio"] label > div:first-child { display:none; }
[data-testid="stSidebar"] [data-testid="stDateInput"] input { color:#0B2545; }
</style>
"""


def cards(items):
    html = '<div class="cards">' + "".join(
        f'<div class="kpi {cls}"><div class="k">{k}</div><div class="v">{v}</div><div class="s">{s}</div></div>'
        for k, v, s, cls in items) + "</div>"
    st.markdown(html, unsafe_allow_html=True)


# =============================================================================
# AUTH
# =============================================================================


def login_screen():
    st.markdown(
        '<div class="hero"><div><div class="hero-t">🛡️ Duty Roster Command Center</div>'
        '<div class="hero-s">Sign in with your station account to manage today\'s duty roster</div></div></div>',
        unsafe_allow_html=True)
    _, mid, _ = st.columns([1, 1.3, 1])
    with mid:
        st.markdown('<div class="login-card">', unsafe_allow_html=True)
        locked = st.session_state.get("locked_until", 0) - _time.time()
        with st.form("login"):
            who = st.selectbox("Station / Role", [ADMIN] + STATIONS)
            pw = st.text_input("Password", type="password")
            go = st.form_submit_button("Sign in", type="primary", width="stretch")
        if go:
            if locked > 0:
                st.error(f"Too many attempts. Try again in {int(locked)}s.")
            elif verify_password(who, pw):
                st.session_state.update(logged_in=True, station=who, fails=0)
                with db() as c:
                    audit(c, who, "LOGIN")
                st.rerun()
            else:
                st.session_state.fails = st.session_state.get("fails", 0) + 1
                if st.session_state.fails >= 5:
                    st.session_state.locked_until = _time.time() + 60
                    st.session_state.fails = 0
                st.error("Incorrect password.")
        st.markdown("</div>", unsafe_allow_html=True)




# =============================================================================
# UI HELPERS
# =============================================================================

PAGES = ["🏠 Dashboard", "🧭 Roster Builder", "👥 Employees", "📄 Reports", "⚙️ Settings"]
P_DASH, P_BUILD, P_EMP, P_REP, P_SET = PAGES
STEP1, STEP2, STEP3 = "① Start", "② Assign duties", "③ Review & print"


def flash(msg, icon="✅"):
    st.session_state.setdefault("_flash", []).append((msg, icon))


def bump():
    st.session_state.ver = st.session_state.get("ver", 0) + 1


def goto_page(p):
    st.session_state["_goto_page"] = p


def goto_step(s):
    st.session_state["_goto_step"] = s


def nice(station):
    return station.title()


def fname_for(stations, date_iso):
    tag = "ALL_STATIONS" if set(stations) == set(STATIONS) else "_".join(s.replace(" ", "") for s in stations)
    return f"Duty_Roster_{tag}_{date_iso}"


def timing(r):
    return f"{fmt12(r.start_time)} – {fmt12(r.end_time)}"


def emp_label(r):
    return f"{r.name}  ·  {r.belt_no or '—'}  ·  {r.designation or '—'}"


def shift_rank_map(shift_df):
    return {n: i for i, n in enumerate(shift_df["name"])}


def pdf_for(df, date_iso, scope, shift_df, expected=()):
    label = "All Stations" if set(scope) == set(STATIONS) else ", ".join(nice(s) for s in scope)
    return cached_pdf(df, date_iso, label, shift_rank_map(shift_df), tuple(expected))


# =============================================================================
# PAGE: DASHBOARD
# =============================================================================


def page_dashboard(sel, sel_iso, is_admin, user, work, shift_df):
    scope = STATIONS if is_admin else [user]
    day = load_day(sel_iso, scope)
    prev_iso = (sel - timedelta(days=1)).strftime("%Y-%m-%d")
    prev = load_day(prev_iso, scope)
    emps = load_employees(scope)

    rows, eligible = [], []
    for s in scope:
        d, e = day[day.station == s], emps[emps.station == s]
        on = int(d["emp_id"].nunique())
        if d.empty and not prev[prev.station == s].empty:
            eligible.append(s)
        rows.append({"Station": nice(s), "Active staff": len(e), "On duty": on,
                     "Not assigned": max(len(e) - on, 0), "Duty entries": len(d),
                     "Status": "✅ Roster ready" if len(d) else "⚠️ Not created yet"})
    tbl = pd.DataFrame(rows)
    issues = health_check(day) if len(day) else pd.DataFrame(columns=["Issue"])
    dbl = int(issues["Issue"].str.startswith("🔴").sum()) if len(issues) else 0
    ready = int((tbl["Duty entries"] > 0).sum())

    cards([
        ("Rosters ready", f"{ready}/{len(scope)}", "stations with a roster", "ok" if ready == len(scope) else "warn"),
        ("On duty", int(tbl["On duty"].sum()), "personnel today", ""),
        ("Not assigned", int(tbl["Not assigned"].sum()), "active staff without duty", "warn" if tbl["Not assigned"].sum() else "ok"),
        ("Double-bookings", dbl, "overlapping duties" if dbl else "none detected", "bad" if dbl else "ok"),
    ])

    left, right = st.columns([2.2, 1])
    with left:
        st.markdown("#### 📍 Station status")
        st.dataframe(tbl, hide_index=True, width="stretch")
        a, b, c = st.columns(3)
        pdf = pdf_for(day, sel_iso, scope, shift_df, expected=scope) if len(day) else b""
        a.download_button("📄  Download PDF" + (" — all stations" if is_admin else ""), pdf,
                          f"{fname_for(scope, sel_iso)}.pdf", "application/pdf", type="primary",
                          width="stretch", disabled=day.empty, on_click="ignore", key="dash_pdf")
        if b.button(f"⚡  Auto-create from yesterday ({len(eligible)})", width="stretch",
                    disabled=not eligible,
                    help="For every station that has no roster today, copy yesterday's roster."):
            total, skip = 0, 0
            for s in eligible:
                n, sk = copy_rows(prev[prev.station == s], s, sel_iso, "Add to existing", user, f"yesterday ({prev_iso})")
                total, skip = total + n, skip + len(sk)
            flash(f"Created rosters for {len(eligible)} station(s): {total} duties copied"
                  + (f", {skip} skipped (staff moved/removed)." if skip else "."))
            st.rerun()
        if c.button("🧭  Open Roster Builder", width="stretch"):
            goto_page(P_BUILD)
            st.rerun()
        if not day.empty and len(issues):
            with st.expander(f"🩺 Data checks — {len(issues)} item(s)"):
                st.dataframe(issues, hide_index=True, width="stretch")
    with right:
        st.markdown(f"#### ✅ Getting started — {nice(work)}")
        n_emp = int((emps.station == work).sum())
        n_beat = len(get_list("beat", work))
        n_duty = int((day.station == work).sum())
        steps = [
            (n_emp > 0, "Employees added", f"{n_emp} active", P_EMP),
            (n_beat > 0, "Beats defined", f"{n_beat} beats", P_SET),
            (n_duty > 0, "Roster created for this date", f"{n_duty} duties", P_BUILD),
        ]
        with st.container(border=True):
            for ok, title, sub, page in steps:
                st.markdown(f"{'✅' if ok else '⬜'} **{title}**  \n<span style='color:#64748B;font-size:.8rem'>{sub}</span>",
                            unsafe_allow_html=True)
            nxt = next((s for s in steps if not s[0]), None)
            if nxt and st.button(f"Next: {nxt[1]} →", width="stretch", key="dash_next"):
                goto_page(nxt[3])
                st.rerun()
            if not nxt:
                st.success("All set — print the PDF from Reports.")


# =============================================================================
# PAGE: ROSTER BUILDER
# =============================================================================


def do_copy(src_df, label, work, sel_iso, mode, user):
    copied, skipped = copy_rows(src_df, work, sel_iso, mode, user, label)
    msg = f"Copied {copied} duties from {label}."
    if skipped:
        msg += f" Skipped {len(skipped)} no longer at this station: " + ", ".join(skipped[:5]) + ("…" if len(skipped) > 5 else "")
    flash(msg, "✅" if copied else "⚠️")
    goto_step(STEP3 if copied else STEP1)
    bump()
    st.rerun()


def step_start(sel, sel_iso, work, user, day):
    n = len(day)
    st.markdown("#### Start from a previous roster — or from scratch")
    mode = "Add to existing"
    confirmed = True
    if n:
        st.info(f"**{n}** duties already exist for {sel:%A, %d %b %Y}.")
        mode = st.radio("When copying", ["Add to existing", "Replace existing"], horizontal=True, key="cp_mode")
        if mode == "Replace existing":
            confirmed = st.checkbox("I understand existing entries for this station and date will be replaced", key="cp_ok")

    y_iso = (sel - timedelta(days=1)).strftime("%Y-%m-%d")
    w_iso = (sel - timedelta(days=7)).strftime("%Y-%m-%d")
    y_df, w_df = load_day(y_iso, [work]), load_day(w_iso, [work])

    c1, c2 = st.columns(2)
    with c1.container(border=True):
        st.markdown("**📅 Yesterday's roster**")
        st.caption(f"{sel - timedelta(days=1):%A, %d %b} — {len(y_df)} duties")
        if st.button("Use yesterday's roster", key="cp_y", type="primary", disabled=y_df.empty or not confirmed,
                     width="stretch"):
            do_copy(y_df, f"yesterday ({y_iso})", work, sel_iso, mode, user)
    with c2.container(border=True):
        st.markdown("**🗓️ Same day last week**")
        st.caption(f"{sel - timedelta(days=7):%A, %d %b} — {len(w_df)} duties")
        if st.button("Use last week's roster", key="cp_w", disabled=w_df.empty or not confirmed, width="stretch"):
            do_copy(w_df, f"last week ({w_iso})", work, sel_iso, mode, user)

    c3, c4 = st.columns(2)
    with c3.container(border=True):
        st.markdown("**🔎 Any other date**")
        other = st.date_input("Pick a date", value=sel - timedelta(days=2), format="DD/MM/YYYY",
                              key=f"cp_other_{sel_iso}", label_visibility="collapsed")
        o_iso = other.strftime("%Y-%m-%d")
        o_df = load_day(o_iso, [work])
        st.caption(f"{len(o_df)} duties on {other:%d %b %Y}")
        if st.button("Copy from this date", key="cp_o", disabled=o_df.empty or not confirmed, width="stretch"):
            do_copy(o_df, o_iso, work, sel_iso, mode, user)
    with c4.container(border=True):
        st.markdown("**📋 Saved template**")
        tpl = list_templates(work)
        if tpl.empty:
            st.caption("No templates yet. Build a roster, then save it as a template in step ③.")
        else:
            names = {int(r.id): f"{r.name}  ({r.n} duties)" for r in tpl.itertuples()}
            tid = st.selectbox("Template", list(names), format_func=names.get, key="tpl_pick",
                               label_visibility="collapsed")
            t1, t2 = st.columns([3, 1])
            if t1.button("Apply template", key="tpl_apply", disabled=not confirmed, width="stretch"):
                do_copy(load_template(tid), f"template '{tpl.set_index('id').loc[tid, 'name']}'",
                        work, sel_iso, mode, user)
            if t2.button("🗑", key="tpl_del", help="Delete this template", width="stretch"):
                delete_template(tid, user)
                flash("Template deleted.", "🗑️")
                st.rerun()

    st.markdown("&nbsp;")
    if st.button("✍️  Start blank — I'll assign duties myself  →", key="start_blank"):
        goto_step(STEP2)
        st.rerun()


def step_assign(sel_iso, work, user, day, emps, beats, squads, shift_df):
    ver = st.session_state.get("ver", 0)
    presets = {r.name: (r.start_time, r.end_time) for r in shift_df.itertuples()}
    shifts = list(shift_df["name"])
    assigned = {}
    for r in day.itertuples():
        if pd.notna(r.emp_id):
            assigned.setdefault(int(r.emp_id), []).append(r.shift or "?")

    if emps.empty:
        st.warning("No employees in this station yet.")
        if st.button("👥  Add employees first", type="primary"):
            goto_page(P_EMP)
            st.rerun()
        return

    st.markdown("#### Assign officers to a beat")
    st.caption("Select one or more employees, choose the beat and shift once, then press **Assign**. "
               "Repeat for every beat. ✓ marks people who already have a duty on this date.")
    labels = {}
    for r in emps.itertuples():
        tag = f"     ✓ {' + '.join(dict.fromkeys(assigned[r.id]))}" if r.id in assigned else ""
        labels[int(r.id)] = emp_label(r) + tag
    free_ids = [i for i in labels if i not in assigned]
    key_p = f"as_people_{ver}"

    def pick_free():
        st.session_state[key_p] = free_ids

    with st.container(border=True):
        l, r_ = st.columns([1.15, 1])
        with l:
            people = st.multiselect("1 · Employees", list(labels), format_func=labels.get, key=key_p,
                                    placeholder="Search by name or belt no. and select…")
            st.button(f"Select everyone not yet assigned ({len(free_ids)})", on_click=pick_free,
                      disabled=not free_ids, key=f"as_all_{ver}")
        with r_:
            if beats:
                beat = st.selectbox("2 · Beat / location", beats, index=None,
                                    placeholder="Choose a beat…", key=f"as_beat_{ver}")
            else:
                beat = None
                st.warning("No beats defined yet for this station.")
            with st.expander("➕ Add a new beat"):
                nb = st.text_input("New beat name", key=f"as_nb_{ver}")
                if st.button("Add beat", key=f"as_nbb_{ver}"):
                    if add_to_list("beat", work, [nb], user):
                        flash(f"Beat '{nb.strip()}' added.")
                        st.rerun()
                    else:
                        st.error("Enter a new, non-duplicate beat name.")
            sq = st.selectbox("Squad / team (optional)", ["— none —"] + squads, key=f"as_sq_{ver}")
            shift = st.selectbox("3 · Shift", shifts, key=f"as_sh_{ver}")
            ps, pe = presets.get(shift, ("09:00", "17:00"))
            t1, t2 = st.columns(2)
            start = t1.time_input("Start", value=to_time(ps), step=300, key=f"as_s_{ver}_{shift}")
            end = t2.time_input("End", value=to_time(pe), step=300, key=f"as_e_{ver}_{shift}")
            remarks = st.text_input("Remarks (optional)", key=f"as_rm_{ver}")

        if st.button("➕  Assign selected employees", type="primary", key=f"as_go_{ver}"):
            if not people:
                st.error("Select at least one employee.")
            elif not beat:
                st.error("Choose a beat.")
            else:
                added, skipped = assign_duty(sel_iso, people, beat, "" if sq.startswith("—") else sq, shift,
                                             start.strftime("%H:%M"), end.strftime("%H:%M"), remarks.strip(), user)
                flash(f"Assigned {added} to {beat} ({shift})." + (f" {skipped} already had this duty." if skipped else ""))
                bump()
                st.rerun()

    st.markdown("&nbsp;")
    a, b = st.columns([1, 1])
    with a.expander(f"🕘 Not yet assigned today ({len(free_ids)})"):
        if free_ids:
            st.dataframe(emps[emps.id.isin(free_ids)][["name", "belt_no", "designation"]]
                         .rename(columns={"name": "Name", "belt_no": "Belt No.", "designation": "Designation"}),
                         hide_index=True, width="stretch")
        else:
            st.success("Everyone has a duty today.")
    if b.button(f"Review & print  →   ({len(day)} duties so far)", key="to_review", disabled=day.empty):
        goto_step(STEP3)
        st.rerun()


def step_review(sel, sel_iso, work, user, day, beats, squads, shift_df):
    if day.empty:
        st.info("Nothing to review yet — start from a previous roster (①) or assign duties (②).")
        return
    presets = {r.name: (r.start_time, r.end_time) for r in shift_df.itertuples()}
    fname = fname_for([work], sel_iso)
    pdf = pdf_for(day, sel_iso, [work], shift_df)
    xls = cached_excel(day, sel_iso)

    c1, c2, c3 = st.columns([1.3, 1.3, 3])
    c1.download_button("📄  Download PDF", pdf, f"{fname}.pdf", "application/pdf", type="primary",
                       width="stretch", on_click="ignore", key="rv_pdf")
    c2.download_button("📊  Excel", xls, f"{fname}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       width="stretch", on_click="ignore", key="rv_xls")
    st.caption("Change a beat, shift, timing or remark right in the table, or tick 🗑 to take someone off duty — "
               "then press **Save changes**.")

    ed = pd.DataFrame({"id": day["id"], "Remove": False, "name": day["name"], "belt_no": day["belt_no"],
                       "designation": day["designation"], "beat_area": day["beat_area"],
                       "squad_team": day["squad_team"], "shift": day["shift"],
                       "start_time": day["start_time"].map(to_time), "end_time": day["end_time"].map(to_time),
                       "remarks": day["remarks"]})
    beat_opts = beats + sorted(set(day["beat_area"]) - set(beats) - {""})
    squad_opts = squads + sorted(set(day["squad_team"]) - set(squads) - {""})
    shift_opts = list(shift_df["name"]) + sorted(set(day["shift"]) - set(shift_df["name"]) - {""})
    cfg = {
        "Remove": st.column_config.CheckboxColumn("🗑", width="small", help="Tick to remove from this roster"),
        "name": st.column_config.TextColumn("Name", disabled=True, width="medium"),
        "belt_no": st.column_config.TextColumn("Belt No.", disabled=True, width="small"),
        "designation": st.column_config.TextColumn("Designation", disabled=True, width="medium"),
        "beat_area": st.column_config.SelectboxColumn("Beat / Location", options=beat_opts, width="large"),
        "squad_team": st.column_config.SelectboxColumn("Squad", options=squad_opts, width="small"),
        "shift": st.column_config.SelectboxColumn("Shift", options=shift_opts, width="small"),
        "start_time": st.column_config.TimeColumn("Start", format="hh:mm a", step=60, width="small"),
        "end_time": st.column_config.TimeColumn("End", format="hh:mm a", step=60, width="small"),
        "remarks": st.column_config.TextColumn("Remarks", width="medium"),
    }
    key = f"rv_{sel_iso}_{work}_{st.session_state.get('ver', 0)}"
    edited = st.data_editor(ed, key=key, column_config=cfg, hide_index=True, num_rows="fixed", width="stretch",
                            column_order=["Remove", "name", "belt_no", "designation", "beat_area", "squad_team",
                                          "shift", "start_time", "end_time", "remarks"],
                            height=int(min(680, 44 + 35 * (len(ed) + 1))))
    state = st.session_state.get(key, {})
    pending = len(state.get("edited_rows", {}))
    b1, b2, b3 = st.columns([1.3, 1.1, 5])
    if b1.button("💾  Save changes", type="primary", disabled=pending == 0, width="stretch", key="rv_save"):
        stat = apply_roster_edits(day, edited, user, presets)
        flash(f"Saved — {stat['updated']} updated, {stat['removed']} removed.")
        bump()
        st.rerun()
    if b2.button("↺  Discard", disabled=pending == 0, width="stretch", key="rv_disc"):
        bump()
        st.rerun()
    if pending:
        b3.markdown(f'<div class="saveline" style="padding-top:8px"><span class="dot dirty"></span>'
                    f'{pending} unsaved change(s)</div>', unsafe_allow_html=True)
    else:
        b3.markdown('<div class="saveline" style="padding-top:8px"><span class="dot"></span>'
                    'All changes saved</div>', unsafe_allow_html=True)

    st.divider()
    t1, t2 = st.columns([2, 1])
    with t1.container(border=True):
        st.markdown("**📋 Save this roster as a template** — reuse it any day in step ①")
        tn = st.text_input("Template name", value=f"{sel:%A} roster", key=f"tpl_name_{sel_iso}")
        if st.button("Save as template", key="tpl_save", disabled=not tn.strip()):
            n = save_template(work, tn, day, user)
            flash(f"Template '{tn.strip()}' saved with {n} duties.", "📋")
            st.rerun()
    with t2.container(border=True):
        st.markdown("**🧨 Clear this roster**")
        sure = st.checkbox(f"Delete all {len(day)} duties on {sel:%d %b %Y}", key="clr_ok")
        if st.button("Clear roster", disabled=not sure, key="clr_go"):
            n = clear_day(work, sel_iso, user)
            flash(f"Cleared {n} duties.", "🧹")
            bump()
            goto_step(STEP1)
            st.rerun()

    issues = health_check(day)
    if len(issues):
        with st.expander(f"🩺 Data checks — {len(issues)} item(s) to review"):
            st.dataframe(issues, hide_index=True, width="stretch")


def page_builder(sel, sel_iso, work, user, is_admin, shift_df):
    day = load_day(sel_iso, [work])
    emps = load_employees([work])
    beats, squads = get_list("beat", work), get_list("squad", work)

    st.markdown(f"### 🧭 Roster for **{nice(work)}** — {sel:%A, %d %B %Y}")
    if "_goto_step" in st.session_state:
        st.session_state["bstep"] = st.session_state.pop("_goto_step")
    if st.session_state.get("bstep") not in (STEP1, STEP2, STEP3):
        st.session_state["bstep"] = STEP3 if len(day) else STEP1
    step = st.radio("Step", [STEP1, STEP2, STEP3], horizontal=True, key="bstep", label_visibility="collapsed")
    st.caption(f"{len(day)} duties · {int(day['emp_id'].nunique())} personnel · {len(emps)} active employees in this station")

    if step == STEP1:
        step_start(sel, sel_iso, work, user, day)
    elif step == STEP2:
        step_assign(sel_iso, work, user, day, emps, beats, squads, shift_df)
    else:
        step_review(sel, sel_iso, work, user, day, beats, squads, shift_df)


# =============================================================================
# PAGE: EMPLOYEES
# =============================================================================


def page_employees(user, is_admin, work, designations):
    show_all = False
    if is_admin:
        show_all = st.checkbox("Show all stations", key="emp_all")
    view_stations = STATIONS if show_all else [work]
    status = st.radio("Show", ["Active", "Inactive", "All"], horizontal=True, key="emp_status")
    emps = load_employees(view_stations, status)
    ver = st.session_state.get("ver", 0)

    active_n = int((emps.status == "active").sum())
    cards([("Employees", len(emps), f"{status.lower()} · {'all stations' if show_all else nice(work)}", ""),
           ("Active", active_n, "available for duty", "ok")])

    with st.expander("➕ Add a new employee", expanded=emps.empty):
        c1, c2, c3, c4 = st.columns([2, 1.2, 1.6, 1.4])
        n = c1.text_input("Full name", key=f"ne_n_{ver}")
        b = c2.text_input("Belt No.", key=f"ne_b_{ver}", placeholder="e.g. SGT-2613")
        d = c3.selectbox("Designation", [""] + designations, key=f"ne_d_{ver}")
        p = c4.text_input("Contact", key=f"ne_p_{ver}", placeholder="03XX-XXXXXXX")
        station = work
        if is_admin:
            station = st.selectbox("Station", STATIONS, index=STATIONS.index(work), key=f"ne_s_{ver}")
        if st.button("Add employee", type="primary", key=f"ne_go_{ver}"):
            ok, msg = add_employee(station, n, b, d, p, user)
            if ok:
                flash(msg)
                bump()
                st.rerun()
            else:
                st.error(msg)
        if p and not PHONE_RE.match(norm_phone(p)):
            st.caption("⚠️ Phone should look like 03XX-XXXXXXX")

    with st.expander("📥 Import many employees from Excel / CSV"):
        st.caption("Columns needed: **Name**; optional: Belt No., Designation, Contact. Duplicates are skipped.")
        up = st.file_uploader("Upload .xlsx or .csv", type=["xlsx", "csv"], key=f"imp_{ver}")
        if up is not None:
            try:
                idf = pd.read_csv(up) if up.name.lower().endswith(".csv") else pd.read_excel(up)
                st.dataframe(idf.head(8), hide_index=True, width="stretch")
                tgt = work
                if is_admin:
                    tgt = st.selectbox("Import into station", STATIONS, index=STATIONS.index(work), key=f"imp_s_{ver}")
                if st.button(f"Import {len(idf)} rows into {nice(tgt)}", key=f"imp_go_{ver}"):
                    a, s_, err = import_employees(idf, tgt, user)
                    if err:
                        st.error(err)
                    else:
                        flash(f"Imported {a} employees, skipped {s_} duplicates/invalid.")
                        bump()
                        st.rerun()
            except Exception as e:  # noqa: BLE001
                st.error(f"Could not read the file: {e}")

    q = st.text_input("Search", placeholder="🔎  Search name, belt no., designation…", label_visibility="collapsed", key="emp_q")
    lst = emps
    if q:
        m = lst[["name", "belt_no", "designation", "contact_no"]].apply(
            lambda col: col.str.lower().str.contains(q.lower(), regex=False)).any(axis=1)
        lst = lst[m]
    show = lst[["station", "name", "belt_no", "designation", "contact_no", "status", "left_on", "left_reason"]].copy()
    show["station"] = show["station"].map(nice)
    show.columns = ["Station", "Name", "Belt No.", "Designation", "Contact", "Status", "Left on", "Reason"]
    if not show_all:
        show = show.drop(columns="Station")
    if status == "Active":
        show = show.drop(columns=["Left on", "Reason", "Status"])
    st.dataframe(show, hide_index=True, width="stretch", height=min(520, 60 + 35 * len(show)))

    if lst.empty:
        return
    st.markdown("#### 🛠️ Manage an employee")
    opts = {int(r.id): f"{r.name}  ·  {r.belt_no or '—'}  ·  {nice(r.station)}" + ("" if r.status == "active" else "  (inactive)")
            for r in lst.itertuples()}
    eid = st.selectbox("Choose employee", list(opts), format_func=opts.get, index=None,
                       placeholder="Select to edit, transfer or remove…", key=f"emp_pick_{ver}")
    if eid is None:
        return
    row = lst[lst.id == eid].iloc[0]
    is_active = row["status"] == "active"
    actions = ["✏️ Edit details", "🔁 Transfer to another station", "🚪 Remove"] if is_active else ["↩️ Reinstate", "✏️ Edit details"]
    act = st.radio("Action", actions, horizontal=True, key=f"emp_act_{ver}_{eid}")

    with st.container(border=True):
        if act.startswith("✏️"):
            c1, c2, c3, c4 = st.columns([2, 1.2, 1.6, 1.4])
            n = c1.text_input("Full name", value=row["name"], key=f"ed_n_{eid}")
            b = c2.text_input("Belt No.", value=row["belt_no"], key=f"ed_b_{eid}")
            dopts = designations if row["designation"] in designations or not row["designation"] else designations + [row["designation"]]
            d = c3.selectbox("Designation", [""] + dopts, index=([""] + dopts).index(row["designation"]), key=f"ed_d_{eid}")
            p = c4.text_input("Contact", value=row["contact_no"], key=f"ed_p_{eid}")
            if st.button("Save details", type="primary", key=f"ed_go_{eid}"):
                ok, msg = edit_employee(eid, n, b, d, p, user)
                if ok:
                    flash(msg)
                    bump()
                    st.rerun()
                else:
                    st.error(msg)
        elif act.startswith("🔁"):
            st.caption(f"Currently at **{nice(row['station'])}**. Beat and squad assignments do not carry over — "
                       "assign them again at the new station.")
            c1, c2 = st.columns(2)
            to = c1.selectbox("Move to station", [s for s in STATIONS if s != row["station"]], format_func=nice, key=f"tr_s_{eid}")
            eff = c2.date_input("Effective from", value=date.today(), format="DD/MM/YYYY", key=f"tr_d_{eid}")
            eff_iso = eff.strftime("%Y-%m-%d")
            cnt = upcoming_count(eid, eff_iso)
            rm = st.checkbox(f"Remove from {nice(row['station'])}'s rosters from that date ({cnt} duty entries)", value=True, key=f"tr_rm_{eid}")
            if st.button("Transfer employee", type="primary", key=f"tr_go_{eid}"):
                ok, msg = transfer_employee(eid, to, eff_iso, rm, user)
                if ok:
                    flash(msg, "🔁")
                    bump()
                    st.rerun()
                else:
                    st.error(msg)
        elif act.startswith("🚪"):
            c1, c2 = st.columns(2)
            reason = c1.selectbox("Reason", REMOVAL_REASONS, key=f"rm_r_{eid}")
            left = c2.date_input("Last day of service here", value=date.today(), format="DD/MM/YYYY", key=f"rm_d_{eid}")
            left_iso = left.strftime("%Y-%m-%d")
            cnt = upcoming_count(eid, left_iso)
            rm = st.checkbox(f"Also remove from rosters from that date ({cnt} duty entries)", value=True, key=f"rm_u_{eid}")
            sure = st.checkbox("Yes, remove this employee from the active list (history is kept)", key=f"rm_ok_{eid}")
            if st.button("Remove employee", disabled=not sure, key=f"rm_go_{eid}"):
                ok, msg = deactivate_employee(eid, reason, left_iso, rm, user)
                flash(msg, "🚪")
                bump()
                st.rerun()
            st.caption("Moving to another station? Use **Transfer** instead — it keeps them active.")
        else:  # reinstate
            st.caption(f"Removed on {row['left_on'] or '—'} · {row['left_reason'] or '—'}")
            to = st.selectbox("Reinstate in station", STATIONS if is_admin else [work],
                              index=(STATIONS if is_admin else [work]).index(row["station"]) if row["station"] in (STATIONS if is_admin else [work]) else 0,
                              format_func=nice, key=f"ri_s_{eid}")
            if st.button("Reinstate employee", type="primary", key=f"ri_go_{eid}"):
                ok, msg = reinstate_employee(eid, to, user)
                if ok:
                    flash(msg, "↩️")
                    bump()
                    st.rerun()
                else:
                    st.error(msg)

    with db() as c:
        hist = pd.read_sql_query(
            "SELECT effective AS 'Effective', from_station AS 'From', to_station AS 'To', by_user AS 'By' "
            "FROM transfers WHERE emp_id=? ORDER BY id DESC", c, params=[int(eid)])
    if len(hist):
        with st.expander("🔁 Transfer history"):
            st.dataframe(hist, hide_index=True, width="stretch")


# =============================================================================
# PAGE: REPORTS
# =============================================================================


def page_reports(sel, sel_iso, is_admin, user, shift_df):
    scope = STATIONS if is_admin else [user]
    day = load_day(sel_iso, scope)
    st.markdown(f"### 📄 Duty roster reports — {sel:%A, %d %B %Y}")
    if is_admin:
        st.markdown("#### One click — all stations")
        a, b, _ = st.columns([1.4, 1.2, 3])
        a.download_button("📄  Download ALL STATIONS PDF",
                          pdf_for(day, sel_iso, STATIONS, shift_df, expected=STATIONS) if len(day) else b"",
                          f"{fname_for(STATIONS, sel_iso)}.pdf", "application/pdf", type="primary",
                          width="stretch", disabled=day.empty, on_click="ignore", key="rp_all")
        b.download_button("📊  All stations Excel", cached_excel(day, sel_iso) if len(day) else b"",
                          f"{fname_for(STATIONS, sel_iso)}.xlsx",
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                          width="stretch", disabled=day.empty, on_click="ignore", key="rp_all_x")
        st.caption("Stations without a roster are listed in the PDF as “not submitted”.")
        st.markdown("#### Individual stations")
    cols = st.columns(len(scope)) if len(scope) <= 5 else st.columns(5)
    for i, s in enumerate(scope):
        d = day[day.station == s]
        with cols[i % len(cols)].container(border=True):
            st.markdown(f"**{nice(s)}**")
            st.caption(f"{len(d)} duties" if len(d) else "No roster yet")
            st.download_button("📄 PDF", pdf_for(d, sel_iso, [s], shift_df) if len(d) else b"",
                               f"{fname_for([s], sel_iso)}.pdf", "application/pdf", width="stretch",
                               disabled=d.empty, on_click="ignore", key=f"rp_{s}")
    if day.empty:
        st.info("No duty entries on this date.")
        return
    st.markdown("#### Preview")
    prev = day.assign(Timing=day.apply(timing, axis=1))
    prev["station"] = prev["station"].map(nice)
    st.dataframe(prev[["station", "name", "belt_no", "designation", "contact_no", "squad_team", "beat_area",
                       "shift", "Timing", "remarks"]].rename(columns=lambda c: c.replace("_", " ").title()),
                 hide_index=True, width="stretch", height=420)


# =============================================================================
# PAGE: SETTINGS
# =============================================================================


def list_manager(kind, station, noun, user):
    ver = st.session_state.get("ver", 0)
    items = get_list(kind, station)
    kk = f"lm_{kind}_{station}"

    def add_cb():
        if add_to_list(kind, station, [st.session_state.get(kk + "_new", "")], user):
            flash(f"{noun.title()} added.")
        else:
            flash("Enter a new, non-duplicate name.", "⚠️")
        st.session_state[kk + "_new"] = ""
        bump()

    c1, c2 = st.columns([4, 1])
    c1.text_input(f"Add a new {noun}", key=kk + "_new", placeholder=f"Type a {noun} name and press Add")
    c2.write("")
    c2.write("")
    c2.button("➕ Add", on_click=add_cb, key=kk + "_btn", width="stretch")
    with st.expander(f"Paste many {noun}s (one per line)"):
        bulk = st.text_area(f"{noun.title()}s", key=f"{kk}_bulk_{ver}", height=120)
        if st.button("Add all", key=f"{kk}_bulkgo_{ver}"):
            n = add_to_list(kind, station, bulk.splitlines(), user)
            flash(f"{n} {noun}(s) added.")
            bump()
            st.rerun()
    st.markdown(f"**{len(items)} {noun}(s) for {nice(station)}** — edit a name, delete rows, then press Save.")
    df = st.data_editor(pd.DataFrame({"name": pd.Series(items, dtype="object")}), num_rows="dynamic", hide_index=True, width="stretch",
                        key=f"{kk}_ed_{ver}", column_config={"name": st.column_config.TextColumn(noun.title() + " name")})
    if st.button(f"💾  Save {noun} list", type="primary", key=f"{kk}_save_{ver}"):
        n = set_list(kind, station, list(df["name"]), user)
        flash(f"{n} {noun}(s) saved.")
        bump()
        st.rerun()
    st.caption("Renaming a beat does not change duties already saved in past rosters.")


def page_settings(user, is_admin, work, shift_df, designations):
    opts = ["📍 Beats", "👥 Squads", "🔐 Account"] + (["⏱️ Shifts", "🏷️ Designations", "🛠️ Admin"] if is_admin else [])
    sec = st.radio("Section", opts, horizontal=True, key="set_sec", label_visibility="collapsed")
    if sec.startswith("📍"):
        st.markdown(f"#### Beats / locations — {nice(work)}")
        list_manager("beat", work, "beat", user)
    elif sec.startswith("👥"):
        st.markdown(f"#### Squads / teams — {nice(work)}")
        list_manager("squad", work, "squad", user)
    elif sec.startswith("🔐"):
        st.markdown("#### Change your password")
        with st.form("pw", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            old = c1.text_input("Current password", type="password")
            new = c2.text_input("New password", type="password")
            new2 = c3.text_input("Confirm new password", type="password")
            if st.form_submit_button("Change password"):
                if not verify_password(user, old):
                    st.error("Current password is incorrect.")
                elif len(new) < 6 or new != new2:
                    st.error("New password must match and be at least 6 characters.")
                else:
                    with db() as c:
                        set_password(c, user, new)
                        audit(c, user, "PASSWORD", "changed own password")
                    st.success("Password updated.")
    elif sec.startswith("⏱️"):
        st.markdown("#### Shift presets (apply to all stations)")
        sh = shift_df.copy()
        sh["start_time"], sh["end_time"] = sh["start_time"].map(to_time), sh["end_time"].map(to_time)
        new_sh = st.data_editor(
            sh, num_rows="dynamic", hide_index=True, width="stretch", key="shift_ed",
            column_config={"name": st.column_config.TextColumn("Shift name", required=True),
                           "start_time": st.column_config.TimeColumn("Start", format="hh:mm a", step=60),
                           "end_time": st.column_config.TimeColumn("End", format="hh:mm a", step=60)})
        if st.button("Save shift presets", type="primary"):
            rows = [(_s(r["name"]), to_hhmm(r["start_time"]), to_hhmm(r["end_time"])) for _, r in new_sh.iterrows() if _s(r["name"])]
            if len({r[0] for r in rows}) != len(rows):
                st.error("Shift names must be unique.")
            else:
                with db() as c:
                    c.execute("DELETE FROM shifts")
                    for i, r in enumerate(rows):
                        c.execute("INSERT INTO shifts VALUES (?,?,?,?)", (*r, i))
                    audit(c, user, "SETTINGS", "shift presets updated")
                flash("Shift presets saved.")
                st.rerun()
    elif sec.startswith("🏷️"):
        st.markdown("#### Designations")
        des = st.data_editor(pd.DataFrame({"designation": pd.Series(designations, dtype="object")}), num_rows="dynamic", hide_index=True,
                             width="stretch", key="des_ed")
        if st.button("Save designations", type="primary"):
            vals = [d for d in dict.fromkeys(_s(x).upper() for x in des["designation"]) if d]
            with db() as c:
                c.execute("DELETE FROM options WHERE kind='designation'")
                for i, v in enumerate(vals):
                    c.execute("INSERT INTO options VALUES ('designation',?,?)", (v, i))
                audit(c, user, "SETTINGS", "designations updated")
            flash("Designations saved.")
            st.rerun()
    else:  # admin tools
        st.markdown("#### Reset an account password")
        r1, r2, r3 = st.columns([2, 2, 1])
        target = r1.selectbox("Account", [ADMIN] + STATIONS, key="rp_t")
        npw = r2.text_input("New password", type="password", key="rp_p")
        r3.write("")
        r3.write("")
        if r3.button("Reset"):
            if len(npw) < 6:
                st.error("Minimum 6 characters.")
            else:
                with db() as c:
                    set_password(c, target, npw)
                    audit(c, user, "PASSWORD", f"reset for {target}")
                st.success(f"Password reset for {target}.")
        st.divider()
        st.markdown("#### 💾 Backup & restore")
        st.caption("Hosted apps such as Streamlit Cloud can wipe local files on restart — download a backup regularly "
                   "and restore it here if data is ever lost.")
        if os.path.exists(DB_FILE):
            with open(DB_FILE, "rb") as f:
                st.download_button("Download full database backup", f.read(),
                                   f"duty_roster_backup_{date.today():%Y-%m-%d}.db", on_click="ignore")
        up = st.file_uploader("Restore from a backup (.db)", type=["db"], key="restore_up")
        if up is not None and st.checkbox("I understand this replaces ALL current data", key="restore_ok"):
            if st.button("Restore now", type="primary"):
                ok, msg = restore_db(up.getvalue())
                (st.success if ok else st.error)(msg)
                if ok:
                    st.rerun()
        st.divider()
        st.markdown("#### 🧾 Activity log")
        with db() as c:
            log = pd.read_sql_query("SELECT ts AS Time, user AS User, action AS Action, detail AS Detail "
                                    "FROM audit ORDER BY id DESC LIMIT 150", c)
        st.dataframe(log, hide_index=True, width="stretch", height=280)


# =============================================================================
# MAIN
# =============================================================================


def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🛡️", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    init_db()

    if not st.session_state.get("logged_in"):
        login_screen()
        return

    user = st.session_state.station
    is_admin = user == ADMIN
    st.session_state.setdefault("sel_date", date.today())
    if "_goto_page" in st.session_state:
        st.session_state["page"] = st.session_state.pop("_goto_page")
    if st.session_state.get("page") not in PAGES:
        st.session_state["page"] = P_DASH

    sb = st.sidebar
    sb.markdown(f'<div class="side-t">🛡️ {APP_TITLE}</div><div class="side-s">{ORG_LINE}</div>'
                f'<div class="user-chip">👤 <b>{"Master Admin" if is_admin else nice(user)}</b><br>'
                f'<span style="opacity:.7;font-size:.75rem">{"All stations" if is_admin else "Station account"}</span></div>',
                unsafe_allow_html=True)
    page = sb.radio("Menu", PAGES, key="page", label_visibility="collapsed")
    sb.divider()

    def shift_day(n):
        st.session_state.sel_date = st.session_state.sel_date + timedelta(days=n)

    def go_today():
        st.session_state.sel_date = date.today()

    p, m, n_ = sb.columns([1, 2.4, 1])
    p.button("◀", on_click=shift_day, args=(-1,), key="prev", help="Previous day")
    n_.button("▶", on_click=shift_day, args=(1,), key="next", help="Next day")
    m.button("Today", on_click=go_today, key="today", width="stretch")
    sel = sb.date_input("Roster date", key="sel_date", format="DD/MM/YYYY")
    sel_iso = sel.strftime("%Y-%m-%d")

    if is_admin:
        work = sb.selectbox("Working station", STATIONS, format_func=nice, key="work_station",
                            help="Used by Roster Builder, Employees and Settings. Reports and Dashboard cover all stations.")
    else:
        work = user
    sb.divider()
    if sb.button("Log out", width="stretch"):
        for k in ("logged_in", "station", "page", "bstep"):
            st.session_state.pop(k, None)
        st.rerun()

    for msg, icon in st.session_state.pop("_flash", []):
        st.toast(msg, icon=icon)

    shift_df = load_shifts()
    st.markdown(
        f'<div class="hero"><div><div class="hero-t">{page[2:].strip()}</div>'
        f'<div class="hero-s">{sel:%A, %d %B %Y}  ·  {"All stations" if is_admin and page in (P_DASH, P_REP) else nice(work)}</div></div>'
        f'<div><span class="pill">{"Sunday" if sel.weekday() == 6 else "Weekday"}</span></div></div>',
        unsafe_allow_html=True)

    if page == P_DASH:
        page_dashboard(sel, sel_iso, is_admin, user, work, shift_df)
    elif page == P_BUILD:
        page_builder(sel, sel_iso, work, user, is_admin, shift_df)
    elif page == P_EMP:
        page_employees(user, is_admin, work, load_designations())
    elif page == P_REP:
        page_reports(sel, sel_iso, is_admin, user, shift_df)
    else:
        page_settings(user, is_admin, work, shift_df, load_designations())


if __name__ == "__main__":
    main()