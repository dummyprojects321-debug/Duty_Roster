
"""
Duty Roster Command Center
--------------------------
Daily duty-roster management for multi-station enforcement teams.

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
DB_FILE = "duty_roster.db"
SEED_FILE = "DUTY_ROSTER_Standard.xlsx"
ORG_LINE = "PERA  |  DERB Gujranwala"          # printed in the PDF header
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

COLS = [
    "station", "name", "belt_no", "designation", "contact_no",
    "squad_team", "beat_area", "shift", "start_time", "end_time", "remarks",
]
PHONE_RE = re.compile(r"^03\d{2}-\d{7}$")




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


@st.cache_resource(show_spinner=False)
def init_db():
    today = date.today().strftime("%Y-%m-%d")
    with db() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS roster (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                station TEXT, roster_date TEXT, name TEXT, belt_no TEXT,
                designation TEXT, contact_no TEXT, squad_team TEXT,
                beat_area TEXT, shift TEXT, start_time TEXT, end_time TEXT,
                remarks TEXT DEFAULT '', updated_at TEXT, updated_by TEXT
            )""")
        have = {r[1] for r in c.execute("PRAGMA table_info(roster)")}
        for col in ("remarks", "updated_at", "updated_by"):  # migrate old DBs
            if col not in have:
                c.execute(f"ALTER TABLE roster ADD COLUMN {col} TEXT DEFAULT ''")
        c.execute("CREATE INDEX IF NOT EXISTS ix_roster_date ON roster(roster_date, station)")
        c.execute("""CREATE TABLE IF NOT EXISTS shifts (
            name TEXT PRIMARY KEY, start_time TEXT, end_time TEXT, sort INTEGER)""")
        c.execute("CREATE TABLE IF NOT EXISTS options (kind TEXT, value TEXT, sort INTEGER)")
        c.execute("""CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY, salt TEXT, pw_hash TEXT)""")
        c.execute("""CREATE TABLE IF NOT EXISTS audit (
            id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, user TEXT,
            action TEXT, detail TEXT)""")

        if c.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
            for u, p in DEFAULT_PASSWORDS.items():
                set_password(c, u, p)

        # normalise legacy rows (dates -> ISO, times -> HH:MM)
        for rid, d, st_, en in c.execute(
            "SELECT id, roster_date, start_time, end_time FROM roster"
        ).fetchall():
            nd, ns, ne = norm_date(d, today), to_hhmm(st_), to_hhmm(en)
            if (nd, ns, ne) != (d, st_, en):
                c.execute(
                    "UPDATE roster SET roster_date=?, start_time=?, end_time=? WHERE id=?",
                    (nd, ns, ne, rid),
                )

        # seed roster from the Excel template on first run
        if c.execute("SELECT COUNT(*) FROM roster").fetchone()[0] == 0 and os.path.exists(SEED_FILE):
            try:
                df = pd.read_excel(SEED_FILE, sheet_name=0)
                df = df.rename(columns={
                    "Station": "station", "Roster Date": "roster_date", "Name": "name",
                    "Belt No.": "belt_no", "Designation": "designation",
                    "Contact No.": "contact_no", "Squad / Team": "squad_team",
                    "Beat Area": "beat_area", "Shift": "shift",
                    "Start Time": "start_time", "End Time": "end_time",
                })
                valid = [norm_date(x, "") for x in df["roster_date"]]
                seed_date = next((d for d in valid if d), today)
                for col in COLS:
                    if col not in df:
                        df[col] = ""
                out = pd.DataFrame({k: df[k].map(_s) for k in COLS if k not in ("start_time", "end_time")})
                out["start_time"] = df["start_time"].map(to_hhmm)
                out["end_time"] = df["end_time"].map(to_hhmm)
                out["contact_no"] = out["contact_no"].map(norm_phone)
                out["roster_date"] = seed_date
                out["updated_at"] = datetime.now().isoformat(timespec="seconds")
                out["updated_by"] = "seed"
                out.to_sql("roster", c, if_exists="append", index=False)
            except Exception as e:  # noqa: BLE001
                print("Seed error:", e)

        if c.execute("SELECT COUNT(*) FROM shifts").fetchone()[0] == 0:
            known = {n for n, _, _ in DEFAULT_SHIFTS}
            extra = [r[0] for r in c.execute(
                "SELECT DISTINCT shift FROM roster WHERE shift<>''") if r[0] not in known]
            for i, (n, s, e) in enumerate(DEFAULT_SHIFTS + [(x, "09:00", "17:00") for x in extra]):
                c.execute("INSERT INTO shifts VALUES (?,?,?,?)", (n, s, e, i))
        if c.execute("SELECT COUNT(*) FROM options WHERE kind='designation'").fetchone()[0] == 0:
            for i, d in enumerate(DEFAULT_DESIGNATIONS):
                c.execute("INSERT INTO options VALUES ('designation',?,?)", (d, i))
    return True


# ---- loaders ----------------------------------------------------------------


def load_shifts():
    with db() as c:
        df = pd.read_sql_query("SELECT name, start_time, end_time FROM shifts ORDER BY sort", c)
    return df


def load_designations(extra=()):
    with db() as c:
        base = [r[0] for r in c.execute(
            "SELECT value FROM options WHERE kind='designation' ORDER BY sort")]
        used = [r[0] for r in c.execute(
            "SELECT DISTINCT designation FROM roster WHERE designation<>''")]
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
    df = df.sort_values(["_st", "_sh", "start_time", "squad_team", "beat_area", "id"])
    return df.drop(columns=["_st", "_sh"]).reset_index(drop=True)


def load_day(date_iso, stations):
    ph = ",".join("?" * len(stations))
    with db() as c:
        df = pd.read_sql_query(
            f"SELECT * FROM roster WHERE roster_date=? AND station IN ({ph})",
            c, params=[date_iso, *stations])
    for col in COLS:
        df[col] = df[col].map(_s)
    df["start_time"] = df["start_time"].map(to_hhmm)
    df["end_time"] = df["end_time"].map(to_hhmm)
    sh = load_shifts()
    return sort_df(df, {n: i for i, n in enumerate(sh["name"])})


def staff_directory(stations):
    ph = ",".join("?" * len(stations))
    with db() as c:
        df = pd.read_sql_query(
            f"""SELECT name, belt_no, designation, contact_no, station, MAX(id) AS mid
                FROM roster WHERE station IN ({ph}) AND name<>''
                GROUP BY lower(name), upper(belt_no) ORDER BY name""",
            c, params=list(stations))
    return df.fillna("")


def beats_for(stations):
    ph = ",".join("?" * len(stations))
    with db() as c:
        return [r[0] for r in c.execute(
            f"SELECT DISTINCT beat_area FROM roster WHERE station IN ({ph}) AND beat_area<>'' ORDER BY beat_area",
            list(stations))]


# =============================================================================
# EDIT ENGINE  (id-based diff: only touches rows that actually changed)
# =============================================================================


def _record(r, forced_station, presets):
    rec = {
        "station": forced_station or _s(r.get("station")),
        "name": _s(r.get("name")),
        "belt_no": _s(r.get("belt_no")).upper(),
        "designation": _s(r.get("designation")).upper(),
        "contact_no": norm_phone(r.get("contact_no")),
        "squad_team": _s(r.get("squad_team")),
        "beat_area": _s(r.get("beat_area")),
        "shift": _s(r.get("shift")),
        "start_time": to_hhmm(r.get("start_time")),
        "end_time": to_hhmm(r.get("end_time")),
        "remarks": _s(r.get("remarks")),
    }
    p = presets.get(rec["shift"])
    if p:
        rec["start_time"] = rec["start_time"] or p[0]
        rec["end_time"] = rec["end_time"] or p[1]
    return rec


def _is_empty(rec):
    return not any(rec[k] for k in COLS if k != "station")


def validate_edits(edited, forced_station, presets):
    errs = []
    for i, r in edited.iterrows():
        if bool(r.get("Remove")):
            continue
        rec = _record(r, forced_station, presets)
        new = _blank(r.get("id"))
        if new and _is_empty(rec):
            continue
        label = f"Row {i + 1}"
        if not rec["name"]:
            errs.append(f"{label}: name is required.")
        if not rec["station"]:
            errs.append(f"{label}: choose a station.")
    return errs


def apply_edits(orig, edited, date_iso, forced_station, user):
    sh = load_shifts()
    presets = {r.name: (r.start_time, r.end_time) for r in sh.itertuples()}
    o = orig.set_index("id")
    now = datetime.now().isoformat(timespec="seconds")
    stat = {"added": 0, "updated": 0, "removed": 0}
    kept_ids = set()
    with db() as c:
        for _, r in edited.iterrows():
            rid = None if _blank(r.get("id")) else int(r["id"])
            if rid is not None:
                kept_ids.add(rid)
            if bool(r.get("Remove")):
                if rid is not None:
                    c.execute("DELETE FROM roster WHERE id=?", (rid,))
                    stat["removed"] += 1
                continue
            rec = _record(r, forced_station, presets)
            if rid is None:
                if _is_empty(rec):
                    continue
                c.execute(
                    f"INSERT INTO roster (roster_date,{','.join(COLS)},updated_at,updated_by) "
                    f"VALUES (?,{','.join('?' * len(COLS))},?,?)",
                    [date_iso, *[rec[k] for k in COLS], now, user])
                stat["added"] += 1
            elif rid in o.index:
                before = {k: _s(o.loc[rid, k]) for k in COLS}
                if before != rec:
                    sets = ",".join(f"{k}=?" for k in COLS)
                    c.execute(
                        f"UPDATE roster SET {sets}, updated_at=?, updated_by=? WHERE id=?",
                        [*[rec[k] for k in COLS], now, user, rid])
                    stat["updated"] += 1
        for rid in set(int(x) for x in orig["id"]) - kept_ids:  # deleted with the Delete key
            c.execute("DELETE FROM roster WHERE id=?", (rid,))
            stat["removed"] += 1
        if any(stat.values()):
            audit(c, user, "EDIT", f"{date_iso}: +{stat['added']} ~{stat['updated']} -{stat['removed']}")
    return stat


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


def build_pdf(df, date_iso, scope_label, shift_rank):
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
    stations = [s for s in STATIONS if s in set(df["station"])] + sorted(set(df["station"]) - set(STATIONS))

    if df.empty:
        story.append(Spacer(1, 40 * mm))
        story.append(Paragraph("No duty entries recorded for this date.", banner_l))

    for si, stn in enumerate(stations):
        sdf = df[df["station"] == stn]
        banner = Table([[Paragraph(escape(stn), banner_l),
                         Paragraph(f"{len(sdf)} duty entries &nbsp;|&nbsp; {sdf['name'].str.lower().nunique()} personnel", banner_r)]],
                       colWidths=[doc.width * 0.6, doc.width * 0.4])
        banner.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#E6F4F2")),
            ("LINEBEFORE", (0, 0), (0, -1), 3, TEAL),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8)]))
        story += [banner, Spacer(1, 3 * mm)]

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
def cached_pdf(df, date_iso, scope_label, shift_rank):
    return build_pdf(df, date_iso, scope_label, shift_rank)


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
# UI SECTIONS
# =============================================================================


def flash(msg, icon="✅"):
    st.session_state.setdefault("_flash", []).append((msg, icon))


def bump():
    st.session_state.ver = st.session_state.get("ver", 0) + 1


def tab_roster(day, sel, scope, is_admin, user, shift_df, designations, pdf_bytes, xls_bytes, fname):
    # ---- toolbar
    c1, c2, c3, c4 = st.columns([3, 2.4, 1.3, 1.3])
    q = c1.text_input("Search", placeholder="🔎  Search name, belt no., beat, squad…",
                      label_visibility="collapsed", key="q")
    shifts_in_day = [s for s in shift_df["name"] if s in set(day["shift"])] + \
        sorted(set(day["shift"]) - set(shift_df["name"]) - {""})
    shf = c2.multiselect("Shift", shifts_in_day, placeholder="All shifts",
                         label_visibility="collapsed", key="shf")
    c3.download_button("📄  PDF", pdf_bytes, f"{fname}.pdf", "application/pdf",
                       type="primary", width="stretch", on_click="ignore", key="pdf1")
    c4.download_button("📊  Excel", xls_bytes, f"{fname}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       width="stretch", on_click="ignore", key="xl1")

    view = day
    if q:
        needle = q.lower()
        mask = view[["name", "belt_no", "beat_area", "squad_team", "designation", "contact_no"]] \
            .apply(lambda col: col.str.lower().str.contains(needle, regex=False)).any(axis=1)
        view = view[mask]
    if shf:
        view = view[view["shift"].isin(shf)]
    view = view.reset_index(drop=True)

    st.caption("✏️ Click any cell to edit · use the **＋** row at the bottom to add · tick **🗑** to remove a row · "
               "then press **Save changes**. Blank start/end times are filled from the shift preset.")

    ed_df = pd.DataFrame({"id": view["id"], "Remove": False})
    for col in COLS:
        ed_df[col] = view[col]
    ed_df["start_time"] = view["start_time"].map(to_time)
    ed_df["end_time"] = view["end_time"].map(to_time)

    shift_opts = list(shift_df["name"]) + sorted(set(day["shift"]) - set(shift_df["name"]) - {""})
    cfg = {
        "Remove": st.column_config.CheckboxColumn("🗑", width="small", help="Tick to delete on save"),
        "station": st.column_config.SelectboxColumn("Station", options=STATIONS, width="medium"),
        "name": st.column_config.TextColumn("Name", required=True, width="medium"),
        "belt_no": st.column_config.TextColumn("Belt No.", width="small"),
        "designation": st.column_config.SelectboxColumn("Designation", options=designations, width="medium"),
        "contact_no": st.column_config.TextColumn("Contact", width="small", help="Format 03XX-XXXXXXX"),
        "squad_team": st.column_config.TextColumn("Squad / Team", width="small"),
        "beat_area": st.column_config.TextColumn("Beat Area / Location", width="large"),
        "shift": st.column_config.SelectboxColumn("Shift", options=shift_opts, width="small"),
        "start_time": st.column_config.TimeColumn("Start", format="hh:mm a", step=60, width="small"),
        "end_time": st.column_config.TimeColumn("End", format="hh:mm a", step=60, width="small"),
        "remarks": st.column_config.TextColumn("Remarks", width="medium"),
    }
    order = ["Remove"] + (["station"] if is_admin else []) + COLS[1:]
    key = f"ed_{sel}_{'-'.join(scope)}_{q}_{'-'.join(shf)}_{st.session_state.get('ver', 0)}"
    edited = st.data_editor(
        ed_df, key=key, column_config=cfg, column_order=order, hide_index=True,
        num_rows="dynamic", width="stretch",
        height=int(min(680, 44 + 35 * (len(ed_df) + 2))))

    state = st.session_state.get(key, {})
    pending = len(state.get("edited_rows", {})) + len(state.get("added_rows", [])) + len(state.get("deleted_rows", []))

    b1, b2, b3 = st.columns([1.3, 1.1, 5])
    forced = None if is_admin else user
    presets = {r.name: (r.start_time, r.end_time) for r in shift_df.itertuples()}
    if b1.button("💾  Save changes", type="primary", disabled=pending == 0, width="stretch"):
        errs = validate_edits(edited, forced, presets)
        if errs:
            for e in errs[:8]:
                st.error(e)
        else:
            stat = apply_edits(view, edited, sel, forced, user)
            flash(f"Saved — {stat['added']} added, {stat['updated']} updated, {stat['removed']} removed.")
            bump()
            st.rerun()
    if b2.button("↺  Discard", disabled=pending == 0, width="stretch"):
        bump()
        st.rerun()
    dirty = "dirty" if pending else ""
    msg = f"{pending} unsaved change(s)" if pending else "All changes saved"
    b3.markdown(f'<div class="saveline" style="padding-top:8px"><span class="dot {dirty}"></span>{msg}'
                f'{" · <i>save before changing search / shift filter</i>" if pending else ""}</div>',
                unsafe_allow_html=True)


def tab_quick_add(sel, scope, is_admin, user, shift_df, designations):
    ver = st.session_state.get("ver", 0)
    st.markdown("#### ⚡ Assign personnel to a beat")
    st.caption("Pick one or many people, set the beat and shift once — everyone is added in a single click.")
    directory = staff_directory(scope)
    labels = {f"{r.name}  ·  {r.belt_no or '—'}  ·  {r.designation or '—'}": r for r in directory.itertuples()}
    beats = beats_for(scope)
    shifts = list(shift_df["name"])
    presets = {r.name: (r.start_time, r.end_time) for r in shift_df.itertuples()}

    with st.container(border=True):
        a, b = st.columns(2)
        station = a.selectbox("Station", scope, key=f"qa_st_{ver}") if is_admin and len(scope) > 1 \
            else (scope[0] if is_admin else user)
        if not (is_admin and len(scope) > 1):
            a.text_input("Station", value=station, disabled=True, key=f"qa_stx_{ver}")
        people = b.multiselect("Personnel", list(labels), placeholder="Search and select one or more…",
                               key=f"qa_people_{ver}")
        with st.expander("➕ Person not in the list (new staff member)"):
            n1, n2, n3, n4 = st.columns(4)
            new_name = n1.text_input("Name", key=f"qa_nn_{ver}")
            new_belt = n2.text_input("Belt No.", key=f"qa_nb_{ver}")
            new_des = n3.selectbox("Designation", [""] + designations, key=f"qa_nd_{ver}")
            new_phone = n4.text_input("Contact", placeholder="03XX-XXXXXXX", key=f"qa_np_{ver}")

        c1, c2, c3 = st.columns([2, 1.2, 1.2])
        beat_pick = c1.selectbox("Beat area / location", ["➕ Type a new beat…"] + beats, key=f"qa_bp_{ver}")
        beat = c1.text_input("New beat name", key=f"qa_bn_{ver}") if beat_pick.startswith("➕") else beat_pick
        squad = c2.text_input("Squad / Team (optional)", key=f"qa_sq_{ver}")
        shift = c3.selectbox("Shift", shifts, key=f"qa_sh_{ver}")
        d1, d2, d3 = st.columns([1, 1, 2])
        ps, pe = presets.get(shift, ("09:00", "17:00"))
        start = d1.time_input("Start", value=to_time(ps), key=f"qa_s_{ver}_{shift}", step=300)
        end = d2.time_input("End", value=to_time(pe), key=f"qa_e_{ver}_{shift}", step=300)
        remarks = d3.text_input("Remarks (optional)", key=f"qa_rm_{ver}")

        if st.button("➕  Add to roster", type="primary"):
            targets = []
            for lab in people:
                r = labels[lab]
                targets.append((r.name, r.belt_no, r.designation, r.contact_no))
            if new_name.strip():
                targets.append((new_name.strip(), new_belt.strip().upper(), new_des, norm_phone(new_phone)))
            if not targets:
                st.error("Select at least one person, or enter a new staff member.")
            elif not beat.strip():
                st.error("Please provide a beat area.")
            else:
                now = datetime.now().isoformat(timespec="seconds")
                with db() as c:
                    for n, bno, des, ph in targets:
                        c.execute(
                            f"INSERT INTO roster (roster_date,{','.join(COLS)},updated_at,updated_by) "
                            f"VALUES (?,{','.join('?' * len(COLS))},?,?)",
                            [sel, station, n, bno, des, ph, squad.strip(), beat.strip(), shift,
                             start.strftime("%H:%M"), end.strftime("%H:%M"), remarks.strip(), now, user])
                    audit(c, user, "ADD", f"{sel}: {len(targets)} to {beat.strip()} ({shift})")
                flash(f"Added {len(targets)} duty entr{'y' if len(targets) == 1 else 'ies'}.")
                bump()
                st.rerun()


def tab_day_tools(sel, scope, user, is_admin):
    st.markdown("#### 🗓️ Copy a roster from another day")
    st.caption("Rosters rarely change much day to day. Copy yesterday (or any past day), then just edit the differences.")
    ph = ",".join("?" * len(scope))
    with st.container(border=True):
        c1, c2 = st.columns(2)
        src = c1.date_input("Copy FROM", value=sel - timedelta(days=1), format="DD/MM/YYYY", key=f"cp_src_{sel:%Y%m%d}")
        c2.text_input("Copy TO (selected date in sidebar)", value=sel.strftime("%A, %d %B %Y"), disabled=True)
        with db() as c:
            n_src = c.execute(f"SELECT COUNT(*) FROM roster WHERE roster_date=? AND station IN ({ph})",
                              [src.strftime("%Y-%m-%d"), *scope]).fetchone()[0]
            n_dst = c.execute(f"SELECT COUNT(*) FROM roster WHERE roster_date=? AND station IN ({ph})",
                              [sel.strftime("%Y-%m-%d"), *scope]).fetchone()[0]
        st.info(f"**{n_src}** entries on {src:%d %b %Y}  →  **{n_dst}** entries already on {sel:%d %b %Y}")
        mode = st.radio("If the target day already has entries", ["Add to existing", "Replace existing"],
                        horizontal=True)
        ok = True
        if mode == "Replace existing" and n_dst:
            ok = st.checkbox(f"I understand this deletes the {n_dst} existing entries on {sel:%d %b %Y}")
        if st.button("📋  Copy roster", type="primary", disabled=(n_src == 0 or not ok or src == sel)):
            now = datetime.now().isoformat(timespec="seconds")
            with db() as c:
                if mode == "Replace existing":
                    c.execute(f"DELETE FROM roster WHERE roster_date=? AND station IN ({ph})",
                              [sel.strftime("%Y-%m-%d"), *scope])
                c.execute(
                    f"INSERT INTO roster (roster_date,{','.join(COLS)},updated_at,updated_by) "
                    f"SELECT ?,{','.join(COLS)},?,? FROM roster WHERE roster_date=? AND station IN ({ph})",
                    [sel.strftime("%Y-%m-%d"), now, user, src.strftime("%Y-%m-%d"), *scope])
                audit(c, user, "COPY", f"{src:%Y-%m-%d} -> {sel:%Y-%m-%d} ({mode}, {n_src})")
            flash(f"Copied {n_src} entries from {src:%d %b %Y}.")
            bump()
            st.rerun()

    with st.expander("🧨 Danger zone — clear this day"):
        sure = st.checkbox(f"Delete ALL {n_dst} entries on {sel:%d %b %Y} for the selected station(s)")
        if st.button("Clear day", disabled=not (sure and n_dst)):
            with db() as c:
                c.execute(f"DELETE FROM roster WHERE roster_date=? AND station IN ({ph})",
                          [sel.strftime("%Y-%m-%d"), *scope])
                audit(c, user, "CLEAR", f"{sel:%Y-%m-%d}: {n_dst} rows")
            flash(f"Cleared {n_dst} entries.", "🧹")
            bump()
            st.rerun()


def tab_reports(sel, all_scope_choices, is_admin, user, shift_df):
    st.markdown("#### 📄 One-click roster reports")
    if is_admin:
        rep_scope = st.multiselect("Stations to include", STATIONS, default=all_scope_choices,
                                   key="rep_scope") or STATIONS
    else:
        rep_scope = [user]
    df = load_day(sel.strftime("%Y-%m-%d"), rep_scope)
    label = "All Stations" if set(rep_scope) == set(STATIONS) else ", ".join(rep_scope).title()
    rank = {n: i for i, n in enumerate(shift_df["name"])}
    fname = f"Duty_Roster_{'ALL' if set(rep_scope) == set(STATIONS) else '_'.join(s.replace(' ', '') for s in rep_scope)}_{sel:%Y-%m-%d}"
    if df.empty:
        st.warning("No entries on this date for the selected stations.")
        return
    a, b, _ = st.columns([1.2, 1.2, 3])
    a.download_button("📄  Download PDF", cached_pdf(df, sel.strftime("%Y-%m-%d"), label, rank),
                      f"{fname}.pdf", "application/pdf", type="primary", width="stretch",
                      on_click="ignore", key="pdf2")
    b.download_button("📊  Download Excel", cached_excel(df, sel.strftime("%Y-%m-%d")), f"{fname}.xlsx",
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                      width="stretch", on_click="ignore", key="xl2")
    st.caption("PDF is landscape A4 — grouped by station and shift, with signature blocks, ready to print.")
    prev = df.assign(Timing=df.apply(lambda r: f"{fmt12(r.start_time)} – {fmt12(r.end_time)}", axis=1))
    st.dataframe(prev[["station", "name", "belt_no", "designation", "contact_no", "squad_team",
                       "beat_area", "shift", "Timing", "remarks"]]
                 .rename(columns=lambda c: c.replace("_", " ").title()),
                 hide_index=True, width="stretch", height=420)


def tab_settings(user, is_admin, shift_df, designations):
    st.markdown("#### 🔐 Account")
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
    if not is_admin:
        return

    st.divider()
    st.markdown("#### ⏱️ Shift presets")
    st.caption("These feed the shift dropdowns and auto-fill start / end times.")
    sh = shift_df.copy()
    sh["start_time"], sh["end_time"] = sh["start_time"].map(to_time), sh["end_time"].map(to_time)
    new_sh = st.data_editor(
        sh, num_rows="dynamic", hide_index=True, width="stretch", key="shift_ed",
        column_config={"name": st.column_config.TextColumn("Shift name", required=True),
                       "start_time": st.column_config.TimeColumn("Start", format="hh:mm a", step=60),
                       "end_time": st.column_config.TimeColumn("End", format="hh:mm a", step=60)})
    if st.button("Save shift presets"):
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

    st.markdown("#### 🏷️ Designations")
    des = st.data_editor(pd.DataFrame({"designation": designations}), num_rows="dynamic", hide_index=True,
                         width="stretch", key="des_ed")
    if st.button("Save designations"):
        vals = [d for d in dict.fromkeys(_s(x).upper() for x in des["designation"]) if d]
        with db() as c:
            c.execute("DELETE FROM options WHERE kind='designation'")
            for i, v in enumerate(vals):
                c.execute("INSERT INTO options VALUES ('designation',?,?)", (v, i))
            audit(c, user, "SETTINGS", "designations updated")
        flash("Designations saved.")
        st.rerun()

    st.divider()
    st.markdown("#### 👥 Reset a station password")
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
    st.markdown("#### 🧾 Activity log & backup")
    with db() as c:
        log = pd.read_sql_query("SELECT ts AS Time, user AS User, action AS Action, detail AS Detail "
                                "FROM audit ORDER BY id DESC LIMIT 150", c)
    st.dataframe(log, hide_index=True, width="stretch", height=260)
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "rb") as f:
            st.download_button("💾  Download full database backup", f.read(),
                               f"duty_roster_backup_{date.today():%Y-%m-%d}.db", on_click="ignore")


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

    # ---- sidebar
    sb = st.sidebar
    sb.markdown(f'<div class="side-t">🛡️ {APP_TITLE}</div><div class="side-s">{ORG_LINE}</div>'
                f'<div class="user-chip">👤 <b>{"Master Admin" if is_admin else user.title()}</b><br>'
                f'<span style="opacity:.7;font-size:.75rem">{"All stations" if is_admin else "Station account"}</span></div>',
                unsafe_allow_html=True)

    def shift_day(n):
        st.session_state.sel_date = st.session_state.sel_date + timedelta(days=n)

    def go_today():
        st.session_state.sel_date = date.today()

    p, m, n = sb.columns([1, 3, 1])
    p.button("◀", on_click=shift_day, args=(-1,), key="prev", help="Previous day")
    n.button("▶", on_click=shift_day, args=(1,), key="next", help="Next day")
    m.button("Today", on_click=go_today, key="today", width="stretch")
    sel = sb.date_input("Roster date", key="sel_date", format="DD/MM/YYYY")
    sel_iso = sel.strftime("%Y-%m-%d")

    if is_admin:
        scope = sb.multiselect("Stations", STATIONS, default=STATIONS, key="scope_ms") or STATIONS
    else:
        scope = [user]
    sb.divider()
    if sb.button("Log out", width="stretch"):
        for k in ("logged_in", "station"):
            st.session_state.pop(k, None)
        st.rerun()

    for msg, icon in st.session_state.pop("_flash", []):
        st.toast(msg, icon=icon)

    # ---- data
    shift_df = load_shifts()
    rank = {n_: i for i, n_ in enumerate(shift_df["name"])}
    day = load_day(sel_iso, scope)
    designations = load_designations()
    issues = health_check(day) if len(day) else pd.DataFrame(columns=["Issue", "Person", "Station", "Details"])
    scope_label = "All Stations" if set(scope) == set(STATIONS) else ", ".join(scope).title()
    fname = f"Duty_Roster_{'ALL' if set(scope) == set(STATIONS) else '_'.join(s.replace(' ', '') for s in scope)}_{sel_iso}"

    # ---- hero + KPIs
    st.markdown(
        f'<div class="hero"><div><div class="hero-t">{sel:%A, %d %B %Y}</div>'
        f'<div class="hero-s">Duty roster · {escape(scope_label)}</div></div>'
        f'<div><span class="pill">{len(day)} entries</span>'
        f'<span class="pill">{"Sunday roster" if sel.weekday() == 6 else "Weekday roster"}</span></div></div>',
        unsafe_allow_html=True)

    people = day.apply(person_key, axis=1).nunique() if len(day) else 0
    beats_n = day.loc[day["beat_area"] != "", "beat_area"].nunique() if len(day) else 0
    crit = int(issues["Issue"].str.startswith("🔴").sum()) if len(issues) else 0
    cards([
        ("Personnel on duty", people, f"{len(day)} duty entries", ""),
        ("Beats covered", beats_n, "distinct locations", ""),
        ("Stations", day["station"].nunique() if len(day) else 0, f"of {len(scope)} selected", ""),
        ("Double-bookings", crit, "overlapping shifts" if crit else "none detected", "bad" if crit else "ok"),
        ("Data checks", len(issues), "items to review" if len(issues) else "all clear", "warn" if len(issues) else "ok"),
    ])

    pdf_bytes = cached_pdf(day, sel_iso, scope_label, rank)
    xls_bytes = cached_excel(day, sel_iso) if len(day) else b""

    tabs = st.tabs(["📋 Roster", "⚡ Quick Add", "🗓️ Day Tools", "📄 Reports", "⚙️ Settings"])
    with tabs[0]:
        if day.empty:
            st.info("No roster for this date yet. Use **Day Tools → Copy roster** to start from a previous day, "
                    "or add people with **Quick Add**.")
        tab_roster(day, sel_iso, scope, is_admin, user, shift_df, designations,
                   pdf_bytes if len(day) else b"", xls_bytes, fname)
        if len(day):
            with st.expander("📊 Coverage summary"):
                ct = pd.crosstab(day["station"], day["shift"].replace("", "Unassigned"), margins=True, margins_name="Total")
                st.dataframe(ct, width="stretch")
        if len(issues):
            with st.expander(f"🩺 Data checks — {len(issues)} item(s) to review"):
                st.dataframe(issues, hide_index=True, width="stretch")
    with tabs[1]:
        tab_quick_add(sel_iso, scope, is_admin, user, shift_df, designations)
    with tabs[2]:
        tab_day_tools(sel, scope, user, is_admin)
    with tabs[3]:
        tab_reports(sel, scope, is_admin, user, shift_df)
    with tabs[4]:
        tab_settings(user, is_admin, shift_df, designations)


if __name__ == "__main__":
    main()

import io
import pandas as pd
import sqlite3
import streamlit as st

# --- DATABASE SETUP ---
DB_FILE = "duty_roster.db"


def init_db():
  conn = sqlite3.connect(DB_FILE)
  cursor = conn.cursor()
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS roster (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            station TEXT,
            roster_date TEXT,
            name TEXT,
            belt_no TEXT,
            designation TEXT,
            contact_no TEXT,
            squad_team TEXT,
            beat_area TEXT,
            shift TEXT,
            start_time TEXT,
            end_time TEXT
        )
    """)
  conn.commit()

  # Check if table is empty, if so, seed from uploaded excel file
  cursor.execute("SELECT COUNT(*) FROM roster")
  if cursor.fetchone()[0] == 0:
    try:
      df_excel = pd.read_excel("DUTY_ROSTER_Standard.xlsx", sheet_name=0)
      # Normalize columns to match db
      df_excel.columns = [
          "Sr No",
          "station",
          "roster_date",
          "name",
          "belt_no",
          "designation",
          "contact_no",
          "squad_team",
          "beat_area",
          "shift",
          "start_time",
          "end_time",
      ]
      # Drop Sr No as id is autoincrement
      df_excel = df_excel.drop(columns=["Sr No"])
      # Fill NaN with empty strings
      df_excel = df_excel.fillna("")
      df_excel.to_sql("roster", conn, if_exists="append", index=False)
    except Exception as e:
      st.error(f"Error seeding database from Excel: {e}")
  conn.close()


init_db()

# --- STREAMLIT UI CONFIG ---
st.set_page_config(
    page_title="Duty Roster Management Portal", page_icon="🛡️", layout="wide"
)

# --- AUTHENTICATION CONFIG ---
# Define station credentials (you can modify passwords here)
STATION_PASSWORDS = {
    "MASTER_ADMIN": "admin123",
    "KAMOKE": "kamoke2026",
    "NOWSHEHRA VIRKAN": "nowshehra2026",
    "GUJRANWALA CITY": "city2026",
    "WAZIRABAD": "wazirabad2026",
    "GUJRANWALA SADDAR": "saddar2026",
}

if "logged_in" not in st.session_state:
  st.session_state.logged_in = False
  st.session_state.station = ""

# --- LOGIN SCREEN ---
if not st.session_state.logged_in:
  st.title("🛡️ Duty Roster Management Portal")
  st.write(
      "Please log in using your assigned station credentials to manage your"
      " roster."
  )

  col1, col2 = st.columns([1, 1])
  with col1:
    selected_station = st.selectbox(
        "Select Station / Role", list(STATION_PASSWORDS.keys())
    )
    password_input = st.text_input("Enter Password", type="password")

    if st.button("Login", type="primary"):
      if password_input == STATION_PASSWORDS.get(selected_station):
        st.session_state.logged_in = True
        st.session_state.station = selected_station
        st.rerun()
      else:
        st.error("Incorrect Password. Please try again.")
  st.stop()

# --- MAIN APP DASHBOARD ---
current_station = st.session_state.station
st.sidebar.title("Navigation")
st.sidebar.write(f"Logged in as: **{current_station}**")

if st.sidebar.button("Logout"):
  st.session_state.logged_in = False
  st.session_state.station = ""
  st.rerun()

st.title(f"📋 Duty Roster Portal - {current_station}")

# Load Data from SQLite
conn = sqlite3.connect(DB_FILE)
if current_station == "MASTER_ADMIN":
  df_roster = pd.read_sql("SELECT * FROM roster", conn)
  stations_list = [
      "KAMOKE",
      "NOWSHEHRA VIRKAN",
      "GUJRANWALA CITY",
      "WAZIRABAD",
      "GUJRANWALA SADDAR",
  ]
else:
  df_roster = pd.read_sql(
      "SELECT * FROM roster WHERE station = ?", conn, params=(current_station,)
  )
  stations_list = [current_station]

# Tabs for management
tab1, tab2, tab3 = st.tabs(
    ["📊 View & Edit Roster", "➕ Add New Staff Member", "📥 Download Excel"]
)

with tab1:
  st.subheader("Live Roster Table (Double click cells to edit directly)")
  st.info(
      "You can edit names, belt numbers, designations, beat areas, and timings"
      " directly in the table below. Changes save automatically."
  )

  # Data Editor for full editing rights
  edited_df = st.data_editor(
      df_roster,
      num_rows="dynamic",
      key="roster_editor",
      use_container_width=True,
      hide_index=True,
  )

  if st.button("Save Changes to Database", type="primary"):
    # Overwrite station records or full table depending on role
    cursor = conn.cursor()
    if current_station == "MASTER_ADMIN":
      cursor.execute("DELETE FROM roster")
      edited_df.to_sql("roster", conn, if_exists="append", index=False)
    else:
      cursor.execute("DELETE FROM roster WHERE station = ?", (current_station,))
      edited_df.to_sql("roster", conn, if_exists="append", index=False)
    conn.commit()
    st.success("Roster successfully updated and saved!")
    st.rerun()

with tab2:
  st.subheader("Add a New Staff Member / Duty Entry")
  with st.form("add_staff_form"):
    col1, col2, col3 = st.columns(3)
    with col1:
      target_station = (
          st.selectbox("Station", stations_list)
          if current_station == "MASTER_ADMIN"
          else current_station
      )
      name = st.text_input("Staff Name")
      belt_no = st.text_input("Belt No (e.g., SGT-2613)")
      designation = st.text_input("Designation (e.g., SERGEANT)")
    with col2:
      contact_no = st.text_input("Contact No (e.g., 0334-6499717)")
      squad_team = st.text_input("Squad / Team (Optional)")
      beat_area = st.text_input("Beat Area / Location")
      shift = st.selectbox("Shift", ["Morning", "Evening", "Night"])
    with col3:
      roster_date = st.text_input("Roster Date (DD-MM-YYYY)", value="")
      start_time = st.text_input("Start Time", value="09:00:00")
      end_time = st.text_input("End Time", value="17:00:00")

    submitted = st.form_submit_button("Add Member")
    if submitted:
      if name:
        cursor = conn.cursor()
        cursor.execute(
            """
                    INSERT INTO roster (station, roster_date, name, belt_no, designation, contact_no, squad_team, beat_area, shift, start_time, end_time)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
            (
                target_station,
                roster_date,
                name,
                belt_no,
                designation,
                contact_no,
                squad_team,
                beat_area,
                shift,
                start_time,
                end_time,
            ),
        )
        conn.commit()
        st.success(f"Successfully added {name} to {target_station} roster!")
        st.rerun()
      else:
        st.error("Staff Name is required.")

with tab3:
  st.subheader("Export Current Roster to Excel")


  def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df.to_excel(writer, index=False, sheet_name="Roster")
    return output.getvalue()


  excel_data = convert_df_to_excel(df_roster)
  st.download_button(
      label="📥 Download Roster as Excel File",
      data=excel_data,
      file_name=f"Duty_Roster_{current_station}.xlsx",
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
  )

conn.close()

