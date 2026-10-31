"""
leave_range.py - two add-on features for the Duty Roster app.

  1. LEAVE      mark an official "on leave" from a date to a date (type + reason)
  2. DATE RANGE apply a roster to a "from - to" period in one click

Safety:
  * This file only ADDS two new tables (staff_leave, roster_period) with
    CREATE TABLE IF NOT EXISTS. Your existing tables / rows are never altered.
  * It never touches a roster day that is marked FINAL.
  * Days that already have a roster are skipped unless you explicitly tick "Replace".
  * Works on both SQLite and PostgreSQL (uses core.db()).
"""
from __future__ import annotations

from datetime import date, timedelta

import pandas as pd

import core

LEAVE_TYPES = ["Casual Leave", "Medical Leave", "Earned Leave", "Special Leave", "Other"]
MAX_RANGE_DAYS = 366


# =============================================================================
# setup
# =============================================================================
def _schema() -> list[str]:
    pk = "SERIAL PRIMARY KEY" if core.using_pg() else "INTEGER PRIMARY KEY AUTOINCREMENT"
    return [
        f"CREATE TABLE IF NOT EXISTS staff_leave("
        f"id {pk}, station TEXT NOT NULL, staff_id INTEGER NOT NULL, date_from TEXT NOT NULL, "
        f"date_to TEXT NOT NULL, leave_type TEXT DEFAULT '', reason TEXT DEFAULT '', "
        f"created_by TEXT, created_at TEXT)",
        "CREATE INDEX IF NOT EXISTS ix_leave_staff ON staff_leave(staff_id, date_from, date_to)",
        f"CREATE TABLE IF NOT EXISTS roster_period("
        f"id {pk}, station TEXT NOT NULL, date_from TEXT NOT NULL, date_to TEXT NOT NULL, "
        f"source_date TEXT, title TEXT DEFAULT '', created_by TEXT, created_at TEXT)",
    ]


def init():
    """Create the two new tables (safe to call on every start)."""
    with core.db() as c:
        for stmt in _schema():
            c.execute(stmt)


# =============================================================================
# small helpers
# =============================================================================
def _d(v) -> date:
    return v if isinstance(v, date) else date.fromisoformat(str(v))


def _days(a: date, b: date) -> list[date]:
    return [a + timedelta(days=i) for i in range((b - a).days + 1)]


def _fmt(v) -> str:
    return _d(v).strftime("%d/%m/%Y")


def _remark(ltype: str, reason: str, d_to) -> str:
    txt = ltype or "Leave"
    if reason:
        txt += f": {reason}"
    return f"On leave - {txt} (till {_fmt(d_to)})"


def _leaves_in(c, station: str, a_iso: str, b_iso: str) -> list[tuple]:
    """Leave records of this station's people that overlap a_iso..b_iso."""
    return c.execute(
        "SELECT l.staff_id, l.date_from, l.date_to, l.leave_type, l.reason "
        "FROM staff_leave l JOIN staff s ON s.id=l.staff_id "
        "WHERE s.station=? AND l.date_from<=? AND l.date_to>=?",
        (station, b_iso, a_iso)).fetchall()


def _leave_on(leaves: list[tuple], iso: str) -> dict:
    """{staff_id: remark text} for everybody on leave on that date."""
    return {int(sid): _remark(lt, rs, to) for sid, f, to, lt, rs in leaves if f <= iso <= to}


def on_leave(station: str, iso: str) -> dict:
    """{staff_id: remark} of people on leave on a given date."""
    with core.db() as c:
        return _leave_on(_leaves_in(c, station, iso, iso), iso)


# =============================================================================
# roster grid that knows about leave  (use instead of core.roster_grid on the Roster page)
# =============================================================================
def roster_grid(station: str, iso: str) -> pd.DataFrame:
    """core.roster_grid + people on leave shown as status 'Leave' (only if not already placed on duty)."""
    g = core.roster_grid(station, iso)
    lv = on_leave(station, iso)
    if g.empty or not lv:
        return g
    for i, r in g.iterrows():
        sid = int(r["staff_id"])
        if sid in lv and r["status"] == "Duty" and not (r["beat"] or r["shift"] or r["slot"]):
            g.at[i, "status"] = "Leave"
            g.at[i, "remarks"] = r["remarks"] or lv[sid]
    return g


# =============================================================================
# LEAVE: add / list / remove
# =============================================================================
def _sync_leave_into_rosters(c, station, sid, d1: date, d2: date, ltype, reason, by):
    """For days that ALREADY have a (non-final) roster, write the person as Leave. Returns (done, final_days)."""
    done, final_days = [], []
    remark = _remark(ltype, reason, d2)
    for d in _days(d1, d2):
        iso = d.isoformat()
        if not c.execute("SELECT 1 FROM roster WHERE station=? AND roster_date=?", (station, iso)).fetchone():
            continue
        if c.execute("SELECT 1 FROM roster_day WHERE station=? AND roster_date=?", (station, iso)).fetchone():
            final_days.append(iso)
            continue
        cur = c.execute("SELECT status FROM roster WHERE roster_date=? AND staff_id=?", (iso, sid)).fetchone()
        if cur and cur[0] == "Off Day":
            continue
        core._upsert(c, station, iso, sid, "Leave", "", "", "", remark, by)
        done.append(iso)
    return done, final_days


def add_leave(station, staff_id, d_from, d_to, leave_type, reason, by):
    d1, d2 = _d(d_from), _d(d_to)
    if d2 < d1:
        return False, "'Leave to' cannot be before 'Leave from'."
    if (d2 - d1).days + 1 > MAX_RANGE_DAYS:
        return False, f"Leave cannot be longer than {MAX_RANGE_DAYS} days."
    sid = int(staff_id)
    reason = " ".join(core._s(reason).split())
    ltype = leave_type if leave_type in LEAVE_TYPES else "Other"
    with core.db() as c:
        row = c.execute("SELECT name FROM staff WHERE id=? AND station=?", (sid, station)).fetchone()
        if not row:
            return False, "Person not found at this station."
        clash = c.execute(
            "SELECT date_from, date_to FROM staff_leave WHERE staff_id=? AND date_from<=? AND date_to>=?",
            (sid, d2.isoformat(), d1.isoformat())).fetchone()
        if clash:
            return False, f"{row[0]} already has leave {_fmt(clash[0])} to {_fmt(clash[1])} that overlaps these dates."
        c.execute(
            "INSERT INTO staff_leave(station, staff_id, date_from, date_to, leave_type, reason, created_by, created_at) "
            "VALUES (?,?,?,?,?,?,?,?)",
            (station, sid, d1.isoformat(), d2.isoformat(), ltype, reason, by, core.now_iso()))
        done, final_days = _sync_leave_into_rosters(c, station, sid, d1, d2, ltype, reason, by)
        core.audit(by, "add_leave", f"{station}: {row[0]} {d1}..{d2} {ltype}", c)
    msg = f"{row[0]} marked on leave {_fmt(d1)} to {_fmt(d2)}."
    if done:
        msg += f" Updated {len(done)} existing roster day(s)."
    if final_days:
        msg += (f" {len(final_days)} FINAL day(s) were not changed ({', '.join(_fmt(x) for x in final_days[:4])}"
                f"{'…' if len(final_days) > 4 else ''}) - reopen them if they must show the leave.")
    return True, msg


def list_leaves(station: str, include_past: bool = False) -> pd.DataFrame:
    q = ("SELECT l.id, s.name, s.belt_no, s.designation, l.date_from, l.date_to, l.leave_type, l.reason "
         "FROM staff_leave l JOIN staff s ON s.id=l.staff_id WHERE s.station=?")
    args = [station]
    if not include_past:
        q += " AND l.date_to>=?"
        args.append(core.today_iso())
    q += " ORDER BY l.date_from DESC, s.name"
    with core.db() as c:
        df = pd.read_sql_query(q, c, params=args).fillna("")
    t = core.today_iso()
    df["state"] = df.apply(lambda r: "Past" if r["date_to"] < t else ("On leave now" if r["date_from"] <= t else "Upcoming"), axis=1)
    return df


def remove_leaves(station: str, ids, by) -> int:
    """Delete leave records. Roster rows already written as 'Leave' are NOT changed - edit them on the Roster page."""
    n = 0
    with core.db() as c:
        for i in ids:
            n += c.execute(
                "DELETE FROM staff_leave WHERE id=? AND staff_id IN (SELECT id FROM staff WHERE station=?)",
                (int(i), station)).rowcount
        if n:
            core.audit(by, "remove_leave", f"{station}: {n}", c)
    return n


# =============================================================================
# DATE RANGE: apply one roster to every day from - to
# =============================================================================
def apply_range(station, src_iso, d_from, d_to, by, replace=False, title=""):
    """Copy the roster of src_iso to every day in d_from..d_to.

    * people on leave on a target day become status 'Leave' (with the reason in remarks)
    * days that already have a roster are skipped unless replace=True
    * FINAL days are never changed
    """
    d1, d2 = _d(d_from), _d(d_to)
    if d2 < d1:
        return {"ok": False, "msg": "'To' date cannot be before 'From' date."}
    if (d2 - d1).days + 1 > MAX_RANGE_DAYS:
        return {"ok": False, "msg": f"A period can be at most {MAX_RANGE_DAYS} days."}
    written, skipped_existing, skipped_final, entries = [], [], [], 0
    with core.db() as c:
        src = c.execute(
            "SELECT r.staff_id, r.status, r.beat, r.shift, r.slot, r.remarks FROM roster r "
            "JOIN staff s ON s.id=r.staff_id WHERE r.station=? AND r.roster_date=? AND s.active=1 AND s.station=?",
            (station, src_iso, station)).fetchall()
        if not src:
            return {"ok": False, "msg": f"No saved roster for {_fmt(src_iso)}. Build and save it first."}
        active = {int(r[0]) for r in c.execute("SELECT id FROM staff WHERE station=? AND active=1", (station,))}
        leaves = _leaves_in(c, station, min(src_iso, d1.isoformat()), d2.isoformat())
        src_leave = _leave_on(leaves, src_iso)
        for d in _days(d1, d2):
            iso = d.isoformat()
            if iso == src_iso:
                continue
            exists = c.execute("SELECT 1 FROM roster WHERE station=? AND roster_date=?", (station, iso)).fetchone()
            final = c.execute("SELECT 1 FROM roster_day WHERE station=? AND roster_date=?", (station, iso)).fetchone()
            if final:
                skipped_final.append(iso)
                continue
            if exists and not replace:
                skipped_existing.append(iso)
                continue
            if exists:
                c.execute("DELETE FROM roster WHERE station=? AND roster_date=?", (station, iso))
            lv = _leave_on(leaves, iso)
            placed = set()
            for sid, status, beat, shift, slot, rem in src:
                sid = int(sid)
                if sid in lv:
                    core._upsert(c, station, iso, sid, "Leave", "", "", "", lv[sid], by)
                elif status == "Leave" and sid in src_leave:
                    continue            # leave that ended before this day: back to 'not assigned'
                else:
                    core._upsert(c, station, iso, sid, status, beat, shift, slot, rem, by)
                placed.add(sid)
                entries += 1
            for sid, remark in lv.items():      # on leave but not in the source roster
                if sid in active and sid not in placed:
                    core._upsert(c, station, iso, sid, "Leave", "", "", "", remark, by)
                    entries += 1
            written.append(iso)
        c.execute(
            "INSERT INTO roster_period(station, date_from, date_to, source_date, title, created_by, created_at) "
            "VALUES (?,?,?,?,?,?,?)",
            (station, d1.isoformat(), d2.isoformat(), src_iso, " ".join(core._s(title).split()), by, core.now_iso()))
        core.audit(by, "apply_range", f"{station} {src_iso} -> {d1}..{d2}: {len(written)} days", c)
    return {"ok": True, "written": written, "skipped_existing": skipped_existing,
            "skipped_final": skipped_final, "entries": entries}


def list_periods(station: str, limit: int = 8) -> pd.DataFrame:
    with core.db() as c:
        df = pd.read_sql_query(
            "SELECT date_from, date_to, source_date, title, created_by, created_at FROM roster_period "
            "WHERE station=? ORDER BY id DESC LIMIT ?", c, params=[station, limit]).fillna("")
    for col in ("date_from", "date_to", "source_date"):
        df[col] = df[col].map(lambda v: _fmt(v) if v else "")
    return df.rename(columns={"date_from": "From", "date_to": "To", "source_date": "Copied from",
                              "title": "Note", "created_by": "By", "created_at": "Saved at"})


# =============================================================================
# STREAMLIT UI (called from app.py)
# =============================================================================
def _flash(msg: str, kind: str = "success"):
    import streamlit as st
    st.session_state.setdefault("_flash", []).append((kind, msg))


def _bump():
    import streamlit as st
    st.session_state["ver"] = st.session_state.get("ver", 0) + 1


def leave_tab(user, iso: str):
    """Content of the new 'Leave' tab on the Staff page."""
    import streamlit as st
    stn, by = user["station"], user["username"]
    ver = st.session_state.get("ver", 0)
    staff = core.staff_df(stn)
    if staff.empty:
        st.info("Add your staff first.")
        return
    labels = {int(r.id): f"{r.name} · {r.belt_no or '—'} · {r.designation or '—'}" for r in staff.itertuples()}
    day = date.fromisoformat(iso)

    st.markdown("#### Mark an official on leave")
    with st.form(f"leave_form_{ver}", clear_on_submit=True):
        who = st.selectbox("Official", list(labels), format_func=labels.get, index=None, placeholder="Select a person")
        c1, c2 = st.columns(2)
        d1 = c1.date_input("Leave from", value=day, format="DD/MM/YYYY")
        d2 = c2.date_input("Leave to (last day of leave)", value=day, format="DD/MM/YYYY")
        c3, c4 = st.columns([1, 2])
        ltype = c3.selectbox("Leave type", LEAVE_TYPES)
        reason = c4.text_input("Reason / remarks (optional)")
        go = st.form_submit_button("Mark on leave", type="primary")
    if go:
        if who is None:
            st.warning("Choose the official.")
        else:
            ok, msg = add_leave(stn, who, d1, d2, ltype, reason, by)
            _flash(msg, "success" if ok else "warning")
            _bump()
            st.rerun()

    st.markdown("#### Leave register")
    past = st.checkbox("Show past leave too", value=False, key=f"lv_past_{ver}")
    df = list_leaves(stn, include_past=past)
    if df.empty:
        st.caption("No leave recorded.")
        return
    view = df.assign(date_from=df["date_from"].map(_fmt), date_to=df["date_to"].map(_fmt))
    view = view[["name", "belt_no", "designation", "date_from", "date_to", "leave_type", "reason", "state"]]
    view.columns = ["Name", "Belt No.", "Designation", "From", "To", "Type", "Reason", "Status"]
    st.dataframe(view, hide_index=True, width="stretch", height=min(40 + 35 * len(view), 420))
    opts = {int(r.id): f"{r.name} · {_fmt(r.date_from)} to {_fmt(r.date_to)}" for r in df.itertuples()}
    rem = st.multiselect("Cancel leave record(s)", list(opts), format_func=opts.get, key=f"lv_rm_{ver}")
    if st.button("Cancel selected leave", disabled=not rem, key=f"lv_rmgo_{ver}"):
        n = remove_leaves(stn, rem, by)
        _flash(f"{n} leave record(s) cancelled. Days already written as Leave on a roster stay as they are - "
               f"change them on the Roster page if needed.", "info")
        _bump()
        st.rerun()


def range_panel(user, iso: str):
    """Expander for the Roster page: apply the roster of the selected date to a from - to period."""
    import streamlit as st
    stn, by = user["station"], user["username"]
    ver = st.session_state.get("ver", 0)
    start = date.fromisoformat(iso)
    with st.expander("📆 Apply this roster to a date range (from – to)"):
        st.caption(f"Uses the saved roster of **{_fmt(iso)}** (the date chosen in the sidebar) as the roster for every day "
                   "in the period. Officials on leave are marked **Leave** automatically. "
                   "Days that already have a roster are not touched unless you tick Replace. Final days are never changed.")
        rng = st.date_input("Roster applicable from – to", value=(start, start + timedelta(days=6)),
                            format="DD/MM/YYYY", key=f"rp_rng_{ver}")
        c1, c2 = st.columns([2, 3])
        replace = c1.checkbox("Replace days that already have a draft roster", value=False, key=f"rp_rep_{ver}")
        title = c2.text_input("Note (optional)", placeholder="e.g. Muharram security duty", key=f"rp_title_{ver}")
        if st.button("Apply to this period", type="primary", key=f"rp_go_{ver}"):
            if not (isinstance(rng, (tuple, list)) and len(rng) == 2):
                st.warning("Pick both a start date and an end date.")
            else:
                res = apply_range(stn, iso, rng[0], rng[1], by, replace=replace, title=title)
                if not res["ok"]:
                    _flash(res["msg"], "warning")
                else:
                    _flash(f"Roster applied to {len(res['written'])} day(s) ({res['entries']} entries).")
                    if res["skipped_existing"]:
                        _flash(f"{len(res['skipped_existing'])} day(s) already had a roster and were kept.", "info")
                    if res["skipped_final"]:
                        _flash(f"{len(res['skipped_final'])} final day(s) were not changed.", "info")
                _bump()
                st.rerun()
        periods = list_periods(stn)
        if len(periods):
            st.markdown("**Recent periods**")
            st.dataframe(periods, hide_index=True, width="stretch")