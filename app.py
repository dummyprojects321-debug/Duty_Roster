"""
Duty Roster  -  station-wise duty roster management.

    streamlit run app.py

Station user : Dashboard -> Roster -> Staff -> Tracker -> Setup (beats, shifts, time slots, signature)
Admin        : Overview -> Consolidated (all stations, PDF / Excel) -> Tracker -> Admin tools

All data rules live in core.py; this file is only the screen.
"""

import html
import time as _time
from datetime import date, datetime, time, timedelta

import pandas as pd
import streamlit as st

import core
from core import nice

# One place to change if your Streamlit version uses `use_container_width=True` instead.
STRETCH = {"width": "stretch"}

st.set_page_config(page_title="Duty Roster", page_icon="🗓️", layout="wide", initial_sidebar_state="expanded")

CSS = """
<style>
#MainMenu, footer {visibility: hidden;}
.block-container {padding-top: 2rem; padding-bottom: 3rem; max-width: 1280px;}
.hero {background: linear-gradient(120deg, #12263F 0%, #1D4E7A 60%, #1AA3A3 100%); color: #fff;
       border-radius: 16px; padding: 22px 28px; margin-bottom: 18px;}
.hero h1 {margin: 0; font-size: 1.7rem; font-weight: 700; color: #fff; padding: 0;}
.hero p {margin: 4px 0 0; opacity: .85; font-size: .95rem;}
[data-testid="stMetric"] {background: rgba(127,127,127,.08); border: 1px solid rgba(127,127,127,.18);
       border-radius: 14px; padding: 12px 16px;}
[data-testid="stMetricLabel"] p {font-size: .8rem; text-transform: uppercase; letter-spacing: .04em; opacity: .75;}
.stButton > button, .stDownloadButton > button, [data-testid="stFormSubmitButton"] > button {border-radius: 10px; font-weight: 600;}
.stTabs [data-baseweb="tab"] {font-weight: 600;}
.chip {display: inline-block; padding: 3px 12px; border-radius: 999px; font-size: .8rem; font-weight: 700;}
.chip.final {background: #DDF3E4; color: #17653A;} .chip.draft {background: #FFF3D1; color: #8A5200;}
.chip.none {background: #FDE4E4; color: #B91C1C;}
.login-card {text-align: center; padding: 8px 0 18px;}
.login-card h1 {font-size: 1.9rem; margin-bottom: 0;}
.login-card p {opacity: .7;}
</style>
"""

NAV_STATION = ["🏠 Dashboard", "📝 Roster", "👥 Staff", "🔎 Tracker", "⚙️ Setup"]
NAV_ADMIN = ["🏠 Overview", "📄 Consolidated", "🔎 Tracker", "🛠️ Admin tools"]
STATE_TXT = {"final": "✅ Final", "draft": "🟡 Draft", "none": "🔴 Not created"}


# =============================================================================
# small UI helpers
# =============================================================================
@st.cache_resource(show_spinner=False)
def _boot():
    return core.init_db()


@st.cache_data(show_spinner="Preparing PDF…", max_entries=40)
def _pdf(iso: str, stations: tuple, cover: bool, fingerprint: str) -> bytes:
    return core.build_pdf(iso, list(stations), cover)


@st.cache_data(show_spinner=False, max_entries=20)
def _xlsx(iso: str, stations: tuple, fingerprint: str) -> bytes:
    return core.build_excel(iso, list(stations))


def flash(msg: str, kind: str = "success"):
    st.session_state.setdefault("_flash", []).append((kind, msg))


def show_flash():
    for kind, msg in st.session_state.pop("_flash", []):
        if kind == "success":
            st.toast(msg, icon="✅")
        else:
            getattr(st, kind)(msg)


def bump():
    st.session_state["ver"] = st.session_state.get("ver", 0) + 1


def goto(page: str):
    st.session_state["nav"] = page


def hero(title: str, sub: str = ""):
    st.markdown(f'<div class="hero"><h1>{html.escape(title)}</h1><p>{html.escape(sub)}</p></div>',
                unsafe_allow_html=True)


def short_date(iso: str) -> str:
    return datetime.strptime(iso, "%Y-%m-%d").strftime("%a %d %b %Y")


def chip(state: str) -> str:
    label = {"final": "FINAL", "draft": "DRAFT", "none": "NOT CREATED"}[state]
    return f'<span class="chip {state}">{label}</span>'


def stations_for(user) -> list:
    return [user["station"]] if user["role"] == "station" else core.list_stations()


# =============================================================================
# login / password
# =============================================================================
def login_view():
    _, mid, _ = st.columns([1, 1.3, 1])
    with mid:
        st.markdown('<div class="login-card"><h1>🗓️ Duty Roster</h1>'
                    f'<p>{html.escape(core.ORG_NAME)}</p></div>', unsafe_allow_html=True)
        who = st.radio("Sign in as", ["Station", "Admin"], horizontal=True)
        with st.form("login"):
            if who == "Station":
                username = st.selectbox("Station", core.list_stations(), format_func=nice)
            else:
                username = core.ADMIN_USER
                st.text_input("Username", value=username, disabled=True)
            pw = st.text_input("Password", type="password")
            go = st.form_submit_button("Sign in", type="primary", **STRETCH)
        if go:
            user = core.authenticate(username, pw)
            if user:
                st.session_state["user"] = user
                st.rerun()
            _time.sleep(0.8)
            st.error("Wrong password.")
        st.caption("First time? Ask your admin for the starting password - you will be asked to change it.")


def password_form(user, key: str):
    with st.form(key, clear_on_submit=True):
        cur = st.text_input("Current password", type="password")
        new = st.text_input("New password (min 6 characters)", type="password")
        again = st.text_input("Repeat new password", type="password")
        go = st.form_submit_button("Change password", type="primary")
    if not go:
        return False
    defaults = {core.DEFAULT_ADMIN_PW, *core.DEFAULT_STATIONS.values()}
    if not core.authenticate(user["username"], cur):
        st.error("Current password is wrong.")
    elif len(new) < 6:
        st.error("New password must be at least 6 characters.")
    elif new != again:
        st.error("The two new passwords do not match.")
    elif new in defaults or new == cur:
        st.error("Choose a password different from the starting / current one.")
    else:
        core.change_password(user["username"], new)
        return True
    return False


def force_password_view(user):
    _, mid, _ = st.columns([1, 1.3, 1])
    with mid:
        st.markdown('<div class="login-card"><h1>🔐 Set your own password</h1>'
                    '<p>You are using a starting password. Choose a new one to continue.</p></div>',
                    unsafe_allow_html=True)
        if password_form(user, "force_pw"):
            user["must_change"] = False
            flash("Password changed.")
            st.rerun()


# =============================================================================
# sidebar + router
# =============================================================================
def sidebar(user) -> str:
    st.session_state.setdefault("work_date", date.today())

    def _set(delta):
        st.session_state["work_date"] = date.today() + timedelta(days=delta)

    with st.sidebar:
        st.markdown("### 🗓️ Duty Roster")
        who = "Administrator" if user["role"] == "admin" else f"{nice(user['station'])} station"
        st.caption(f"Signed in: **{who}**")
        st.date_input("Roster date", key="work_date", format="DD/MM/YYYY")
        c1, c2 = st.columns(2)
        c1.button("Today", on_click=_set, args=(0,), **STRETCH)
        c2.button("Tomorrow", on_click=_set, args=(1,), **STRETCH)
        pages = NAV_ADMIN if user["role"] == "admin" else NAV_STATION
        if st.session_state.get("nav") not in pages:
            st.session_state["nav"] = pages[0]
        st.radio("Menu", pages, key="nav", label_visibility="collapsed")
        st.divider()
        if st.button("Sign out", **STRETCH):
            st.session_state.clear()
            st.rerun()
    return st.session_state["work_date"].isoformat()


# =============================================================================
# STATION: dashboard
# =============================================================================
def page_dashboard(user, iso):
    stn = user["station"]
    hero(nice(stn), core.long_date(iso))
    s = core.day_summary(iso, [stn]).iloc[0]
    cols = st.columns(5)
    cols[0].metric("Active staff", int(s.staff))
    cols[1].metric("On duty", int(s.on_duty))
    cols[2].metric("Leave / off", int(s.leave_off))
    cols[3].metric("Not assigned", int(s.unassigned))
    cols[4].metric("Beats covered", int(s.beats))
    st.markdown(f"Roster status: {chip(s.state)}", unsafe_allow_html=True)

    if s.on_duty + s.leave_off == 0:
        st.info("No roster for this date yet. Build it in a few clicks - you can copy it from a previous day.")
        st.button("Create roster for this date →", type="primary", on_click=goto, args=("📝 Roster",))
        return
    left, right = st.columns([3, 2])
    with left:
        st.markdown("#### Beat coverage")
        cov = core.beat_coverage(stn, iso)
        cov["status"] = cov["people"].map(lambda n: "✅ covered" if n else "⚠️ nobody")
        cov.columns = ["Beat / area", "People", "Status"]
        st.dataframe(cov, hide_index=True, **STRETCH, height=min(40 + 35 * len(cov), 420))
    with right:
        rows = core.roster_rows(iso, [stn])
        duty = rows[rows["status"] == "Duty"]
        st.markdown("#### People by shift")
        if duty.empty:
            st.caption("No duty entries.")
        else:
            st.bar_chart(duty["shift"].replace("", "Not set").value_counts())
        grid = core.roster_grid(stn, iso)
        free = grid[(grid.status == "Duty") & (grid.beat == "") & (grid.shift == "") & (grid.slot == "")]
        if len(free):
            with st.expander(f"🕘 Not assigned yet ({len(free)})"):
                st.dataframe(free[["name", "designation", "belt_no"]].rename(columns=str.title),
                             hide_index=True, **STRETCH)
        part = rows[(rows["status"] == "Duty") & ((rows["beat"] == "") | (rows["shift"] == "") | (rows["slot"] == ""))]
        if len(part):
            st.warning(f"{len(part)} duty entr{'y is' if len(part) == 1 else 'ies are'} missing a beat, shift or time slot.")


# =============================================================================
# STATION: roster builder
# =============================================================================
def page_roster(user, iso):
    stn, by = user["station"], user["username"]
    ver = st.session_state.get("ver", 0)
    hero("Duty roster", f"{nice(stn)}  ·  {core.long_date(iso)}")
    staff = core.staff_df(stn)
    beats, shifts, slots = core.get_list("beat", stn), core.get_list("shift", stn), core.get_slots(stn)
    if staff.empty:
        st.warning("Add your staff first.")
        st.button("Go to Staff →", type="primary", on_click=goto, args=("👥 Staff",))
        return
    if not beats or not slots:
        st.warning("Add at least one beat and one time slot first (Setup page).")
        st.button("Go to Setup →", type="primary", on_click=goto, args=("⚙️ Setup",))
        return

    build, preview = st.tabs(["📝 Build roster", "👁️ Preview, sign & download"])
    grid = core.roster_grid(stn, iso)
    placed = int(((grid.status != "Duty") | (grid.beat != "") | (grid.shift != "") | (grid.slot != "")).sum())

    with build:
        st.progress(placed / max(len(grid), 1), text=f"{placed} of {len(grid)} people placed")

        # 1 - copy a previous day
        prev = core.roster_dates(stn, before=iso)
        with st.expander("⚡ Start from a previous roster (fastest)", expanded=(placed == 0 and bool(prev))):
            if prev:
                c1, c2, c3 = st.columns([2, 2, 1])
                src = c1.selectbox("Copy from", prev, format_func=short_date, key=f"cp_src_{ver}")
                rep = c2.checkbox("Replace what is already here", value=True, key=f"cp_rep_{ver}")
                c3.write("")
                if c3.button("Copy", type="primary", key=f"cp_go_{ver}"):
                    n = core.copy_roster(stn, src, iso, by, replace=rep)
                    flash(f"Copied {n} entries from {short_date(src)}. Adjust anything that changed, then save.")
                    bump()
                    st.rerun()
            else:
                st.caption("No earlier rosters yet - use Quick assign below for the first one.")

        # 2 - quick assign
        with st.container(border=True):
            st.markdown("#### ⚡ Quick assign")
            st.caption("Pick several people, choose beat, shift and time once, press Assign. Type a new beat or shift name to create it.")
            k1, k2 = st.columns(2)
            only_free = k1.checkbox("Only people not placed yet", value=True, key=f"qa_free_{ver}")
            pool = grid[(grid.status == "Duty") & (grid.beat == "") & (grid.shift == "") & (grid.slot == "")] if only_free else grid
            pick_all = k2.checkbox(f"Select everyone shown ({len(pool)})", value=False, key=f"qa_all_{ver}")
            labels = {int(r.staff_id): f"{r.name}  ·  {r.belt_no or '—'}  ·  {r.designation or '—'}" for r in pool.itertuples()}
            with st.form(f"qa_{ver}_{only_free}_{pick_all}"):
                people = st.multiselect("People", list(labels), default=list(labels) if pick_all else [],
                                        format_func=labels.get, placeholder="Choose people")
                a, b, c = st.columns(3)
                beat = a.selectbox("Beat / area", beats, index=None, accept_new_options=True, placeholder="Select or type")
                shift = b.selectbox("Shift", shifts, index=None, accept_new_options=True, placeholder="Select or type")
                slot = c.selectbox("Time slot", slots, index=None, placeholder="Select")
                go = st.form_submit_button("Assign to selected people", type="primary")
            if go:
                if not people:
                    st.warning("Choose at least one person.")
                elif not beat:
                    st.warning("Choose a beat / area.")
                else:
                    n = core.assign(stn, iso, people, beat, shift, slot, by)
                    flash(f"{n} assigned to {beat}.")
                    if not shift or not slot:
                        flash("Shift or time slot was left empty - fill it in the table below.", "warning")
                    bump()
                    st.rerun()

        # 3 - per-person grid
        st.markdown("#### Fine-tune each person")
        st.caption("Change any dropdown directly in the row, then press **Save roster**. "
                   "Leave Beat, Shift and Time empty to mark someone as not assigned. Use Status for leave / off days.")
        g = grid.copy()
        for col in ("beat", "shift", "slot"):
            g[col] = g[col].astype(object).where(g[col] != "", None)
        keep = lambda opts, cur: opts + [x for x in cur.dropna().unique() if x not in opts]  # noqa: E731
        edited = st.data_editor(
            g, key=f"grid_{stn}_{iso}_{ver}", hide_index=True, num_rows="fixed", **STRETCH,
            height=min(38 * len(g) + 40, 680),
            column_order=["name", "designation", "belt_no", "status", "beat", "shift", "slot", "remarks"],
            disabled=["name", "designation", "belt_no"],
            column_config={
                "name": st.column_config.TextColumn("Name", width="medium"),
                "designation": st.column_config.TextColumn("Designation", width="medium"),
                "belt_no": st.column_config.TextColumn("Belt No.", width="small"),
                "status": st.column_config.SelectboxColumn("Status", options=core.STATUSES, required=True, width="small"),
                "beat": st.column_config.SelectboxColumn("Beat / area", options=keep(beats, g["beat"]), width="large"),
                "shift": st.column_config.SelectboxColumn("Shift", options=keep(shifts, g["shift"]), width="small"),
                "slot": st.column_config.SelectboxColumn("Time slot", options=keep(slots, g["slot"]), width="medium"),
                "remarks": st.column_config.TextColumn("Remarks", width="medium"),
            })
        s1, s2 = st.columns([1, 3])
        if s1.button("💾 Save roster", type="primary", **STRETCH, key=f"save_{ver}"):
            res = core.save_roster(stn, iso, edited, by)
            flash(f"Saved - {res['saved']} entries." + (f" {res['cleared']} cleared." if res["cleared"] else ""))
            if res["warnings"]:
                flash("Incomplete entries: " + "; ".join(res["warnings"][:8]) + ("…" if len(res["warnings"]) > 8 else ""), "warning")
            bump()
            st.rerun()
        with s2.expander("🗑️ Clear this date's roster"):
            if st.checkbox("Yes, remove every entry for this date", key=f"clr_ok_{ver}") and st.button("Clear roster", key=f"clr_{ver}"):
                core.clear_roster(stn, iso, by)
                flash("Roster cleared.")
                bump()
                st.rerun()

    with preview:
        rows = core.roster_rows(iso, [stn])
        if rows.empty:
            st.info("Nothing to preview yet. Build the roster first.")
            return
        final = core.is_final(stn, iso)
        duty = rows[rows["status"] == "Duty"]
        off = rows[rows["status"] != "Duty"]
        incomplete = duty[(duty["beat"] == "") | (duty["shift"] == "") | (duty["slot"] == "")]
        free = int(grid[(grid.status == "Duty") & (grid.beat == "") & (grid.shift == "") & (grid.slot == "")].shape[0])

        top = st.columns(4)
        top[0].metric("On duty", len(duty))
        top[1].metric("Leave / off", len(off))
        top[2].metric("Not placed", free)
        top[3].metric("Incomplete", len(incomplete))
        if len(incomplete):
            st.warning("Some duty entries are missing a beat, shift or time slot: " + ", ".join(incomplete["name"].head(8)))
        if free:
            st.info(f"{free} active people have no duty yet. That is fine if they are not on the roster today.")

        # signature
        sig = core.get_signature(stn)
        with st.container(border=True):
            c1, c2 = st.columns([2, 3])
            with c1:
                st.markdown("**Signature on the PDF**")
                st.write(f"**{sig['label']}**" + (f" — {sig['officer']}" if sig["officer"] else ""))
                st.button("Edit signature →", on_click=goto, args=("⚙️ Setup",), key=f"sig_go_{ver}")
            with c2:
                if sig["image"]:
                    st.image(sig["image"], width=200)
                else:
                    st.caption("No signature image saved. The PDF leaves a blank space to sign by hand.")

        # status + downloads
        st.markdown(f"Status: {chip('final' if final else 'draft')}", unsafe_allow_html=True)
        d1, d2, d3 = st.columns(3)
        fp = core.fingerprint(iso, [stn])
        fname = f"Duty_Roster_{stn.replace(' ', '_')}_{iso}"
        d1.download_button("⬇️ Download PDF", _pdf(iso, (stn,), False, fp), f"{fname}.pdf", "application/pdf",
                           type="primary", **STRETCH)
        d2.download_button("⬇️ Download Excel", _xlsx(iso, (stn,), fp), f"{fname}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", **STRETCH)
        if final:
            if d3.button("↩️ Reopen as draft", **STRETCH):
                core.set_final(stn, iso, False, by)
                flash("Roster reopened.")
                st.rerun()
        elif d3.button("✅ Mark as final & submit", **STRETCH):
            core.set_final(stn, iso, True, by)
            flash("Roster marked final - the admin can see it.")
            st.rerun()

        st.markdown("#### Roster preview")
        for shift, gdf in duty.groupby(duty["shift"].replace("", "Shift not set"), sort=False):
            st.markdown(f"**{shift}**  ·  {len(gdf)} staff")
            st.dataframe(gdf[["name", "designation", "belt_no", "contact", "beat", "slot", "remarks"]]
                         .rename(columns={"belt_no": "belt no", "slot": "time"}).rename(columns=str.title),
                         hide_index=True, **STRETCH)
        if len(off):
            st.markdown(f"**Not on duty**  ·  {len(off)}")
            st.dataframe(off[["name", "designation", "belt_no", "status", "remarks"]].rename(columns=str.title),
                         hide_index=True, **STRETCH)


# =============================================================================
# STATION: staff
# =============================================================================
def page_staff(user, iso):
    stn, by = user["station"], user["username"]
    ver = st.session_state.get("ver", 0)
    hero("Staff", f"{nice(stn)}  ·  add people one by one or import a whole list")
    lst, add, imp = st.tabs(["👥 Staff list", "➕ Add person", "📥 Bulk import"])

    with lst:
        c1, c2 = st.columns([3, 1])
        q = c1.text_input("Search", placeholder="🔎  name, belt no., designation, contact", label_visibility="collapsed")
        inactive = c2.checkbox("Show inactive", value=False)
        df = core.staff_df(stn, active_only=not inactive)
        if q:
            m = df[["name", "belt_no", "designation", "contact"]].apply(
                lambda col: col.str.contains(q, case=False, na=False, regex=False)).any(axis=1)
            df = df[m]
        st.caption(f"{len(df)} shown. Edit any cell, tick/untick Active, then press Save. "
                   "Un-ticking Active keeps the history but hides the person from new rosters.")
        edited = st.data_editor(
            df, key=f"staff_{stn}_{ver}_{q}_{inactive}", hide_index=True, num_rows="fixed", **STRETCH,
            height=min(38 * max(len(df), 1) + 40, 640), disabled=["id"],
            column_config={
                "id": None,
                "name": st.column_config.TextColumn("Name", required=True),
                "designation": st.column_config.TextColumn("Designation"),
                "belt_no": st.column_config.TextColumn("Belt No."),
                "contact": st.column_config.TextColumn("Contact", help="Format 03XX-XXXXXXX"),
                "active": st.column_config.CheckboxColumn("Active"),
            })
        if st.button("💾 Save changes", type="primary", key=f"staff_save_{ver}"):
            n, errs = core.update_staff(stn, edited, by)
            flash(f"{n} person(s) updated." if n else "No changes to save.", "success" if n else "info")
            for e in errs:
                flash(e, "warning")
            bump()
            st.rerun()

    with add:
        with st.form(f"add_staff_{ver}", clear_on_submit=True):
            a, b = st.columns(2)
            name = a.text_input("Full name *")
            des = b.selectbox("Designation", core.designations(), index=None, accept_new_options=True,
                              placeholder="Select or type")
            c, d = st.columns(2)
            belt = c.text_input("Belt number")
            phone = d.text_input("Contact number", placeholder="0300-1234567")
            go = st.form_submit_button("Add person", type="primary")
        if go:
            ok, msg = core.add_staff(stn, name, des or "", belt, phone, by)
            if ok:
                flash(msg)
                bump()
                st.rerun()
            st.error(msg)

    with imp:
        st.markdown("Columns: **Name, Designation, Belt No, Contact No** - in that order or with those headings.")
        st.download_button("⬇️ Download Excel template", core.import_template_bytes(), "Staff_Import_Template.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        mode = st.radio("Source", ["Upload Excel / CSV", "Paste from a list"], horizontal=True)
        raw = None
        try:
            if mode.startswith("Upload"):
                up = st.file_uploader("Excel or CSV file", type=["xlsx", "xls", "csv"], key=f"up_{ver}")
                if up is not None:
                    raw = core.read_import_file(up)
            else:
                txt = st.text_area("One person per line: name, designation, belt no, contact (comma or tab separated)",
                                   height=160, key=f"paste_{ver}")
                if txt.strip():
                    raw = core.parse_pasted(txt)
        except Exception as exc:  # noqa: BLE001
            st.error(f"Could not read that: {exc}")
        if raw is not None and len(raw):
            prep = core.prepare_import(raw, stn)
            ok = prep["status"].isin(["New", "Check phone"])
            m1, m2, m3 = st.columns(3)
            m1.metric("Will be added", int(ok.sum()))
            m2.metric("Check phone", int((prep["status"] == "Check phone").sum()))
            m3.metric("Skipped", int((~ok).sum()))
            shown = prep.assign(status=prep["status"].map(
                lambda s: {"New": "🟢 New", "Check phone": "🟡 Check phone"}.get(s, f"⚪ Skip - {s}")))
            st.dataframe(shown, hide_index=True, **STRETCH, height=min(40 + 35 * len(shown), 420))
            if st.button(f"Import {int(ok.sum())} people", type="primary", disabled=not ok.any(), key=f"imp_{ver}"):
                n = core.import_staff(stn, prep, by)
                flash(f"Imported {n} people.")
                bump()
                st.rerun()


# =============================================================================
# TRACKER (station + admin)
# =============================================================================
def page_tracker(user, iso):
    hero("Tracker", "Where is each person deputed?")
    stations = stations_for(user)
    if user["role"] == "admin":
        stations = st.multiselect("Stations", stations, default=stations, format_func=nice) or stations
    by_beat, by_person = st.tabs([f"📍 Who is where on {short_date(iso)}", "🧑 Find a person"])

    with by_beat:
        rows = core.roster_rows(iso, stations)
        duty = rows[rows["status"] == "Duty"]
        if rows.empty:
            st.info("No roster for this date.")
        else:
            duty = duty.assign(person=duty["name"] + duty["belt_no"].map(lambda b: f" ({b})" if b else ""),
                               beat=duty["beat"].replace("", "— not set —"), shift=duty["shift"].replace("", "—"),
                               slot=duty["slot"].replace("", "—"), station=duty["station"].map(nice))
            tbl = (duty.groupby(["station", "beat", "shift", "slot"], sort=False)
                   .agg(People=("person", "count"), Staff=("person", ", ".join)).reset_index()
                   .rename(columns={"station": "Station", "beat": "Beat / area", "shift": "Shift", "slot": "Time"}))
            st.dataframe(tbl, hide_index=True, **STRETCH, height=min(40 + 36 * len(tbl), 640),
                         column_config={"Staff": st.column_config.TextColumn(width="large")})
            off = rows[rows["status"] != "Duty"]
            if len(off):
                with st.expander(f"Not on duty ({len(off)})"):
                    st.dataframe(off.assign(station=off["station"].map(nice))[["station", "name", "belt_no", "status", "remarks"]]
                                 .rename(columns=str.title), hide_index=True, **STRETCH)

    with by_person:
        a, b = st.columns([2, 2])
        q = a.text_input("Name, belt no. or contact", placeholder="🔎  e.g. Ali  or  SGT-2613")
        rng = b.date_input("Date range", value=(date.today() - timedelta(days=13), date.today()), format="DD/MM/YYYY")
        d1, d2 = (rng[0], rng[1]) if isinstance(rng, (tuple, list)) and len(rng) == 2 else (rng if not isinstance(rng, (tuple, list)) else rng[0],) * 2
        if not q.strip():
            st.caption("Type a name, belt number or phone number to see every duty of that person.")
        else:
            h = core.track(q.strip(), d1.isoformat(), d2.isoformat(), stations)
            if h.empty:
                st.info("No roster entries found for that search in this date range.")
            else:
                h["station"] = h["station"].map(nice)
                st.caption(f"{len(h)} entries · {h['name'].nunique()} person(s)")
                st.dataframe(h.rename(columns={"belt_no": "belt no"}).rename(columns=str.title),
                             hide_index=True, **STRETCH, height=min(40 + 36 * len(h), 600))


# =============================================================================
# STATION: setup
# =============================================================================
def list_manager(kind, noun, example, stn, by, ver):
    items = core.get_list(kind, stn)
    a, b = st.columns([3, 2])
    with a:
        with st.form(f"add_{kind}_{ver}", clear_on_submit=True):
            txt = st.text_area(f"Add {noun}s - one per line", height=130, placeholder=example)
            if st.form_submit_button(f"Add {noun}s", type="primary"):
                n = core.add_to_list(kind, stn, txt.splitlines(), by)
                flash(f"{n} {noun}(s) added." if n else "Nothing new to add.", "success" if n else "info")
                bump()
                st.rerun()
    with b:
        st.markdown(f"**Your {noun}s ({len(items)})**")
        if items:
            st.dataframe(pd.DataFrame({noun.title(): items}), hide_index=True, **STRETCH, height=min(40 + 35 * len(items), 300))
            rem = st.multiselect(f"Remove {noun}s", items, key=f"rm_{kind}_{ver}")
            if st.button("Remove selected", disabled=not rem, key=f"rmgo_{kind}_{ver}"):
                core.remove_from_list(kind, stn, rem, by)
                flash(f"{len(rem)} removed. Old rosters keep their text.")
                bump()
                st.rerun()
        else:
            st.caption("Nothing yet.")


def page_setup(user, iso):
    stn, by = user["station"], user["username"]
    ver = st.session_state.get("ver", 0)
    hero("Setup", f"{nice(stn)}  ·  the dropdown lists used while building the roster")
    t_beat, t_shift, t_slot, t_sig, t_pw = st.tabs(["📍 Beats / areas", "🔄 Shifts", "⏱️ Time slots", "✍️ Signature", "🔐 Password"])
    with t_beat:
        list_manager("beat", "beat", "Main Bazar\nGT Road\nRailway Station", stn, by, ver)
    with t_shift:
        list_manager("shift", "shift", "Morning\nEvening\nNight", stn, by, ver)
    with t_slot:
        a, b = st.columns([3, 2])
        with a:
            with st.form(f"add_slot_{ver}"):
                c1, c2 = st.columns(2)
                t1 = c1.time_input("From", value=time(9, 0), step=900)
                t2 = c2.time_input("To", value=time(17, 0), step=900)
                if st.form_submit_button("Add time slot", type="primary"):
                    ok, msg = core.add_slot(stn, t1, t2, by)
                    flash(msg, "success" if ok else "warning")
                    bump()
                    st.rerun()
        with b:
            slots = core.get_slots(stn)
            st.markdown(f"**Your time slots ({len(slots)})**")
            st.dataframe(pd.DataFrame({"Time slot": slots}), hide_index=True, **STRETCH)
            rem = st.multiselect("Remove time slots", slots, key=f"rm_slot_{ver}")
            if st.button("Remove selected", disabled=not rem, key=f"rmgo_slot_{ver}"):
                core.remove_slots(stn, rem, by)
                flash(f"{len(rem)} removed.")
                bump()
                st.rerun()
    with t_sig:
        sig = core.get_signature(stn)
        st.caption("Saved once, printed on every roster PDF of this station. Your photo is cleaned automatically: "
                   "use dark ink on white paper, a straight-on photo or scan.")
        a, b = st.columns([3, 2])
        with a:
            with st.form(f"sig_{ver}"):
                label = st.text_input("Title printed under the signature line", value=sig["label"])
                officer = st.text_input("Officer name (optional)", value=sig["officer"])
                up = st.file_uploader("Signature image (PNG / JPG) - leave empty to keep the saved one", type=["png", "jpg", "jpeg"])
                go = st.form_submit_button("Save signature", type="primary")
            if go:
                try:
                    png = core.process_signature(up.getvalue()) if up is not None else None
                    core.save_signature(stn, label, officer, by)
                    if png:
                        core.set_signature_image(stn, png, by)
                    flash("Signature saved.")
                    bump()
                    st.rerun()
                except ValueError as exc:
                    st.error(str(exc))
        with b:
            st.markdown("**Preview**")
            with st.container(border=True):
                if sig["image"]:
                    st.image(sig["image"], width=220)
                else:
                    st.write("\n\n\n")
                st.markdown(f"<hr style='margin:4px 0 6px;border-color:#64748B'><b>{html.escape(sig['label'])}</b>"
                            + (f"<br><span style='opacity:.7'>{html.escape(sig['officer'])}</span>" if sig["officer"] else ""),
                            unsafe_allow_html=True)
            if sig["image"] and st.button("Remove saved image", key=f"sig_rm_{ver}"):
                core.set_signature_image(stn, None, by)
                flash("Signature image removed.")
                bump()
                st.rerun()
    with t_pw:
        if password_form(user, f"pw_{ver}"):
            flash("Password changed.")
            st.rerun()


# =============================================================================
# ADMIN
# =============================================================================
def page_overview(user, iso):
    hero("All stations", core.long_date(iso))
    s = core.day_summary(iso)
    cols = st.columns(5)
    cols[0].metric("Stations final", f"{(s.state == 'final').sum()} / {len(s)}")
    cols[1].metric("Active staff", int(s.staff.sum()))
    cols[2].metric("On duty", int(s.on_duty.sum()))
    cols[3].metric("Leave / off", int(s.leave_off.sum()))
    cols[4].metric("Not assigned", int(s.unassigned.sum()))
    pending = s[s.state != "final"]
    if len(pending):
        st.warning("Not yet final: " + ", ".join(f"{nice(r.station)} ({'draft' if r.state == 'draft' else 'not created'})"
                                                  for r in pending.itertuples()))
    else:
        st.success("Every station has submitted its final roster.")
    tbl = s.assign(station=s.station.map(nice), state=s.state.map(STATE_TXT),
                   coverage=(s.on_duty / s.staff.clip(lower=1)))
    tbl = tbl[["station", "state", "staff", "on_duty", "leave_off", "unassigned", "beats", "coverage"]]
    tbl.columns = ["Station", "Roster", "Staff", "On duty", "Leave / off", "Not assigned", "Beats covered", "On duty %"]
    st.dataframe(tbl, hide_index=True, **STRETCH, column_config={
        "On duty %": st.column_config.ProgressColumn("On duty %", min_value=0, max_value=1, format="percent")})
    if s.on_duty.sum():
        st.markdown("#### On duty by station")
        st.bar_chart(s.set_index(s.station.map(nice))["on_duty"])


def page_consolidated(user, iso):
    hero("Consolidated roster", core.long_date(iso))
    all_st = core.list_stations()
    sel = st.multiselect("Stations", all_st, default=all_st, format_func=nice) or all_st
    rows = core.roster_rows(iso, sel)
    if rows.empty:
        st.info("No station has a roster for this date.")
        return
    fp = core.fingerprint(iso, sel)
    fname = f"Consolidated_Duty_Roster_{iso}"
    d1, d2, _ = st.columns([1, 1, 2])
    d1.download_button("⬇️ Download PDF", _pdf(iso, tuple(sel), len(sel) > 1, fp), f"{fname}.pdf", "application/pdf",
                       type="primary", **STRETCH)
    d2.download_button("⬇️ Download Excel", _xlsx(iso, tuple(sel), fp), f"{fname}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", **STRETCH)
    st.caption("The PDF starts with a station-wise summary, then each station's roster with its own saved SDEO signature.")
    f1, f2 = st.columns(2)
    shifts = sorted(set(rows["shift"]) - {""})
    fs = f1.multiselect("Filter by shift", shifts)
    q = f2.text_input("Search name / belt / beat", placeholder="🔎")
    v = rows.copy()
    if fs:
        v = v[v["shift"].isin(fs)]
    if q:
        v = v[v[["name", "belt_no", "beat", "designation"]].apply(lambda c: c.str.contains(q, case=False, na=False, regex=False)).any(axis=1)]
    v = v.assign(station=v["station"].map(nice))[["station", "name", "designation", "belt_no", "contact", "status",
                                                  "beat", "shift", "slot", "remarks"]]
    st.dataframe(v.rename(columns={"belt_no": "belt no", "slot": "time"}).rename(columns=str.title),
                 hide_index=True, **STRETCH, height=min(40 + 36 * len(v), 700))
    st.caption(f"{len(v)} entries shown.")


def page_admin_tools(user, iso):
    by = user["username"]
    ver = st.session_state.get("ver", 0)
    hero("Admin tools", "Stations, passwords, transfers and backup")
    t_st, t_tr, t_bk, t_log, t_pw = st.tabs(["🏢 Stations", "🔁 Transfer staff", "💾 Backup & restore", "🧾 Activity", "🔐 My password"])
    stations = core.list_stations()
    with t_st:
        s = core.day_summary(iso)[["station", "staff"]].assign(station=lambda d: d.station.map(nice))
        st.dataframe(s.rename(columns={"station": "Station", "staff": "Active staff"}), hide_index=True, **STRETCH)
        a, b = st.columns(2)
        with a:
            st.markdown("**Add a station**")
            with st.form(f"add_station_{ver}", clear_on_submit=True):
                name = st.text_input("Station name")
                pw = st.text_input("Starting password", type="password")
                if st.form_submit_button("Create station", type="primary"):
                    ok, msg = core.add_station(name, pw, by)
                    flash(msg, "success" if ok else "warning")
                    bump()
                    st.rerun()
        with b:
            st.markdown("**Reset a station password**")
            with st.form(f"reset_pw_{ver}", clear_on_submit=True):
                stn = st.selectbox("Station", stations, format_func=nice)
                pw = st.text_input("New starting password", type="password")
                if st.form_submit_button("Reset password"):
                    if len(pw) < 6:
                        st.error("At least 6 characters.")
                    else:
                        core.change_password(stn, pw, by=by)
                        flash(f"Password for {nice(stn)} reset. They must change it at next login.")
                        bump()
                        st.rerun()
    with t_tr:
        st.caption("Moves a person to another station. Their upcoming duties at the old station are removed; past rosters stay as they were.")
        src = st.selectbox("From station", stations, format_func=nice, key="tr_src")
        people = core.staff_df(src)
        if people.empty:
            st.info("No active staff at that station.")
        else:
            who = st.selectbox("Person", list(people["id"]), key="tr_who",
                               format_func=lambda i: "  ·  ".join(x for x in people.set_index("id").loc[i, ["name", "belt_no", "designation"]] if x))
            dst = st.selectbox("To station", [x for x in stations if x != src], format_func=nice, key="tr_dst")
            if st.button("Transfer", type="primary"):
                ok, msg = core.transfer_staff(int(who), dst, by)
                flash(msg, "success" if ok else "warning")
                bump()
                st.rerun()
    with t_bk:
        st.markdown("**Backup** - download regularly, especially if the app runs on a host that can reset its disk.")
        st.download_button("⬇️ Download backup", core.backup_bytes(),
                           f"duty_roster_backup_{datetime.now():%Y%m%d_%H%M}.db", "application/octet-stream")
        st.markdown("**Restore** - replaces ALL current data with the backup file.")
        up = st.file_uploader("Backup file (.db)", type=["db"], key=f"restore_{ver}")
        sure = st.checkbox("I understand this replaces everything and signs everyone out")
        if st.button("Restore backup", disabled=not (up and sure)):
            ok, msg = core.restore_backup(up.getvalue())
            if ok:
                st.cache_data.clear()
                st.session_state.clear()
                flash(msg)
                st.rerun()
            st.error(msg)
    with t_log:
        log = core.get_audit()
        st.dataframe(log, hide_index=True, **STRETCH, height=520)
    with t_pw:
        if password_form(user, f"pw_admin_{ver}"):
            flash("Password changed.")
            st.rerun()


# =============================================================================
# main
# =============================================================================
PAGES = {
    "🏠 Dashboard": page_dashboard, "📝 Roster": page_roster, "👥 Staff": page_staff,
    "🔎 Tracker": page_tracker, "⚙️ Setup": page_setup,
    "🏠 Overview": page_overview, "📄 Consolidated": page_consolidated, "🛠️ Admin tools": page_admin_tools,
}


def main():
    st.markdown(CSS, unsafe_allow_html=True)
    _boot()
    user = st.session_state.get("user")
    if not user:
        login_view()
        return
    if user.get("must_change"):
        force_password_view(user)
        return
    iso = sidebar(user)
    show_flash()
    PAGES[st.session_state["nav"]](user, iso)


main()