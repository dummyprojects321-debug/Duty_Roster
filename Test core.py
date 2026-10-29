"""Run with:  python -m pytest -q   (or just  python test_core.py)"""
import io
import os
import tempfile

import pandas as pd

_tmp = tempfile.mkdtemp()
os.environ["DUTY_DB_PATH"] = os.path.join(_tmp, "test.db")
import core  # noqa: E402

core.SEED_FILE = "/nonexistent.xlsx"  # tests start from an empty database
core.init_db()
ST = "KAMOKE"
D1, D2 = "2026-11-02", "2026-11-03"


def test_login_and_password():
    assert core.authenticate("ADMIN", "wrong") is None
    u = core.authenticate("admin", core.DEFAULT_ADMIN_PW)
    assert u["role"] == "admin" and u["must_change"]
    core.change_password("ADMIN", "NewPass123")
    assert not core.authenticate("ADMIN", "NewPass123")["must_change"]
    s = core.authenticate("kamoke", "kamoke2026")
    assert s["station"] == ST and s["must_change"]
    core.change_password(ST, "kamoke-new", by="ADMIN")          # admin reset -> must change again
    assert core.authenticate(ST, "kamoke-new")["must_change"]


def test_staff_add_edit_duplicates():
    assert core.add_staff(ST, "ali  raza", "sergeant", "sgt 1001", "3001234567", "t")[0]
    assert not core.add_staff("WAZIRABAD", "Other", "", "SGT1001", "", "t")[0]   # belt is global
    assert core.add_staff(ST, "No Belt", "", "", "", "t")[0]
    assert not core.add_staff(ST, "no belt", "", "", "", "t")[0]
    df = core.staff_df(ST)
    assert set(df["name"]) == {"Ali Raza", "No Belt"}
    r = df[df.name == "Ali Raza"].iloc[0]
    assert r.belt_no == "SGT1001" and r.contact == "0300-1234567" and r.designation == "SERGEANT"
    df.loc[df.name == "No Belt", "contact"] = "03111234567"
    n, err = core.update_staff(ST, df, "t")
    assert n == 1 and not err


def test_bulk_import():
    raw = pd.DataFrame({"Staff Name": ["A One", "B Two", "", "A One"], "Rank": ["DRIVER"] * 4,
                        "Belt Number": ["D-1", "SGT1001", "D-3", "D-1"], "Mobile": ["0300-1111111", "x", "0300-2222222", "1"]})
    prep = core.prepare_import(core.standardize_import(raw), ST)
    assert list(prep["status"]) == ["New", "Already exists", "Name missing", "Duplicate in file"]
    assert core.import_staff(ST, prep, "t") == 1
    # unknown headers -> positional
    pos = core.standardize_import(pd.DataFrame({"a": ["P Q"], "b": ["DRIVER"], "c": ["D-9"], "d": ["3451234567"]}))
    assert core.prepare_import(pos, ST).iloc[0]["contact"] == "0345-1234567"
    pasted = core.parse_pasted("Name\tDesignation\tBelt\tContact\nZed Y\tDRIVER\tD-10\t0300-9999999")
    assert len(pasted) == 1 and pasted.iloc[0]["name"] == "Zed Y"
    xl = core.import_template_bytes()
    assert len(core.read_import_file(io.BytesIO(xl))) == 2


def test_lists_and_slots():
    assert core.add_to_list("beat", ST, ["Main Bazar", "main bazar", " GT Road "], "t") == 2
    assert core.get_list("shift", ST) == core.DEFAULT_SHIFTS
    assert core.add_slot(ST, "09:00", "17:00", "t")[0] is False   # default already there
    assert core.add_slot(ST, "06:00", "14:00", "t")[0]
    assert core.get_slots(ST)[0].startswith("6:00 AM")
    assert core.remove_slots(ST, ["6:00 AM \u2013 2:00 PM"], "t") == 1


def test_roster_flow():
    grid = core.roster_grid(ST, D1)
    assert len(grid) == len(core.staff_df(ST)) and (grid["beat"] == "").all()
    ids = list(grid["staff_id"][:2])
    slot = core.get_slots(ST)[0]
    assert core.assign(ST, D1, ids, "New Beat", "Morning", slot, "t") == 2
    assert "New Beat" in core.get_list("beat", ST)
    grid = core.roster_grid(ST, D1)
    grid.loc[grid.index[2], ["status", "remarks"]] = ["Leave", "sick"]
    grid.loc[grid.index[0], "beat"] = ""                       # incomplete -> warning, still saved
    res = core.save_roster(ST, D1, grid, "t")
    assert res["saved"] == 3 and len(res["warnings"]) == 1 and res["changed"]
    assert not core.save_roster(ST, D1, core.roster_grid(ST, D1), "t")["changed"]   # idempotent
    core.set_final(ST, D1, True, "t")
    assert core.is_final(ST, D1)
    g = core.roster_grid(ST, D1); g.loc[g.index[1], "beat"] = "GT Road"
    core.save_roster(ST, D1, g, "t")
    assert not core.is_final(ST, D1)                           # editing re-opens as draft
    g = core.roster_grid(ST, D1); g[["status", "beat", "shift", "slot"]] = ["Duty", "", "", ""]
    assert core.save_roster(ST, D1, g, "t")["cleared"] == 3


def test_copy_summary_track():
    core.assign(ST, D1, list(core.staff_df(ST)["id"]), "Main Bazar", "Evening", core.get_slots(ST)[1], "t")
    assert core.copy_roster(ST, D1, D2, "t") == len(core.staff_df(ST))
    assert core.roster_dates(ST, before=D2) == [D1]
    s = core.day_summary(D2, [ST]).iloc[0]
    assert s.on_duty == s.staff and s.unassigned == 0 and s.state == "draft"
    t = core.track("ali", D1, D2, [ST])
    assert len(t) == 2 and set(t["beat"]) == {"Main Bazar"}
    cov = core.beat_coverage(ST, D2)
    assert cov.set_index("beat").loc["Main Bazar", "people"] == len(core.staff_df(ST))
    assert core.copy_roster(ST, D1, D2, "t", replace=False) == 0


def test_transfer():
    sid = int(core.staff_df(ST).iloc[0]["id"])
    assert core.transfer_staff(sid, "WAZIRABAD", "t")[0]
    assert sid not in set(core.staff_df(ST)["id"]) and sid in set(core.staff_df("WAZIRABAD")["id"])


def test_signature_and_pdf():
    from PIL import Image, ImageDraw
    im = Image.new("RGB", (400, 200), "white")
    ImageDraw.Draw(im).line([(30, 150), (150, 40), (250, 140), (370, 60)], fill=(20, 20, 160), width=6)
    b = io.BytesIO(); im.save(b, "JPEG")
    png = core.process_signature(b.getvalue())
    assert Image.open(io.BytesIO(png)).mode == "RGBA"
    blank = io.BytesIO(); Image.new("RGB", (50, 50), "white").save(blank, "PNG")
    try:
        core.process_signature(blank.getvalue()); assert False
    except ValueError:
        pass
    assert core.get_signature(ST)["label"] == "SDEO (KAMOKE)"
    core.save_signature(ST, "SDEO (Kamoke)", "Mr. Khan", "t")
    core.set_signature_image(ST, png, "t")
    sig = core.get_signature(ST)
    assert sig["label"] == "SDEO (Kamoke)" and sig["officer"] == "Mr. Khan" and sig["image"]
    pdf = core.build_pdf(D2, [ST])
    assert pdf[:4] == b"%PDF"
    allpdf = core.build_pdf(D2, core.list_stations(), cover=True)
    assert allpdf[:4] == b"%PDF" and len(allpdf) > 2000
    assert core.build_pdf("2030-01-01", [ST])[:4] == b"%PDF"          # empty day still prints
    f1 = core.fingerprint(D2, [ST])
    core.save_signature(ST, "SDEO (Kamoke City)", "Mr. Khan", "t")
    assert core.fingerprint(D2, [ST]) != f1
    assert core.build_excel(D2, core.list_stations())[:2] == b"PK"


def test_backup_restore():
    data = core.backup_bytes()
    core.add_staff(ST, "Temp Person", "", "TMP-1", "", "t")
    ok, _ = core.restore_backup(data)
    assert ok and "Temp Person" not in set(core.staff_df(ST)["name"])
    assert not core.restore_backup(b"garbage")[0]


def test_add_station():
    assert core.add_station("Sambrial", "abcdef", "ADMIN")[0]
    assert not core.add_station("sambrial", "abcdef", "ADMIN")[0]
    assert core.get_list("shift", "SAMBRIAL") == core.DEFAULT_SHIFTS


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn(); print("PASS", name)