import streamlit as st
import pandas as pd
import time
import os
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="5K Race Timer", page_icon="🏃", layout="wide")

# ── File paths ────────────────────────────────────────────────────────────────
ROSTER_FILE   = "roster_backup.xlsx"
FINISHER_FILE = "finishers_backup.xlsx"

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700;800&family=Barlow:wght@400;500;600&display=swap');

html, body, [class*="css"] { font-family: 'Barlow', sans-serif; }
h1, h2, h3 { font-family: 'Barlow Condensed', sans-serif; font-weight: 800; letter-spacing: 0.02em; }
.stApp { background: #0d0d0d; color: #f0f0f0; }

.stTabs [data-baseweb="tab-list"] {
    gap: 0; background: #1a1a1a; border-radius: 0; border-bottom: 2px solid #ff4d00;
}
.stTabs [data-baseweb="tab"] {
    font-family: 'Barlow Condensed', sans-serif; font-size: 1.1rem; font-weight: 700;
    letter-spacing: 0.08em; text-transform: uppercase; padding: 14px 36px; color: #888; border-radius: 0;
}
.stTabs [aria-selected="true"] { background: #ff4d00 !important; color: #fff !important; }

.section-label {
    font-family: 'Barlow Condensed', sans-serif; font-size: 0.75rem; font-weight: 700;
    letter-spacing: 0.18em; text-transform: uppercase; color: #ff4d00; margin-bottom: 4px;
}
.section-title {
    font-family: 'Barlow Condensed', sans-serif; font-size: 1.9rem; font-weight: 800;
    color: #f0f0f0; margin-top: 0; margin-bottom: 20px;
    border-bottom: 1px solid #2a2a2a; padding-bottom: 10px;
}
.save-indicator {
    background: #0a1a0a; border: 1px solid #1a4a1a; border-left: 3px solid #00cc44;
    padding: 6px 14px; border-radius: 3px; font-family: 'Barlow Condensed', sans-serif;
    font-size: 0.85rem; color: #00cc44; letter-spacing: 0.05em; display: inline-block;
}
.race-status-stopped {
    background: #1a1a1a; border: 1px solid #333; border-left: 4px solid #555;
    padding: 16px 20px; border-radius: 4px; font-family: 'Barlow Condensed', sans-serif;
    font-size: 1.1rem; color: #888; letter-spacing: 0.05em; text-transform: uppercase;
}
.race-status-active {
    background: #0f1f0f; border: 1px solid #1a4a1a; border-left: 4px solid #00cc44;
    padding: 16px 20px; border-radius: 4px; font-family: 'Barlow Condensed', sans-serif;
    font-size: 1.1rem; color: #00cc44; letter-spacing: 0.05em; text-transform: uppercase;
}
.cat-card {
    background: #1a1a1a; border: 1px solid #2a2a2a; border-top: 3px solid #ff4d00;
    border-radius: 4px; padding: 14px 16px; margin-bottom: 16px;
}
.cat-card-title {
    font-family: 'Barlow Condensed', sans-serif; font-size: 1rem; font-weight: 700;
    letter-spacing: 0.1em; text-transform: uppercase; color: #ff4d00; margin-bottom: 10px;
}
.cat-card-female { border-top-color: #ff69b4; }
.cat-card-female .cat-card-title { color: #ff69b4; }
.podium-row {
    display: flex; align-items: center; gap: 10px; padding: 5px 0;
    border-bottom: 1px solid #222; font-size: 0.88rem;
}
.podium-row:last-child { border-bottom: none; }
.podium-medal { font-size: 1rem; width: 22px; text-align: center; flex-shrink: 0; }
.podium-name { font-weight: 600; color: #f0f0f0; flex: 1; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.podium-time { font-family: 'Barlow Condensed', sans-serif; font-size: 1rem; color: #aaa; font-weight: 600; }
.podium-empty { color: #444; font-style: italic; font-size: 0.82rem; padding: 3px 0; }
.race-status-ended {
    background: #1a0f0f; border: 1px solid #4a1a1a; border-left: 4px solid #cc2200;
    padding: 16px 20px; border-radius: 4px; font-family: 'Barlow Condensed', sans-serif;
    font-size: 1.1rem; color: #cc4422; letter-spacing: 0.05em; text-transform: uppercase;
}
.confirm-banner {
    background: #1a1200; border: 1px solid #4a3a00; border-left: 4px solid #ffaa00;
    padding: 14px 18px; border-radius: 4px; font-family: 'Barlow Condensed', sans-serif;
    font-size: 1rem; color: #ffcc44; letter-spacing: 0.04em;
}
div[data-testid="stButton"] button[kind="secondary"] {
    font-family: 'Barlow Condensed', sans-serif; font-weight: 700;
    letter-spacing: 0.06em; text-transform: uppercase;
}
div[data-testid="stButton"] button[kind="primary"] {
    background: #ff4d00; border: none; font-family: 'Barlow Condensed', sans-serif;
    font-weight: 800; letter-spacing: 0.1em; text-transform: uppercase; font-size: 1.1rem;
}
div[data-testid="stButton"] button[kind="primary"]:hover { background: #e64400; }
</style>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
for key, default in [
    ("runners", {}),
    ("race_started", False),
    ("race_stopped", False),
    ("race_start_time", None),
    ("race_start_wall", None),
    ("race_stop_wall", None),
    ("finishers", []),
    ("last_roster_save", None),
    ("last_finisher_save", None),
    ("confirm_stop", False),
    ("confirm_delete_idx", None),
    ("age_groups", [
        {"label": "65+",   "min": 65, "max": 999},
        {"label": "50-65", "min": 50, "max": 64},
        {"label": "35-50", "min": 35, "max": 49},
        {"label": "20-35", "min": 20, "max": 34},
        {"label": "13-20", "min": 13, "max": 19},
        {"label": "U13",   "min": 0,  "max": 12},
    ]),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ── Constants ─────────────────────────────────────────────────────────────────
DISPLAY_COLS = ["Place", "Bib", "Name", "Age", "Gender", "Elapsed Time", "Clock Time"]
MEDALS = ["🥇", "🥈", "🥉"]

# ── Shared style helpers ───────────────────────────────────────────────────────
def _thin_border():
    s = Side(style="thin", color="444444")
    return Border(left=s, right=s, top=s, bottom=s)

def _hdr_cell(ws, row, col, value, bg="FF4D00"):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = _thin_border()
    return cell

def _data_cell(ws, row, col, value, color="F0F0F0", bg="111111", center=True):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font      = Font(name="Calibri", size=10, color=color)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center" if center else "left", vertical="center")
    cell.border    = _thin_border()
    return cell

# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_elapsed(seconds: float) -> str:
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    return f"{h:02d}:{m:02d}:{s:02d}" if h else f"{m:02d}:{s:02d}"

def lookup(bib):
    return st.session_state.runners.get(str(bib).strip(), {})

def _parse_age(val):
    """Safely parse age from string or number, handling '32.0' style floats."""
    try:
        return int(float(str(val).strip()))
    except (ValueError, TypeError):
        return None

def _normalize_gender(val):
    """Normalize gender to 'Male' or 'Female', case-insensitive."""
    v = str(val).strip().lower()
    if v in ("male", "m"):
        return "Male"
    if v in ("female", "f"):
        return "Female"
    return None

def get_overall_winners():
    """Return the bib numbers of the single fastest male and female with full data."""
    male_winner = female_winner = None
    for f in st.session_state.finishers:
        if not f.get("Bib"):
            continue
        gender = _normalize_gender(f.get("Gender", ""))
        age    = _parse_age(f.get("Age", ""))
        if gender is None or age is None:
            continue
        secs = f.get("_secs", float("inf"))
        if gender == "Male":
            if male_winner is None or secs < male_winner.get("_secs", float("inf")):
                male_winner = f
        elif gender == "Female":
            if female_winner is None or secs < female_winner.get("_secs", float("inf")):
                female_winner = f
    excluded = set()
    if male_winner:
        excluded.add(male_winner["Bib"])
    if female_winner:
        excluded.add(female_winner["Bib"])
    return excluded

def category_top3(gender, age_min, age_max, excluded_bibs=None):
    """Get top 3 finishers for a category, excluding bibs in excluded_bibs set."""
    excluded_bibs = excluded_bibs or set()
    eligible = []
    for f in st.session_state.finishers:
        if not f.get("Bib") or f["Bib"] in excluded_bibs:
            continue
        f_gender = _normalize_gender(f.get("Gender", ""))
        f_age    = _parse_age(f.get("Age", ""))
        if f_gender is None or f_age is None:
            continue
        if f_gender != gender:
            continue
        if age_min <= f_age <= age_max:
            eligible.append(f)
    eligible.sort(key=lambda x: x.get("_secs", float("inf")))
    return eligible[:3]

def read_file_bytes(path):
    with open(path, "rb") as fh:
        return fh.read()

# ── Excel: save roster ────────────────────────────────────────────────────────
def save_roster_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Roster"

    headers    = ["Bib Number", "Name", "Age", "Gender"]
    col_widths = [14, 28, 8, 12]
    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        _hdr_cell(ws, 1, ci, h)
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 20
    ws.freeze_panes = "A2"

    for ri, (bib, v) in enumerate(
        sorted(st.session_state.runners.items(), key=lambda x: x[0].zfill(6)), 2
    ):
        bg = "1A1A1A" if ri % 2 == 0 else "131313"
        _data_cell(ws, ri, 1, bib,         bg=bg)
        _data_cell(ws, ri, 2, v["name"],   bg=bg, center=False)
        _data_cell(ws, ri, 3, v["age"],    bg=bg)
        _data_cell(ws, ri, 4, v["gender"], bg=bg)

    wb.save(ROSTER_FILE)
    st.session_state.last_roster_save = datetime.now().strftime("%I:%M:%S %p")

# ── Excel: save finishers ─────────────────────────────────────────────────────
def save_finishers_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Finishers"

    headers    = ["Place", "Bib", "Name", "Age", "Gender", "Elapsed Time", "Clock Time"]
    col_widths = [8, 10, 28, 8, 10, 14, 14]
    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        _hdr_cell(ws, 1, ci, h)
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.row_dimensions[1].height = 20
    ws.freeze_panes = "A2"

    place_bg = {1: "2A2000", 2: "1C1C1C", 3: "1A1000"}
    for ri, f in enumerate(st.session_state.finishers, 2):
        bg  = place_bg.get(f.get("Place", 99), "111111" if ri % 2 else "1A1A1A")
        row = [f.get(k, "") for k in ["Place","Bib","Name","Age","Gender","Elapsed Time","Clock Time"]]
        for ci, val in enumerate(row, 1):
            _data_cell(ws, ri, ci, val, bg=bg, center=(ci != 3))

    ws2 = wb.create_sheet("Info")
    ws2["A1"] = f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M:%S %p')}"
    ws2["A1"].font = Font(name="Calibri", size=9, color="888888")
    if st.session_state.race_start_wall:
        ws2["A2"] = f"Race start: {st.session_state.race_start_wall.strftime('%I:%M:%S %p')}"
        ws2["A2"].font = Font(name="Calibri", size=9, color="888888")

    wb.save(FINISHER_FILE)
    st.session_state.last_finisher_save = datetime.now().strftime("%I:%M:%S %p")

# ── Excel: leaderboard workbook ───────────────────────────────────────────────
def build_leaderboard_excel() -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    medal_labels = ["🥇 1st", "🥈 2nd", "🥉 3rd"]
    medal_bgs    = ["2A2000", "1C1C1C", "1A1000"]
    sub_headers  = ["Place", "Name", "Age", "Gender", "Time"]
    col_widths   = [12, 28, 8, 12, 14]

    excluded_bibs = get_overall_winners()

    def write_cat_block(ws, start_row, label, gender, age_min, age_max, excl=None):
        accent = "FF69B4" if gender.lower() == "female" else "FF4D00"
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=5)
        tc = ws.cell(row=start_row, column=1, value=label)
        tc.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=12)
        tc.fill      = PatternFill("solid", fgColor=accent)
        tc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[start_row].height = 18

        for ci, h in enumerate(sub_headers, 1):
            cell = ws.cell(row=start_row + 1, column=ci, value=h)
            cell.font      = Font(name="Calibri", bold=True, color="CCCCCC", size=10)
            cell.fill      = PatternFill("solid", fgColor="2A2A2A")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border    = _thin_border()

        top3 = category_top3(gender, age_min, age_max, excl)
        for rank in range(3):
            r = start_row + 2 + rank
            if rank < len(top3):
                f    = top3[rank]
                vals = [
                    medal_labels[rank],
                    f.get("Name") or f"Bib #{f.get('Bib','')}",
                    f.get("Age", ""),
                    f.get("Gender", ""),
                    f.get("Elapsed Time", ""),
                ]
            else:
                vals = [medal_labels[rank], "—", "", "", ""]
            bg = medal_bgs[rank]
            for ci, val in enumerate(vals, 1):
                _data_cell(ws, r, ci, val, bg=bg, center=(ci != 2))
            ws.row_dimensions[r].height = 16
        return start_row + 6

    def setup_sheet(ws, title_text):
        for ci, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.merge_cells("A1:E1")
        t = ws["A1"]
        t.value     = title_text
        t.font      = Font(name="Calibri", bold=True, color="FF4D00", size=16)
        t.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 26
        ws["A2"] = f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}"
        ws["A2"].font = Font(name="Calibri", size=9, color="666666")
        ws.row_dimensions[2].height = 14
        return 4

    # Overall sheet — overall winners ARE eligible here
    ws_ov = wb.create_sheet("Overall")
    row   = setup_sheet(ws_ov, "OVERALL RESULTS")
    row   = write_cat_block(ws_ov, row, "Male — Overall",   "Male",   0, 999, excl=set())
    row   = write_cat_block(ws_ov, row, "Female — Overall", "Female", 0, 999, excl=set())

    # Age group sheets — overall winners excluded
    for grp in st.session_state.age_groups:
        suffix, amin, amax = grp["label"], grp["min"], grp["max"]
        ws_ag = wb.create_sheet(f"Age {suffix}"[:31])  # Excel sheet name max 31 chars
        row   = setup_sheet(ws_ag, f"AGE GROUP: {suffix}")
        row   = write_cat_block(ws_ag, row, f"Male {suffix}",   "Male",   amin, amax, excl=excluded_bibs)
        row   = write_cat_block(ws_ag, row, f"Female {suffix}", "Female", amin, amax, excl=excluded_bibs)

    # All Finishers sheet
    ws_fin = wb.create_sheet("All Finishers")
    fin_headers    = ["Place", "Bib", "Name", "Age", "Gender", "Elapsed Time", "Clock Time"]
    fin_col_widths = [8, 10, 28, 8, 10, 14, 14]
    for ci, (h, w) in enumerate(zip(fin_headers, fin_col_widths), 1):
        _hdr_cell(ws_fin, 1, ci, h)
        ws_fin.column_dimensions[get_column_letter(ci)].width = w
    ws_fin.row_dimensions[1].height = 20
    ws_fin.freeze_panes = "A2"

    place_bg = {1: "2A2000", 2: "1C1C1C", 3: "1A1000"}
    for ri, f in enumerate(st.session_state.finishers, 2):
        bg       = place_bg.get(f.get("Place", 99), "111111" if ri % 2 else "1A1A1A")
        row_data = [f.get(k,"") for k in ["Place","Bib","Name","Age","Gender","Elapsed Time","Clock Time"]]
        for ci, val in enumerate(row_data, 1):
            _data_cell(ws_fin, ri, ci, val, bg=bg, center=(ci != 3))

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tab1, tab2 = st.tabs(["📋  Registration", "⏱  Race Timing"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — REGISTRATION
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-label">Setup</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Runner Registration</div>', unsafe_allow_html=True)

    # Import Excel
    st.markdown("#### 📂 Import from Excel")
    uploaded = st.file_uploader("Upload pre-registration spreadsheet (.xlsx)", type=["xlsx"])
    if uploaded:
        try:
            df_up = pd.read_excel(uploaded, dtype=str)
            df_up.columns = df_up.columns.str.strip()
            col_map = {}
            for col in df_up.columns:
                lc = col.lower().replace(" ", "")
                if "bib" in lc:                      col_map["bib"]    = col
                elif "name" in lc:                   col_map["name"]   = col
                elif "age" in lc:                    col_map["age"]    = col
                elif "gender" in lc or "sex" in lc:  col_map["gender"] = col

            missing = [k for k in ("bib","name","age","gender") if k not in col_map]
            if missing:
                st.error(f"Could not find columns: {', '.join(missing)}. Found: {list(df_up.columns)}")
            else:
                added = 0
                for _, row in df_up.iterrows():
                    bib = str(row[col_map["bib"]]).strip()
                    if bib and bib.lower() != "nan":
                        raw_gender = str(row[col_map["gender"]]).strip()
                        raw_age    = str(row[col_map["age"]]).strip()
                        st.session_state.runners[bib] = {
                            "name":   str(row[col_map["name"]]).strip(),
                            "age":    str(_parse_age(raw_age)) if _parse_age(raw_age) is not None else raw_age,
                            "gender": _normalize_gender(raw_gender) or raw_gender,
                        }
                        added += 1
                save_roster_excel()
                st.success(
                    f"✅ Loaded {added} runners — {len(st.session_state.runners)} total. "
                    f"Roster auto-saved to {ROSTER_FILE}."
                )
        except Exception as e:
            st.error(f"Error reading file: {e}")

    st.divider()

    # Add day-of registrant
    st.markdown("#### ➕ Add Day-Of Registrant")
    with st.form("add_runner", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns(4)
        nb = c1.text_input("Bib Number")
        nn = c2.text_input("Full Name")
        na = c3.text_input("Age")
        ng = c4.selectbox("Gender", ["Male", "Female"])
        add_btn = st.form_submit_button("Add Runner", type="primary", use_container_width=True)

    if add_btn:
        nb = nb.strip()
        if not nb:
            st.warning("Bib number is required.")
        else:
            verb = "Updated" if nb in st.session_state.runners else "Added"
            st.session_state.runners[nb] = {"name": nn, "age": na, "gender": ng}
            save_roster_excel()
            st.success(f"{verb} bib #{nb} — {nn}. Roster saved at {st.session_state.last_roster_save}.")

    st.divider()

    # Roster display + downloads
    total = len(st.session_state.runners)
    rh_col, rs_col = st.columns([3, 1])
    with rh_col:
        st.markdown(f"#### 📋 Full Roster ({total} runners)")
    with rs_col:
        if st.session_state.last_roster_save:
            st.markdown(
                f'<div class="save-indicator">💾 Saved {st.session_state.last_roster_save}</div>',
                unsafe_allow_html=True,
            )

    if total:
        roster_df = pd.DataFrame([
            {"Bib": b, "Name": v["name"], "Age": v["age"], "Gender": v["gender"]}
            for b, v in sorted(st.session_state.runners.items(), key=lambda x: x[0].zfill(6))
        ])
        st.dataframe(roster_df, use_container_width=True, hide_index=True, height=380)

        dl1, dl2 = st.columns(2)
        with dl1:
            file_data = read_file_bytes(ROSTER_FILE) if os.path.exists(ROSTER_FILE) else b""
            st.download_button(
                "⬇️ Download Roster (Excel)",
                data=file_data,
                file_name="roster_backup.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with dl2:
            st.download_button(
                "⬇️ Download Roster (CSV)",
                data=roster_df.to_csv(index=False).encode(),
                file_name="roster.csv",
                mime="text/csv",
                use_container_width=True,
            )
    else:
        st.info("No runners loaded yet. Import an Excel file or add runners manually above.")

    st.divider()

    # Age Group Manager
    st.markdown('<div class="section-label">Configuration</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Age Groups</div>', unsafe_allow_html=True)
    st.caption("Define the age groups that will appear in the leaderboard. Overall Male and Female categories are always included automatically.")

    if st.session_state.age_groups:
        hc1, hc2, hc3, hc4 = st.columns([2, 1, 1, 0.5])
        for h, lbl in zip([hc1, hc2, hc3, hc4], ["Group Label", "Min Age", "Max Age", ""]):
            h.markdown(
                f"<span style='font-size:0.78rem;color:#888;font-weight:700;"
                f"text-transform:uppercase;letter-spacing:0.1em'>{lbl}</span>",
                unsafe_allow_html=True,
            )
        for gi, grp in enumerate(st.session_state.age_groups):
            gc1, gc2, gc3, gc4 = st.columns([2, 1, 1, 0.5])
            gc1.markdown(f"<div style='padding-top:8px;font-weight:600;color:#f0f0f0'>{grp['label']}</div>", unsafe_allow_html=True)
            gc2.markdown(f"<div style='padding-top:8px;color:#ccc'>{grp['min']}</div>", unsafe_allow_html=True)
            max_display = grp['max'] if grp['max'] < 999 else "No limit"
            gc3.markdown(f"<div style='padding-top:8px;color:#ccc'>{max_display}</div>", unsafe_allow_html=True)
            if gc4.button("✕", key=f"del_grp_{gi}", help="Remove this age group"):
                st.session_state.age_groups.pop(gi)
                st.rerun()
    else:
        st.info("No age groups defined. The leaderboard will only show Overall results until groups are added.")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("#### ➕ Add Age Group")
    with st.form("add_age_group", clear_on_submit=True):
        ag1, ag2, ag3, ag4 = st.columns([2, 1, 1, 1])
        ag_label  = ag1.text_input("Group Label", placeholder='e.g. "20-35" or "Seniors"')
        ag_min    = ag2.number_input("Min Age", min_value=0, max_value=120, value=0, step=1)
        ag_max    = ag3.number_input("Max Age (999 = no limit)", min_value=0, max_value=999, value=19, step=1)
        ag_submit = ag4.form_submit_button("Add Group", type="primary", use_container_width=True)

    if ag_submit:
        ag_label = ag_label.strip()
        if not ag_label:
            st.warning("Please provide a label for the group.")
        elif ag_min > ag_max:
            st.warning("Min age must be less than or equal to max age.")
        elif ag_label in [g["label"] for g in st.session_state.age_groups]:
            st.warning(f'A group named "{ag_label}" already exists.')
        else:
            st.session_state.age_groups.append({"label": ag_label, "min": int(ag_min), "max": int(ag_max)})
            max_str = str(int(ag_max)) if int(ag_max) < 999 else "+"
            st.success(f"Added: {ag_label} (ages {int(ag_min)}–{max_str})")
            st.rerun()

    if st.session_state.age_groups:
        if st.button("↩ Reset to Default Age Groups", type="secondary"):
            st.session_state.age_groups = [
                {"label": "65+",   "min": 65, "max": 999},
                {"label": "50-65", "min": 50, "max": 64},
                {"label": "35-50", "min": 35, "max": 49},
                {"label": "20-35", "min": 20, "max": 34},
                {"label": "13-20", "min": 13, "max": 19},
                {"label": "U13",   "min": 0,  "max": 12},
            ]
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — RACE TIMING
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="section-label">Live</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Race Timing</div>', unsafe_allow_html=True)

    race_active  = st.session_state.race_started and not st.session_state.race_stopped
    race_stopped = st.session_state.race_stopped
    race_pending = st.session_state.race_started

    # ── Start / Stop / Status row ─────────────────────────────────────────────
    col_btn, col_stop, col_stat = st.columns([1, 1, 3])

    with col_btn:
        if not race_pending:
            if st.button("🚦 START RACE", type="primary", use_container_width=True):
                st.session_state.race_started    = True
                st.session_state.race_stopped    = False
                st.session_state.race_start_time = time.time()
                st.session_state.race_start_wall = datetime.now()
                st.session_state.race_stop_wall  = None
                st.session_state.confirm_stop    = False
                st.rerun()
        elif race_active:
            st.button("🚦 IN PROGRESS", disabled=True, use_container_width=True)
        else:
            st.button("🏁 RACE ENDED", disabled=True, use_container_width=True)

    with col_stop:
        if race_active:
            if not st.session_state.confirm_stop:
                if st.button("🛑 STOP RACE", type="secondary", use_container_width=True):
                    st.session_state.confirm_stop = True
                    st.rerun()
            else:
                if st.button("✅ Yes, stop it", use_container_width=True):
                    st.session_state.race_stopped   = True
                    st.session_state.race_stop_wall = datetime.now()
                    st.session_state.confirm_stop   = False
                    save_finishers_excel()
                    st.rerun()

    with col_stat:
        if st.session_state.confirm_stop:
            st.markdown(
                '<div class="confirm-banner">⚠️ &nbsp;Are you sure you want to stop the race? '
                'You can still edit results after stopping. &nbsp;'
                'Click <strong>Yes, stop it</strong> to confirm, or do nothing to cancel.</div>',
                unsafe_allow_html=True,
            )
        elif race_active:
            elapsed = time.time() - st.session_state.race_start_time
            wall    = st.session_state.race_start_wall.strftime("%I:%M:%S %p")
            st.markdown(
                f'<div class="race-status-active">🟢 &nbsp;Race started at {wall}'
                f' &nbsp;|&nbsp; Elapsed: {fmt_elapsed(elapsed)}</div>',
                unsafe_allow_html=True,
            )
        elif race_stopped:
            start_wall = st.session_state.race_start_wall.strftime("%I:%M:%S %p")
            stop_wall  = st.session_state.race_stop_wall.strftime("%I:%M:%S %p")
            total_secs = (st.session_state.race_stop_wall - st.session_state.race_start_wall).total_seconds()
            st.markdown(
                f'<div class="race-status-ended">🔴 &nbsp;Race ended at {stop_wall}'
                f' &nbsp;|&nbsp; Total duration: {fmt_elapsed(total_secs)}'
                f' &nbsp;|&nbsp; Started: {start_wall}</div>',
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                '<div class="race-status-stopped">⬜ Race not started</div>',
                unsafe_allow_html=True,
            )

    if st.session_state.confirm_stop and not race_active:
        st.session_state.confirm_stop = False

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Finish line button ────────────────────────────────────────────────────
    if race_active:
        if st.button("🏁  RUNNER CROSSED FINISH LINE", type="primary", use_container_width=True):
            now     = time.time()
            elapsed = now - st.session_state.race_start_time
            st.session_state.finishers.append({
                "Place":        len(st.session_state.finishers) + 1,
                "Bib":          "",
                "Name":         "",
                "Age":          "",
                "Gender":       "",
                "Elapsed Time": fmt_elapsed(elapsed),
                "Clock Time":   datetime.now().strftime("%I:%M:%S %p"),
                "_secs":        elapsed,
            })
            save_finishers_excel()
            st.rerun()

    st.divider()

    # ── Finisher table ────────────────────────────────────────────────────────
    if st.session_state.finishers:
        fh_col, fs_col = st.columns([3, 1])
        with fh_col:
            st.markdown(f"#### 🏆 Finishers — {len(st.session_state.finishers)} recorded")
        with fs_col:
            if st.session_state.last_finisher_save:
                st.markdown(
                    f'<div class="save-indicator">💾 Saved {st.session_state.last_finisher_save}</div>',
                    unsafe_allow_html=True,
                )

        st.caption("Type a bib number in the **Bib** column — Name, Age, and Gender auto-fill. Use ✕ to remove a row.")

        # ── Delete-confirmation banner ────────────────────────────────────────
        del_idx = st.session_state.confirm_delete_idx
        if del_idx is not None and del_idx < len(st.session_state.finishers):
            f     = st.session_state.finishers[del_idx]
            label = f.get("Name") or (f"Bib #{f['Bib']}" if f.get("Bib") else f"Place #{f['Place']}")
            st.markdown(
                f'<div class="confirm-banner">⚠️ &nbsp;Remove <strong>{label}</strong> '
                f'(Place #{f["Place"]}, {f.get("Elapsed Time","")}) from the finisher list?</div>',
                unsafe_allow_html=True,
            )
            conf_yes, conf_no, _ = st.columns([1, 1, 4])
            with conf_yes:
                if st.button("✅ Yes, remove", key="del_confirm_yes", use_container_width=True):
                    st.session_state.finishers.pop(del_idx)
                    for pi, fin in enumerate(st.session_state.finishers, 1):
                        fin["Place"] = pi
                    st.session_state.confirm_delete_idx = None
                    save_finishers_excel()
                    st.rerun()
            with conf_no:
                if st.button("✕ Cancel", key="del_confirm_no", use_container_width=True):
                    st.session_state.confirm_delete_idx = None
                    st.rerun()

        # ── Finisher rows ─────────────────────────────────────────────────────
        HDR = st.columns([0.5, 1, 2, 0.6, 1, 1.2, 1.2, 0.4])
        for h, label in zip(HDR, ["#", "Bib", "Name", "Age", "Gender", "Elapsed", "Clock", ""]):
            h.markdown(
                f"<span style='font-size:0.78rem;color:#888;font-weight:700;"
                f"text-transform:uppercase;letter-spacing:0.1em'>{label}</span>",
                unsafe_allow_html=True,
            )

        for i, fin in enumerate(st.session_state.finishers):
            cols = st.columns([0.5, 1, 2, 0.6, 1, 1.2, 1.2, 0.4])
            cols[0].markdown(
                f"<div style='padding-top:8px;color:#888;font-family:Barlow Condensed,sans-serif;"
                f"font-weight:700'>{fin['Place']}</div>", unsafe_allow_html=True
            )
            new_bib = cols[1].text_input(
                f"bib_{i}", value=fin.get("Bib",""), label_visibility="collapsed",
                key=f"bib_input_{i}", placeholder="Bib #"
            )
            cols[2].markdown(f"<div style='padding-top:8px;font-weight:600;color:#f0f0f0'>{fin.get('Name','') or '—'}</div>", unsafe_allow_html=True)
            cols[3].markdown(f"<div style='padding-top:8px;color:#ccc'>{fin.get('Age','') or '—'}</div>", unsafe_allow_html=True)
            cols[4].markdown(f"<div style='padding-top:8px;color:#ccc'>{fin.get('Gender','') or '—'}</div>", unsafe_allow_html=True)
            cols[5].markdown(f"<div style='padding-top:8px;font-family:Barlow Condensed,sans-serif;color:#aaa;font-weight:600'>{fin.get('Elapsed Time','')}</div>", unsafe_allow_html=True)
            cols[6].markdown(f"<div style='padding-top:8px;font-family:Barlow Condensed,sans-serif;color:#aaa'>{fin.get('Clock Time','')}</div>", unsafe_allow_html=True)

            btn_disabled = (del_idx == i)
            if cols[7].button("✕", key=f"del_{i}", disabled=btn_disabled, help="Remove this finisher"):
                st.session_state.confirm_delete_idx = i
                st.rerun()

            if str(new_bib).strip() != str(fin.get("Bib","")).strip():
                bib_clean = str(new_bib).strip()
                st.session_state.finishers[i]["Bib"] = bib_clean
                info = lookup(bib_clean)
                raw_gender = info.get("gender", "")
                norm_gender = _normalize_gender(raw_gender) or raw_gender
                raw_age = info.get("age", "")
                norm_age = str(_parse_age(raw_age)) if _parse_age(raw_age) is not None else raw_age
                st.session_state.finishers[i]["Name"]   = info.get("name", "")
                st.session_state.finishers[i]["Age"]    = norm_age
                st.session_state.finishers[i]["Gender"] = norm_gender
                save_finishers_excel()
                st.rerun()

        # ── Downloads ─────────────────────────────────────────────────────────
        st.divider()
        st.markdown("#### 📥 Download Results")
        dc1, dc2, dc3 = st.columns(3)
        with dc1:
            fin_data = read_file_bytes(FINISHER_FILE) if os.path.exists(FINISHER_FILE) else b""
            st.download_button(
                "⬇️ Finisher List (Excel)", data=fin_data,
                file_name="finishers_backup.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with dc2:
            export_df = pd.DataFrame(st.session_state.finishers)[DISPLAY_COLS]
            st.download_button(
                "⬇️ Finisher List (CSV)", data=export_df.to_csv(index=False).encode(),
                file_name="race_results.csv", mime="text/csv", use_container_width=True,
            )
        with dc3:
            st.download_button(
                "⬇️ Full Leaderboard (Excel)", data=build_leaderboard_excel(),
                file_name="leaderboard.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    else:
        st.info("No finishers recorded yet. Start the race and press the finish line button as runners cross.")

    st.divider()

    # ── Category leaderboards ─────────────────────────────────────────────────
    st.markdown('<div class="section-label">Live Standings</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Top 3 by Category</div>', unsafe_allow_html=True)

    # Compute the excluded bib set once — the overall male & female winners
    excluded_bibs = get_overall_winners()

    def render_card(label, gender, age_min, age_max, excl=None):
        top       = category_top3(gender, age_min, age_max, excl)
        extra_cls = "cat-card-female" if gender.lower() == "female" else ""
        rows_html = ""
        for rank, f in enumerate(top):
            name = f.get("Name") or f"Bib #{f.get('Bib','?')}"
            rows_html += (
                f'<div class="podium-row">'
                f'<span class="podium-medal">{MEDALS[rank]}</span>'
                f'<span class="podium-name">{name}</span>'
                f'<span class="podium-time">{f.get("Elapsed Time","")}</span>'
                f'</div>'
            )
        for _ in range(3 - len(top)):
            rows_html += '<div class="podium-row"><span class="podium-empty">— awaiting finishers —</span></div>'
        st.markdown(
            f'<div class="cat-card {extra_cls}"><div class="cat-card-title">{label}</div>{rows_html}</div>',
            unsafe_allow_html=True,
        )

    # Overall — no exclusions; the overall winner IS the category winner here
    st.markdown("##### Overall")
    ov1, ov2 = st.columns(2)
    with ov1: render_card("Male",   "Male",   0, 999, excl=set())
    with ov2: render_card("Female", "Female", 0, 999, excl=set())

    # Age groups — exclude the overall winners
    for grp in st.session_state.age_groups:
        suffix, amin, amax = grp["label"], grp["min"], grp["max"]
        st.markdown(f"##### Age {suffix}")
        gc1, gc2 = st.columns(2)
        with gc1: render_card(f"Male {suffix}",   "Male",   amin, amax, excl=excluded_bibs)
        with gc2: render_card(f"Female {suffix}", "Female", amin, amax, excl=excluded_bibs)

    # ── Reset ─────────────────────────────────────────────────────────────────
    st.divider()
    if st.button("🔄 Reset Race (clears timer & all finishers)", type="secondary"):
        st.session_state.race_started        = False
        st.session_state.race_stopped        = False
        st.session_state.race_start_time     = None
        st.session_state.race_start_wall     = None
        st.session_state.race_stop_wall      = None
        st.session_state.finishers           = []
        st.session_state.last_finisher_save  = None
        st.session_state.confirm_stop        = False
        st.session_state.confirm_delete_idx  = None
        if os.path.exists(FINISHER_FILE):
            os.remove(FINISHER_FILE)
        st.rerun()
