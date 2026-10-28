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