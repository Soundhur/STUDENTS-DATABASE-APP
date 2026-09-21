import pandas as pd
import sqlite3
import streamlit as st
import re
import time
import difflib
import os
import shutil
from datetime import datetime

st.set_page_config(page_title="MCA Student Portal", page_icon=":material/school:", layout="wide")

st.markdown("""
<style>
/* Main Typography */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

/* Metric Cards Styling (Glassmorphism) */
div[data-testid="stMetric"] {
    background: rgba(255, 255, 255, 0.03);
    border: 1px solid rgba(255, 255, 255, 0.1);
    border-radius: 12px;
    padding: 15px 20px;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    backdrop-filter: blur(10px);
    transition: transform 0.2s ease-in-out, box-shadow 0.2s ease-in-out;
}
div[data-testid="stMetric"]:hover {
    transform: translateY(-5px);
    box-shadow: 0 8px 15px rgba(0, 0, 0, 0.2);
    border-color: rgba(255, 255, 255, 0.2);
}

/* Primary Button Hover Animations */
button[kind="primary"] {
    transition: transform 0.2s ease-in-out, box-shadow 0.2s ease-in-out !important;
}
button[kind="primary"]:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 4px 12px rgba(255, 75, 75, 0.4) !important;
}

/* Profile Card Styling */
.profile-card {
    background: rgba(255, 255, 255, 0.03);
    border: 1px solid rgba(255, 255, 255, 0.1);
    border-radius: 12px;
    padding: 20px;
    margin-bottom: 25px;
    backdrop-filter: blur(10px);
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
}
.profile-header {
    color: #ff4b4b;
    font-size: 1.3em;
    font-weight: 600;
    margin-bottom: 15px;
    border-bottom: 1px solid rgba(255, 255, 255, 0.1);
    padding-bottom: 8px;
    display: flex;
    align-items: center;
    gap: 8px;
}
.profile-data {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 15px;
}
.profile-item {
    display: flex;
    flex-direction: column;
}
.profile-label {
    font-size: 0.8em;
    color: #888;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 4px;
}
.profile-value {
    font-size: 1.05em;
    font-weight: 500;
    color: #eee;
}
</style>
""", unsafe_allow_html=True)

DB_NAME = "mca_students_umis.db"
EXCEL_FILE = "UMIS_API STUDENT LIST_2026-27.xlsx"

def validate_student_data(name, mobile, email, aadhaar):
    errors = []
    if not name.strip():
        errors.append("Student Name is required.")
    
    mobile_str = str(mobile).strip()
    if mobile_str and not re.match(r"^\d{10}$", mobile_str):
        errors.append("Mobile Number must be exactly 10 digits.")
        
    aadhaar_str = str(aadhaar).strip()
    if aadhaar_str and not re.match(r"^\d{12}$", aadhaar_str):
        errors.append("Aadhaar Number must be exactly 12 digits.")
        
    email_str = str(email).strip()
    if email_str and "@" not in email_str:
        errors.append("Email must be a valid format.")
        
    return errors

@st.cache_resource
def initialize_database():
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()

  # Check if table already exists
  cursor.execute(
      "SELECT name FROM sqlite_master WHERE type='table' AND name='students'"
  )
  table_exists = cursor.fetchone()

  if not table_exists:
    # Read the exact format from the uploaded Excel file
    df = pd.read_excel(EXCEL_FILE, sheet_name="Format")
    # Clean column names for SQLite compatibility
    df.columns = [
        "s_no",
        "admission_year",
        "tnea_code",
        "college_name",
        "student_name",
        "branch",
        "mobile_number",
        "email",
        "aadhaar",
        "umis_number",
        "citizenship",
        "passport",
        "transfer_readmission",
        "course_type",
        "year_of_study",
    ]
    # Fill NaN with empty strings for text storage
    df = df.fillna("")
    # Remove any empty placeholder rows from the Excel file
    df = df[df["student_name"] != ""]
    df.to_sql("students", conn, index=False, if_exists="replace")

  conn.close()


initialize_database()

st.title("🎓 MCA Student Details Portal (UMIS Format)")
# --- TIME MACHINE (SILENT BACKUP) ---
if "backup_completed" not in st.session_state:
    backup_dir = ".backups"
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir)
    if os.path.exists(DB_NAME):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(backup_dir, f"mca_students_{timestamp}.db")
        shutil.copy2(DB_NAME, backup_path)
    st.session_state.backup_completed = True

if "god_mode" not in st.session_state:
    st.session_state.god_mode = False

# --- OMNI-BAR ---
st.sidebar.markdown("### ⚡ Omni-Bar")
omni_input = st.sidebar.text_input("Execute Command (/help)", key="omni_input_key")

force_nav = None
if omni_input.startswith("/"):
    cmd = omni_input.lower().strip()
    if cmd == "/godmode":
        st.session_state.god_mode = not st.session_state.god_mode
        if st.session_state.god_mode:
            st.sidebar.success("God Mode Unlocked")
        time.sleep(0.5)
        st.rerun()
    elif cmd == "/audit":
        force_nav = "Data Quality Audit"
    elif cmd == "/detective":
        force_nav = "🕵️‍♂️ Data Detective"
    elif cmd.startswith("/teams "):
        try:
            size = int(cmd.split(" ")[1])
            st.session_state.team_size = size
            force_nav = "🤝 Team Builder"
        except:
            pass
    elif cmd == "/clean":
        conn = sqlite3.connect(DB_NAME)
        conn.execute("""
            UPDATE students 
            SET student_name = UPPER(TRIM(student_name)), 
                branch = UPPER(TRIM(branch)),
                email = LOWER(TRIM(email)),
                mobile_number = REPLACE(mobile_number, ' ', ''),
                aadhaar = REPLACE(aadhaar, ' ', '')
        """)
        conn.commit()
        st.sidebar.success("Database Magically Formatted!")
        time.sleep(1)
        st.rerun()
    elif cmd == "/help":
        st.sidebar.info("Commands: /audit, /detective, /clean, /teams [number], /godmode")

st.sidebar.title("Navigation")
nav_options = ["Advisor Master View", "Advanced Analytics", "Data Quality Audit", "🕵️‍♂️ Data Detective", "🤝 Team Builder", "Update Student Portal", "Add New Student", "Delete Student", "Bulk Upload"]

if st.session_state.god_mode:
    nav_options.append("👁️ God Mode (SQL)")

default_idx = nav_options.index(force_nav) if force_nav in nav_options else 0
nav_option = st.sidebar.radio("Go to", nav_options, index=default_idx)

conn = sqlite3.connect(DB_NAME)

if nav_option == "Advisor Master View":
  df_all = pd.read_sql("SELECT * FROM students", conn)
  
  # Clean up .0 decimals that pandas adds to numeric columns with missing values
  for col in ["tnea_code", "mobile_number", "aadhaar", "umis_number"]:
      if col in df_all.columns:
          df_all[col] = df_all[col].astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '')
          
  # --- DASHBOARD METRICS ---
  st.subheader("📊 Dashboard Overview")
  m1, m2, m3, m4 = st.columns(4)
  m1.metric("Total Students", len(df_all))
  m2.metric("Total Branches", df_all["branch"].nunique() if not df_all.empty else 0)
  m3.metric("Total UG", len(df_all[df_all["course_type"].astype(str).str.upper() == "UG"]) if not df_all.empty else 0)
  m4.metric("Total PG", len(df_all[df_all["course_type"].astype(str).str.upper() == "PG"]) if not df_all.empty else 0)
  
  st.divider()
  
  # --- SEARCH & FILTERING ---
  st.subheader(f"📋 Master Database (Total: {len(df_all)})")
  
  col_search, col_filter = st.columns(2)
  with col_search:
      search_query = st.text_input("🔍 Smart Search (Name, UMIS, or College)", "")
  with col_filter:
      unique_branches = sorted([b for b in df_all["branch"].unique() if str(b).strip() != ""])
      branch_filter = st.selectbox("📁 Filter by Branch", ["All"] + unique_branches)
      
  # Apply filters
  filtered_df = df_all.copy()
  if search_query:
      search_query = str(search_query).lower()
      
      def fuzzy_match(row):
          # Exact UMIS match
          if search_query in str(row['umis_number']).lower():
              return True
          
          # Fuzzy Name match
          name = str(row['student_name']).lower()
          if search_query in name or difflib.SequenceMatcher(None, search_query, name).ratio() > 0.6:
              return True
              
          # Fuzzy College match
          college = str(row['college_name']).lower()
          if search_query in college or difflib.SequenceMatcher(None, search_query, college).ratio() > 0.7:
              return True
              
          return False
          
      filtered_df = filtered_df[filtered_df.apply(fuzzy_match, axis=1)]
      
  if branch_filter != "All":
      filtered_df = filtered_df[filtered_df["branch"] == branch_filter]
      
  # --- DATA VISUALIZATIONS ---
  if not filtered_df.empty:
      c1, c2 = st.columns(2)
      with c1:
          st.write("**Student Distribution by Branch**")
          branch_counts = filtered_df["branch"].value_counts()
          st.bar_chart(branch_counts, color="#ff4b4b", height=250)
      with c2:
          st.write("**Enrollment by Admission Year**")
          year_counts = filtered_df["admission_year"].value_counts()
          st.bar_chart(year_counts, color="#1e88e5", height=250)
          
  # Clean up .0 from numeric columns for display
  cols_to_clean = ["tnea_code", "mobile_number", "aadhaar", "umis_number", "year_of_study"]
  for col in cols_to_clean:
      if col in filtered_df.columns:
          filtered_df[col] = filtered_df[col].astype(str).str.replace(r"\.0$", "", regex=True).replace("nan", "")
          
  # Make emails clickable links that specifically open Gmail
  if "email" in filtered_df.columns:
      filtered_df["email"] = filtered_df["email"].apply(lambda x: f"https://mail.google.com/mail/?view=cm&fs=1&to={x}" if str(x).strip() else "")
          
  # Hide unnecessary columns to reduce clutter
  cols_to_hide = ["citizenship", "passport", "transfer_readmission"]
  filtered_df = filtered_df.drop(columns=[c for c in cols_to_hide if c in filtered_df.columns])
  
  # Add WhatsApp Web integration
  if "mobile_number" in filtered_df.columns:
      filtered_df["WhatsApp"] = filtered_df["mobile_number"].apply(
          lambda x: f"https://web.whatsapp.com/send?phone=91{x}" if str(x).strip() else ""
      )
          
  st.write("💡 *Tip: You can double-click any cell below to edit it directly. Click 'Save Inline Edits' when you are done!*")
  edited_df = st.data_editor(
      filtered_df, 
      use_container_width=True, 
      hide_index=True,
      disabled=["s_no", "WhatsApp"], # Don't allow editing the primary key or derived link
      column_config={
          "email": st.column_config.LinkColumn(
              "email",
              display_text=r"to=(.*)"
          ),
          "WhatsApp": st.column_config.LinkColumn(
              "WhatsApp",
              display_text="Message 💬"
          )
      }
  )

  col_export, col_save = st.columns([1, 1])
  with col_export:
      # Export back to CSV
      if st.button("Export Filtered Data to CSV", icon=":material/download:", type="primary"):
          csv = filtered_df.to_csv(index=False).encode('utf-8')
          st.download_button(
              label="Download CSV",
              data=csv,
              file_name='filtered_students.csv',
              mime='text/csv',
          )
          
  with col_save:
      # Save Inline Edits
      if st.button("Save Inline Edits", icon=":material/save:", type="primary"):
          changes = 0
          for idx in edited_df.index:
              original_row = filtered_df.loc[idx]
              edited_row = edited_df.loc[idx]
              if not original_row.equals(edited_row):
                  s_no = int(edited_row['s_no'])
                  updates = []
                  for col in edited_df.columns:
                      if col not in ['s_no', 'WhatsApp']:
                          val = str(edited_row[col]).replace("'", "''") if pd.notnull(edited_row[col]) else ''
                          updates.append(f"{col} = '{val}'")
                  query = f"UPDATE students SET {', '.join(updates)} WHERE s_no = {s_no}"
                  conn.execute(query)
                  changes += 1
          
          if changes > 0:
              conn.commit()
              st.success(f"Successfully saved {changes} row(s)!", icon=":material/check_circle:")
              time.sleep(1)
              st.rerun()
          else:
              st.info("No changes detected.")

elif nav_option == "Advanced Analytics":
  st.header("📊 Advanced Pivot Analytics", divider="blue")
  st.write("Cross-reference your data dynamically to gain insights.")
  
  col1, col2, col3 = st.columns(3)
  with col1:
      index_col = st.selectbox("Rows (Group By)", ["college_name", "branch", "admission_year", "course_type"])
  with col2:
      columns_col = st.selectbox("Columns (Breakdown)", ["branch", "admission_year", "course_type", "college_name"], index=1)
  with col3:
      val_col = st.selectbox("Values (Count)", ["s_no"])
      
  df_all = pd.read_sql("SELECT * FROM students", conn)
  if not df_all.empty:
      try:
          pivot = pd.pivot_table(
              df_all, 
              values=val_col, 
              index=index_col, 
              columns=columns_col, 
              aggfunc='count', 
              fill_value=0
          )
          st.dataframe(pivot, use_container_width=True)
          
          st.write(f"**Total Breakdown by {index_col}**")
          st.bar_chart(df_all[index_col].value_counts())
      except Exception as e:
          st.error(f"Could not generate pivot table: {e}")

elif nav_option == "Data Quality Audit":
  st.header("🚨 Data Quality Audit", divider="red")
  st.write("This dashboard automatically flags students missing critical information so you can easily follow up with them.")
  
  df_all = pd.read_sql("SELECT * FROM students", conn)
  
  c1, c2, c3 = st.columns(3)
  
  missing_email = df_all[df_all["email"].isna() | (df_all["email"] == "") | (df_all["email"] == "nan")]
  missing_mobile = df_all[df_all["mobile_number"].isna() | (df_all["mobile_number"] == "") | (df_all["mobile_number"] == "nan")]
  missing_aadhaar = df_all[df_all["aadhaar"].isna() | (df_all["aadhaar"] == "") | (df_all["aadhaar"] == "nan")]
  
  with c1:
      st.error(f"**Missing Email: {len(missing_email)}**", icon=":material/mail:")
      if not missing_email.empty:
          st.dataframe(missing_email[["s_no", "student_name", "branch"]], hide_index=True, use_container_width=True)
          
  with c2:
      st.warning(f"**Missing Mobile: {len(missing_mobile)}**", icon=":material/phone:")
      if not missing_mobile.empty:
          st.dataframe(missing_mobile[["s_no", "student_name", "branch"]], hide_index=True, use_container_width=True)
          
  with c3:
      st.info(f"**Missing Aadhaar: {len(missing_aadhaar)}**", icon=":material/badge:")
      if not missing_aadhaar.empty:
          st.dataframe(missing_aadhaar[["s_no", "student_name", "branch"]], hide_index=True, use_container_width=True)

elif nav_option == "🕵️‍♂️ Data Detective":
  st.header("🕵️‍♂️ The Data Detective", divider="red")
  st.write("Scanning database for deep anomalies (invalid formats, exact duplicates)...")
  
  df_all = pd.read_sql("SELECT * FROM students", conn)
  for col in ["tnea_code", "mobile_number", "aadhaar", "umis_number"]:
      if col in df_all.columns:
          df_all[col] = df_all[col].astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '')
          
  if not df_all.empty:
      c1, c2 = st.columns(2)
      
      invalid_phones = df_all[df_all['mobile_number'].apply(lambda x: len(re.sub(r'\D', '', str(x))) != 10 if x else False)]
      with c1:
          st.error(f"**Invalid Mobile Numbers (Not 10 digits): {len(invalid_phones)}**")
          if not invalid_phones.empty:
              st.dataframe(invalid_phones[["s_no", "student_name", "mobile_number"]], hide_index=True, use_container_width=True)
              
      invalid_aadhaar = df_all[df_all['aadhaar'].apply(lambda x: len(re.sub(r'\D', '', str(x))) != 12 if x else False)]
      with c2:
          st.warning(f"**Invalid Aadhaar Numbers (Not 12 digits): {len(invalid_aadhaar)}**")
          if not invalid_aadhaar.empty:
              st.dataframe(invalid_aadhaar[["s_no", "student_name", "aadhaar"]], hide_index=True, use_container_width=True)

      c3, c4 = st.columns(2)
      invalid_emails = df_all[df_all['email'].apply(lambda x: '@' not in str(x) if x else False)]
      with c3:
          st.error(f"**Invalid Emails (Missing @): {len(invalid_emails)}**")
          if not invalid_emails.empty:
              st.dataframe(invalid_emails[["s_no", "student_name", "email"]], hide_index=True, use_container_width=True)
              
      df_phones = df_all[df_all['mobile_number'] != '']
      duplicate_phones = df_phones[df_phones.duplicated(subset=['mobile_number'], keep=False)]
      with c4:
          st.warning(f"**Duplicate Mobile Numbers: {len(duplicate_phones)}**")
          if not duplicate_phones.empty:
              st.dataframe(duplicate_phones.sort_values('mobile_number')[["s_no", "student_name", "mobile_number"]], hide_index=True, use_container_width=True)

elif nav_option == "🤝 Team Builder":
  st.header("🤝 Smart Team Builder", divider="blue")
  st.write("Randomly assign students into perfectly balanced project groups.")
  
  df_all = pd.read_sql("SELECT * FROM students", conn)
  if df_all.empty:
      st.info("No students in database.")
  else:
      c1, c2 = st.columns(2)
      with c1:
          branch_filter = st.selectbox("Select Branch", ["All"] + list(df_all["branch"].dropna().unique()))
      with c2:
          team_size = st.number_input("Students per Team", min_value=2, max_value=20, value=st.session_state.get('team_size', 4))
          
      if branch_filter != "All":
          df_pool = df_all[df_all["branch"] == branch_filter]
      else:
          df_pool = df_all.copy()
          
      if st.button("🎲 Generate Teams", type="primary"):
          df_shuffled = df_pool.sample(frac=1).reset_index(drop=True)
          teams = []
          for i in range(0, len(df_shuffled), team_size):
              chunk = df_shuffled.iloc[i:i+team_size]
              team_members = chunk["student_name"].tolist()
              teams.append(team_members)
              
          team_df = pd.DataFrame({
              "Team Number": [f"Team {i+1}" for i in range(len(teams))],
              "Members": [", ".join(t) for t in teams],
              "Count": [len(t) for t in teams]
          })
          
          st.success(f"Successfully generated {len(teams)} teams!")
          st.dataframe(team_df, use_container_width=True, hide_index=True)

elif nav_option == "Update Student Portal":
  st.subheader("✏️ Student Information Entry & Update")
  st.write(
      "Select your serial number or name to fill in or update your UMIS"
      " details."
  )

  df_list = pd.read_sql("SELECT s_no, student_name FROM students", conn)

  # Create a display label combining S.No and Name (handling empty names)
  df_list["display_label"] = (
      df_list["s_no"].astype(str)
      + " - "
      + df_list["student_name"].fillna("").astype(str).replace("", "[Empty Slot]")
  )

  selected_label = st.selectbox(
      "Select Student Row", df_list["display_label"].tolist()
  )
  selected_sno = int(selected_label.split(" - ")[0])

  # Fetch specific student row
  student_row = pd.read_sql(
      f"SELECT * FROM students WHERE s_no = {selected_sno}", conn
  ).iloc[0]
  
  # Render Premium Profile Card
  email_val = f"<a href='https://mail.google.com/mail/?view=cm&fs=1&to={student_row['email']}' target='_blank' style='color: #1e88e5; text-decoration: none;'>{student_row['email']}</a>" if student_row['email'] else 'N/A'
  
  st.markdown(f"""
  <div class="profile-card">
      <div class="profile-header">
          🎓 {student_row['student_name'] if student_row['student_name'] else 'Unknown Student'} (S.No: {selected_sno})
      </div>
      <div class="profile-data">
          <div class="profile-item">
              <span class="profile-label">Branch</span>
              <span class="profile-value">{student_row['branch'] if student_row['branch'] else 'N/A'}</span>
          </div>
          <div class="profile-item">
              <span class="profile-label">Course Type</span>
              <span class="profile-value">{student_row['course_type'] if student_row['course_type'] else 'N/A'}</span>
          </div>
          <div class="profile-item">
              <span class="profile-label">Mobile Number</span>
              <span class="profile-value">{str(student_row['mobile_number']).replace(".0", "") if student_row['mobile_number'] else 'N/A'}</span>
          </div>
          <div class="profile-item">
              <span class="profile-label">Email ID</span>
              <span class="profile-value">{email_val}</span>
          </div>
      </div>
  </div>
  """, unsafe_allow_html=True)

  # ID Card CSS for printing
  st.markdown("""
  <style>
  @media print {
      body * { visibility: hidden; }
      .id-card, .id-card * { visibility: visible !important; }
      .id-card { position: absolute; left: 0; top: 0; width: 400px; margin: 0 !important; }
  }
  .id-card {
      border: 2px solid #333;
      border-radius: 15px;
      padding: 20px;
      background: linear-gradient(135deg, #ffffff 0%, #f0f0f0 100%);
      color: #333;
      width: 400px;
      margin: 20px 0;
      box-shadow: 0 10px 20px rgba(0,0,0,0.2);
      font-family: 'Inter', sans-serif;
  }
  .id-card h3 { text-align: center; color: #1e88e5; margin-top: 0; margin-bottom: 5px; }
  .id-card h5 { text-align: center; color: #555; margin-top: 0; margin-bottom: 20px; }
  .id-card .info-row { display: flex; justify-content: space-between; margin-bottom: 10px; border-bottom: 1px solid #ccc; padding-bottom: 5px; }
  .id-card .info-label { font-weight: bold; color: #666; font-size: 0.9em; }
  .id-card .info-val { font-weight: 600; color: #111; font-size: 0.95em; text-align: right; }
  </style>
  """, unsafe_allow_html=True)
  
  if st.button("🖨️ Generate & Print ID Card", type="secondary"):
      st.markdown(f"""
      <div class="id-card">
          <h3>UMIS STUDENT ID</h3>
          <h5>{student_row['college_name'] if student_row['college_name'] else 'College N/A'}</h5>
          <div class="info-row"><span class="info-label">Name</span><span class="info-val">{student_row['student_name']}</span></div>
          <div class="info-row"><span class="info-label">Branch</span><span class="info-val">{student_row['branch']}</span></div>
          <div class="info-row"><span class="info-label">Mobile</span><span class="info-val">{str(student_row['mobile_number']).replace('.0', '')}</span></div>
          <div class="info-row"><span class="info-label">UMIS No.</span><span class="info-val">{student_row['umis_number']}</span></div>
          <div class="info-row"><span class="info-label">Aadhaar</span><span class="info-val">{str(student_row['aadhaar']).replace('.0', '')}</span></div>
          <p style="text-align: center; margin-top: 20px; font-size: 0.8em; color: #888;">Press Ctrl+P (or Cmd+P) to Print</p>
      </div>
      """, unsafe_allow_html=True)
      st.info("Press **Ctrl+P** (Windows) or **Cmd+P** (Mac) to print the ID Card!")

  with st.form("update_form"):
    st.write("📝 **Edit Details Below**")

    col1, col2, col3 = st.columns(3)

    # Exact fields from your UMIS Excel format
    with col1:
      student_name = st.text_input(
          "Student Name (as in certificate)", value=student_row["student_name"]
      )
      admission_year = st.text_input(
          "Year of Admission (yyyy-yyyy)",
          value=(
              student_row["admission_year"]
              if student_row["admission_year"]
              else "2026-2027"
          ),
      )
      mobile_number = st.text_input(
          "Mobile Number",
          value=(
              str(student_row["mobile_number"]).replace(".0", "")
              if student_row["mobile_number"]
              else ""
          ),
      )
      aadhaar = st.text_input(
          "Aadhaar Number",
          value=(
              str(student_row["aadhaar"]).replace(".0", "")
              if student_row["aadhaar"]
              else ""
          ),
      )

    with col2:
      branch = st.text_input(
          "Branch Name",
          value=student_row["branch"] if student_row["branch"] else "MCA",
      )
      course_type = st.text_input(
          "Course Type (UG/PG)",
          value=student_row["course_type"] if student_row["course_type"] else "PG",
      )
      email = st.text_input("Email ID", value=student_row["email"])
      umis_number = st.text_input(
          "UMIS Number",
          value=(
              str(student_row["umis_number"]).replace(".0", "")
              if student_row["umis_number"]
              else ""
          ),
      )

    with col3:
      tnea_code = st.text_input(
          "TNEA College Code",
          value=(
              str(student_row["tnea_code"]).replace(".0", "")
              if student_row["tnea_code"]
              else "2633"
          ),
      )
      college_name = st.text_input(
          "College Name",
          value=(
              student_row["college_name"]
              if student_row["college_name"]
              else "Vidyaa Vikas College of Engineering and Technology"
          ),
      )
      year_of_study = st.text_input(
          "Year of Study",
          value=(
              str(student_row["year_of_study"]).replace(".0", "")
              if student_row["year_of_study"]
              else "1"
          ),
      )

    st.write("") # Small vertical space
    with st.container(horizontal=True, horizontal_alignment="right"):
      submitted = st.form_submit_button("Save Details", icon=":material/save:", type="primary")

    if submitted:
      errors = validate_student_data(student_name, mobile_number, email, aadhaar)
      if errors:
          for err in errors:
              st.error(err, icon=":material/error:")
      else:
        cursor = conn.cursor()
        cursor.execute(
            """
                  UPDATE students 
                  SET student_name = ?, admission_year = ?, tnea_code = ?, college_name = ?, 
                      branch = ?, mobile_number = ?, email = ?, aadhaar = ?, umis_number = ?, 
                      course_type = ?, year_of_study = ?
                  WHERE s_no = ?
              """,
            (
                student_name,
                admission_year,
                tnea_code,
                college_name,
                branch,
                mobile_number,
                email,
                aadhaar,
                umis_number,
                course_type,
                year_of_study,
                selected_sno,
            ),
        )
        conn.commit()
        st.success(
            f"Successfully updated details for S.No {selected_sno} ({student_name})!",
            icon=":material/check_circle:"
        )

elif nav_option == "Add New Student":
  st.subheader("➕ Add New Student")
  st.write("Fill in the details below to add a new student to the database.")

  with st.form("add_new_form"):
    col1, col2, col3 = st.columns(3)

    with col1:
      student_name = st.text_input("Student Name (as in certificate)", value="")
      admission_year = st.text_input("Year of Admission (yyyy-yyyy)", value="2026-2027")
      mobile_number = st.text_input("Mobile Number", value="")
      aadhaar = st.text_input("Aadhaar Number", value="")

    with col2:
      branch = st.text_input("Branch Name", value="MCA")
      course_type = st.text_input("Course Type (UG/PG)", value="PG")
      email = st.text_input("Email ID", value="")
      umis_number = st.text_input("UMIS Number", value="")

    with col3:
      tnea_code = st.text_input("TNEA College Code", value="2633")
      college_name = st.text_input("College Name", value="Vidyaa Vikas College of Engineering and Technology")
      year_of_study = st.text_input("Year of Study", value="1")

    st.write("") # Small vertical space
    with st.container(horizontal=True, horizontal_alignment="right"):
      submitted = st.form_submit_button("Add Student", icon=":material/person_add:", type="primary")

    if submitted:
      errors = validate_student_data(student_name, mobile_number, email, aadhaar)
      if errors:
          for err in errors:
              st.error(err, icon=":material/error:")
      else:
        cursor = conn.cursor()
        
        # Check if there is an empty slot available
        cursor.execute("SELECT MIN(s_no) FROM students WHERE student_name = '' OR student_name IS NULL")
        empty_slot = cursor.fetchone()[0]
        
        if empty_slot is not None:
            # We found an empty slot, let's update it instead of inserting a new row
            next_sno = int(empty_slot)
            cursor.execute(
                """
                    UPDATE students 
                    SET student_name = ?, admission_year = ?, tnea_code = ?, college_name = ?, 
                        branch = ?, mobile_number = ?, email = ?, aadhaar = ?, umis_number = ?, 
                        course_type = ?, year_of_study = ?
                    WHERE s_no = ?
                """,
                (
                    student_name, admission_year, tnea_code, college_name,
                    branch, mobile_number, email, aadhaar, umis_number,
                    course_type, year_of_study, next_sno
                ),
            )
        else:
            # No empty slots, insert a new row
            cursor.execute("SELECT MAX(s_no) FROM students")
            result = cursor.fetchone()
            next_sno = 1 if result[0] is None else int(result[0]) + 1
            
            cursor.execute(
                """
                      INSERT INTO students (
                          s_no, student_name, admission_year, tnea_code, college_name, 
                          branch, mobile_number, email, aadhaar, umis_number, 
                          course_type, year_of_study
                      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                  """,
                (
                    next_sno, student_name, admission_year, tnea_code, college_name,
                    branch, mobile_number, email, aadhaar, umis_number,
                    course_type, year_of_study,
                ),
            )
        conn.commit()
        st.success(
            f"Successfully added {student_name} with S.No {next_sno}!",
            icon=":material/check_circle:"
        )

elif nav_option == "Bulk Upload":
  st.subheader("📤 Bulk Upload Students")
  st.write("Upload an Excel or CSV file matching the UMIS format to add multiple students at once.")
  
  uploaded_file = st.file_uploader("Choose a file", type=["xlsx", "xls", "csv"])
  
  if uploaded_file is not None:
      try:
          if uploaded_file.name.endswith(".csv"):
              df_upload = pd.read_csv(uploaded_file)
          else:
              df_upload = pd.read_excel(uploaded_file)
              
          st.write("Preview of uploaded data (First 5 rows):")
          st.dataframe(df_upload.head())
          
          if st.button("Import Data", type="primary", icon=":material/publish:"):
              expected_cols = [
                  "s_no", "admission_year", "tnea_code", "college_name",
                  "student_name", "branch", "mobile_number", "email",
                  "aadhaar", "umis_number", "citizenship", "passport",
                  "transfer_readmission", "course_type", "year_of_study"
              ]
              
              if len(df_upload.columns) == len(expected_cols):
                  df_upload.columns = expected_cols
                  df_upload = df_upload.fillna("")
                  df_upload = df_upload[df_upload["student_name"] != ""]
                  
                  if not df_upload.empty:
                      cursor = conn.cursor()
                      cursor.execute("SELECT MAX(s_no) FROM students")
                      res = cursor.fetchone()
                      start_sno = 1 if res[0] is None else int(res[0]) + 1
                      
                      # Re-assign s_no to avoid primary key conflicts
                      df_upload["s_no"] = range(start_sno, start_sno + len(df_upload))
                      df_upload.to_sql("students", conn, index=False, if_exists="append")
                      
                      st.success(f"Successfully imported {len(df_upload)} students! They are now in the Master Database.", icon=":material/check_circle:")
                  else:
                      st.warning("No valid students found in the file. Ensure the 'student_name' column has data.", icon=":material/warning:")
              else:
                  st.error(f"Column mismatch! The file must have exactly {len(expected_cols)} columns (it currently has {len(df_upload.columns)}).", icon=":material/error:")
      except Exception as e:
          st.error(f"Error reading file: {e}")

elif nav_option == "Delete Student":
  st.subheader("🗑️ Delete Student")
  st.write("Select a student to remove them from the database permanently.")

  # Show success message if we just deleted someone
  if "delete_success" in st.session_state:
      st.success(st.session_state.delete_success, icon=":material/check_circle:")
      del st.session_state.delete_success

  df_list = pd.read_sql("SELECT s_no, student_name, branch FROM students", conn)

  if df_list.empty:
      st.info("No students found in the database.")
  else:
      # Create a display label combining S.No and Name
      df_list["display_label"] = (
          df_list["s_no"].astype(str)
          + " - "
          + df_list["student_name"].fillna("").astype(str).replace("", "[Empty Slot]")
      )

      selected_label = st.selectbox(
          "Select Student to Delete", df_list["display_label"].tolist()
      )
      selected_sno = int(selected_label.split(" - ")[0])
      
      # Fetch specific student row
      student_row = pd.read_sql(
          f"SELECT * FROM students WHERE s_no = {selected_sno}", conn
      ).iloc[0]
      
      # Render Premium Profile Card
      email_val = f"<a href='https://mail.google.com/mail/?view=cm&fs=1&to={student_row['email']}' target='_blank' style='color: #1e88e5; text-decoration: none;'>{student_row['email']}</a>" if student_row['email'] else 'N/A'
      
      st.markdown(f"""
      <div class="profile-card" style="border-color: rgba(255, 75, 75, 0.3);">
          <div class="profile-header" style="color: #ff4b4b;">
              ⚠️ Delete Candidate: {student_row['student_name'] if student_row['student_name'] else 'Unknown Student'}
          </div>
          <div class="profile-data">
              <div class="profile-item">
                  <span class="profile-label">Branch</span>
                  <span class="profile-value">{student_row['branch'] if student_row['branch'] else 'N/A'}</span>
              </div>
              <div class="profile-item">
                  <span class="profile-label">Email ID</span>
                  <span class="profile-value">{email_val}</span>
              </div>
              <div class="profile-item">
                  <span class="profile-label">S.No</span>
                  <span class="profile-value">{selected_sno}</span>
              </div>
          </div>
      </div>
      """, unsafe_allow_html=True)

      st.warning(f"Are you sure you want to delete **{selected_label}**? This action cannot be undone.", icon=":material/warning:")

      if st.button("Delete Student", icon=":material/delete:", type="primary"):
          cursor = conn.cursor()
          cursor.execute("DELETE FROM students WHERE s_no = ?", (selected_sno,))
          conn.commit()
          st.session_state.delete_success = f"Successfully deleted {selected_label}."
          st.rerun()

elif nav_option == "👁️ God Mode (SQL)":
  st.header("👁️ God Mode: Raw SQL Sandbox", divider="red")
  st.warning("⚠️ **DANGER ZONE:** You are executing raw SQL queries directly against the production database. You can drop tables, delete data, or corrupt the schema.")
  
  query = st.text_area("SQL Query", "SELECT * FROM students LIMIT 10;", height=150)
  
  if st.button("Execute Query", type="primary"):
      try:
          if query.strip().upper().startswith("SELECT"):
              result_df = pd.read_sql(query, conn)
              
              # Clean up .0 decimals for known numeric columns
              for col in ["tnea_code", "mobile_number", "aadhaar", "umis_number"]:
                  if col in result_df.columns:
                      result_df[col] = result_df[col].astype(str).str.replace(r'\.0$', '', regex=True).replace('nan', '')
                      
              st.dataframe(result_df, use_container_width=True)
              st.success(f"Query returned {len(result_df)} row(s).")
          else:
              cursor = conn.cursor()
              cursor.execute(query)
              conn.commit()
              st.success(f"Query executed successfully. Rows affected: {cursor.rowcount}")
      except Exception as e:
          st.error(f"SQL Error: {e}")

conn.close()