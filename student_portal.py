import streamlit as st
import sqlite3
import pandas as pd
import time
import re

# --- CONFIGURATION ---
DB_NAME = "mca_students_umis.db"

# --- PAGE CONFIG & CSS ---
st.set_page_config(
    page_title="Student Data Portal",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS for Premium Look and Hiding Streamlit Elements
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    .stApp {
        background-color: #0e1117;
        color: #e0e6ed;
    }
    
    .login-box {
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 12px;
        padding: 30px;
        margin-top: 50px;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.3);
    }
    
    /* Hide the top header bar and sidebar completely for students */
    header {visibility: hidden;}
    [data-testid="stSidebarNav"] {display: none;}
    [data-testid="collapsedControl"] {display: none;}
</style>
""", unsafe_allow_html=True)

# Initialize Session State
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'student_sno' not in st.session_state:
    st.session_state.student_sno = None

conn = sqlite3.connect(DB_NAME)

try:
    _ = pd.read_sql("SELECT 1 FROM students LIMIT 1", conn)
    db_ready = True
except pd.errors.DatabaseError:
    db_ready = False

if not db_ready:
    st.error("Database is not initialized or is missing. Please contact the administrator.")
    st.stop()

if not st.session_state.logged_in:
    st.markdown("<h1 style='text-align: center; color: #1e88e5;'>🎓 Student Data Portal</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Welcome. Please authenticate to view and update your records.</p>", unsafe_allow_html=True)
    
    st.markdown('<div class="login-box">', unsafe_allow_html=True)
    with st.form("login_form"):
        st.subheader("Secure Login")
        umis_input = st.text_input("UMIS Number", placeholder="Enter your UMIS Number")
        aadhaar_input = st.text_input("Aadhaar Number", type="password", placeholder="Enter your Aadhaar Number (Acts as Password)")
        
        submit_login = st.form_submit_button("Login", type="primary", use_container_width=True)
        
        if submit_login:
            if not umis_input or not aadhaar_input:
                st.error("Please enter both UMIS and Aadhaar.")
            else:
                # Clean inputs to prevent matching issues
                u_in = str(umis_input).strip().replace(".0", "")
                a_in = str(aadhaar_input).strip().replace(".0", "").replace(" ", "")
                
                # Pull credentials safely
                df = pd.read_sql("SELECT s_no, umis_number, aadhaar FROM students", conn)
                
                # Cleanup db values for comparison
                df['umis_clean'] = df['umis_number'].astype(str).str.replace(r'\\.0$', '', regex=True).str.strip()
                df['aadhaar_clean'] = df['aadhaar'].astype(str).str.replace(r'\\.0$', '', regex=True).str.strip().str.replace(" ", "")
                
                match = df[(df['umis_clean'] == u_in) & (df['aadhaar_clean'] == a_in)]
                
                if not match.empty:
                    st.session_state.logged_in = True
                    st.session_state.student_sno = int(match.iloc[0]['s_no'])
                    st.rerun()
                else:
                    st.error("Invalid UMIS Number or Aadhaar Number. Please try again.")
    st.markdown('</div>', unsafe_allow_html=True)

else:
    # --- LOGGED IN VIEW ---
    st.markdown("<h2 style='color: #1e88e5;'>📝 Your Student Profile</h2>", unsafe_allow_html=True)
    
    if st.button("🚪 Logout", type="secondary"):
        st.session_state.logged_in = False
        st.session_state.student_sno = None
        st.rerun()
        
    # Fetch user data
    student_row_df = pd.read_sql(f"SELECT * FROM students WHERE s_no = {st.session_state.student_sno}", conn)
    
    if student_row_df.empty:
        st.error("Your account record could not be found. It may have been deleted.")
        st.session_state.logged_in = False
        st.session_state.student_sno = None
        time.sleep(2)
        st.rerun()
        
    student_row = student_row_df.iloc[0]
    
    st.info("Please review and update your information below. Click 'Submit Update' when finished.")
    
    with st.form("student_update_form"):
        st.subheader("Personal Details")
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("Full Name", value=str(student_row['student_name']) if pd.notnull(student_row['student_name']) else "")
            admission_year = st.text_input("Admission Year (e.g. 2026-2027)", value=str(student_row['admission_year']) if pd.notnull(student_row['admission_year']) else "")
        with col2:
            mobile = st.text_input("Mobile Number", value=str(student_row['mobile_number']).replace(".0", "").strip() if pd.notnull(student_row['mobile_number']) else "")
            email = st.text_input("Email ID", value=str(student_row['email']).strip() if pd.notnull(student_row['email']) else "")
            
        st.subheader("Academic Details")
        c1, c2 = st.columns(2)
        with c1:
            college_name = st.text_input("College Name", value=str(student_row['college_name']).strip() if pd.notnull(student_row['college_name']) else "")
            branch = st.text_input("Branch", value=str(student_row['branch']).strip() if pd.notnull(student_row['branch']) else "")
            
            # Safe index fallback for year_of_study
            yos_options = ["1", "2", "3", "4"]
            try:
                yos_idx = yos_options.index(str(student_row['year_of_study']).strip())
            except ValueError:
                yos_idx = 0
            year_of_study = st.selectbox("Year of Study", yos_options, index=yos_idx)
            
        with c2:
            tnea = st.text_input("TNEA Code", value=str(student_row['tnea_code']).replace(".0", "").strip() if pd.notnull(student_row['tnea_code']) else "")
            
            # Safe index fallback for course_type
            ct_options = ["UG", "PG"]
            try:
                ct_idx = ct_options.index(str(student_row['course_type']).strip().upper())
            except ValueError:
                ct_idx = 1 # Default to PG
            course_type = st.selectbox("Course Type", ct_options, index=ct_idx)
            
        submit = st.form_submit_button("Submit Update", type="primary", use_container_width=True)
        
        if submit:
            if not name or not mobile:
                st.error("Name and Mobile Number are required fields.")
            elif mobile and not re.match(r"^\d{10}$", mobile):
                st.error("Mobile Number must be exactly 10 digits.")
            elif email and "@" not in email:
                st.error("Please provide a valid Email ID.")
            else:
                try:
                    cursor = conn.cursor()
                    cursor.execute(
                        """
                        UPDATE students 
                        SET student_name = ?, admission_year = ?, tnea_code = ?, college_name = ?, 
                            branch = ?, mobile_number = ?, email = ?, course_type = ?, year_of_study = ?
                        WHERE s_no = ?
                        """,
                        (
                            name, admission_year, tnea, college_name,
                            branch, mobile, email, course_type, year_of_study, 
                            st.session_state.student_sno
                        ),
                    )
                    conn.commit()
                    st.success("✅ Profile successfully updated! Your changes have been recorded.")
                    time.sleep(1.5)
                    st.rerun()
                except sqlite3.Error as e:
                    st.error(f"Database Error: {e}. Please try again.")

conn.close()
