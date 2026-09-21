import pandas as pd
import sqlite3
import streamlit as st
import time
import re
import altair as alt
from fpdf import FPDF
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from werkzeug.security import generate_password_hash, check_password_hash
import pdfplumber
import io

# --- CONFIGURATION ---
DB_NAME = "mca_students_2026.db"
EXCEL_FILE = "UMIS_API STUDENT LIST_2026-27.xlsx"

# --- PAGE CONFIG & CSS ---
st.set_page_config(
    page_title="MCA Student Portal",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Premium Look
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    /* Global Background and Text */
    .stApp {
        background-color: #0e1117;
        color: #e0e6ed;
    }
    
    /* Glassmorphism Cards */
    .glass-card {
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 12px;
        padding: 25px;
        box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.3);
        backdrop-filter: blur(10px);
        margin-bottom: 20px;
    }
    
    /* Metrics Styling */
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
    
    /* Headings with Gradients */
    h1, h2, h3 {
        background: -webkit-linear-gradient(45deg, #4b9fff, #ff4b4b);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 700;
    }
    
    /* Animations */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    @keyframes pulseGlow {
        0% { box-shadow: 0 0 5px rgba(75, 159, 255, 0.4); }
        50% { box-shadow: 0 0 15px rgba(75, 159, 255, 0.8); }
        100% { box-shadow: 0 0 5px rgba(75, 159, 255, 0.4); }
    }

    /* Apply fade in to main containers */
    .block-container {
        animation: fadeIn 0.8s ease-out;
    }

    /* Input Field Enhancements */
    .stTextInput input, .stSelectbox > div[data-baseweb="select"] {
        transition: all 0.3s ease;
        border: 1px solid rgba(255, 255, 255, 0.1);
        background: rgba(0, 0, 0, 0.2) !important;
    }
    .stTextInput input:focus, .stSelectbox > div[data-baseweb="select"]:focus-within {
        border-color: #4b9fff !important;
        box-shadow: 0 0 10px rgba(75, 159, 255, 0.3) !important;
    }
    
    /* Primary Button Hover & Pulse Animations */
    button[kind="primary"] {
        transition: all 0.3s ease-in-out !important;
        background: linear-gradient(45deg, #4b9fff, #ff4b4b) !important;
        border: none !important;
    }
    button[kind="primary"]:hover {
        transform: translateY(-3px) scale(1.02) !important;
        animation: pulseGlow 1.5s infinite;
    }
    
    /* Secondary Button enhancements */
    button[kind="secondary"] {
        transition: all 0.2s ease-in-out !important;
    }
    button[kind="secondary"]:hover {
        transform: translateY(-2px) !important;
        border-color: #ff4b4b !important;
        color: #ff4b4b !important;
    }
</style>
""", unsafe_allow_html=True)

# --- DATABASE SETUP ---
def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS students (
            s_no INTEGER PRIMARY KEY,
            admission_year TEXT,
            tnea_code TEXT,
            college_name TEXT,
            student_name TEXT,
            branch TEXT,
            mobile_number TEXT,
            email TEXT,
            aadhaar TEXT,
            umis_number TEXT,
            course_type TEXT,
            year_of_study TEXT,
            status TEXT DEFAULT 'Pending'
        )
    """)
    
    # Safely migrate existing table to include fee columns
    cursor.execute("PRAGMA table_info(students)")
    columns = [col[1] for col in cursor.fetchall()]
    
    if 'college_fee_total' not in columns:
        cursor.execute("ALTER TABLE students ADD COLUMN college_fee_total REAL DEFAULT 0.0")
    if 'college_fee_paid' not in columns:
        cursor.execute("ALTER TABLE students ADD COLUMN college_fee_paid REAL DEFAULT 0.0")
    if 'exam_fee_total' not in columns:
        cursor.execute("ALTER TABLE students ADD COLUMN exam_fee_total REAL DEFAULT 0.0")
    if 'exam_fee_paid' not in columns:
        cursor.execute("ALTER TABLE students ADD COLUMN exam_fee_paid REAL DEFAULT 0.0")
        
    # Full Database columns
    if 'dob' not in columns:
        cursor.execute("ALTER TABLE students ADD COLUMN dob TEXT DEFAULT ''")
    if 'gender' not in columns:
        cursor.execute("ALTER TABLE students ADD COLUMN gender TEXT DEFAULT ''")
    if 'blood_group' not in columns:
        cursor.execute("ALTER TABLE students ADD COLUMN blood_group TEXT DEFAULT ''")
    if 'guardian_name' not in columns:
        cursor.execute("ALTER TABLE students ADD COLUMN guardian_name TEXT DEFAULT ''")
    if 'guardian_mobile' not in columns:
        cursor.execute("ALTER TABLE students ADD COLUMN guardian_mobile TEXT DEFAULT ''")
    if 'address' not in columns:
        cursor.execute("ALTER TABLE students ADD COLUMN address TEXT DEFAULT ''")
    if 'cgpa' not in columns:
        cursor.execute("ALTER TABLE students ADD COLUMN cgpa REAL DEFAULT 0.0")
    if 'attendance_percent' not in columns:
        cursor.execute("ALTER TABLE students ADD COLUMN attendance_percent REAL DEFAULT 0.0")
    if 'docs_verified' not in columns:
        cursor.execute("ALTER TABLE students ADD COLUMN docs_verified TEXT DEFAULT 'No'")
        
    conn.commit()
    conn.close()

def load_excel_to_db():
    conn = sqlite3.connect(DB_NAME)
    df_check = pd.read_sql("SELECT COUNT(*) FROM students", conn)
    if df_check.iloc[0, 0] == 0:
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name="Format")
            df.columns = [
                "s_no", "admission_year", "tnea_code", "college_name", "student_name", 
                "branch", "mobile_number", "email", "aadhaar", "umis_number", 
                "citizenship", "passport", "transfer_readmission", "course_type", "year_of_study"
            ]
            
            # Drop unused columns
            df = df.drop(columns=["citizenship", "passport", "transfer_readmission"], errors='ignore')
            
            # Drop empty rows and fill NaNs
            df = df.dropna(subset=["student_name"])
            df = df.fillna("")
            
            # Clean up default '.0' from excel floats
            for col in ['s_no', 'umis_number', 'aadhaar', 'mobile_number', 'tnea_code']:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
            df['status'] = 'Pending'
            df = df.astype(str)
            df.to_sql("students", conn, if_exists="append", index=False)
        except Exception as e:
            st.error(f"Error loading Excel data: {e}")
    conn.close()

init_db()
load_excel_to_db()

# --- APP STATE & ROUTING ---
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'student_sno' not in st.session_state:
    st.session_state.student_sno = None

st.sidebar.title("Navigation")
role = st.sidebar.radio("Select Portal View", ["Student Portal", "Advisor Dashboard"])

conn = sqlite3.connect(DB_NAME)

# --- STUDENT PORTAL ---
if role == "Student Portal":
    
    if not st.session_state.logged_in:
        st.markdown("<h1 style='text-align: center;'>🎓 Student Authentication</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #888;'>Please authenticate to view and update your records.</p>", unsafe_allow_html=True)
        
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        with st.form("login_form"):
            st.subheader("Secure Login")
            umis_input = st.text_input("UMIS Number", placeholder="e.g. 123456")
            aadhaar_input = st.text_input("Aadhaar Number", type="password", placeholder="Used as your secure password")
            
            submit_login = st.form_submit_button("Login to Portal", type="primary", use_container_width=True)
            
            if submit_login:
                if not umis_input or not aadhaar_input:
                    st.error("Please enter both UMIS and Aadhaar.")
                else:
                    u_in = str(umis_input).strip().replace(".0", "")
                    a_in = str(aadhaar_input).strip().replace(".0", "").replace(" ", "")
                    
                    df = pd.read_sql("SELECT s_no, umis_number, aadhaar FROM students", conn)
                    df['umis_clean'] = df['umis_number'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                    df['aadhaar_clean'] = df['aadhaar'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.replace(" ", "")
                    
                    match = df[(df['umis_clean'] == u_in) & (df['aadhaar_clean'] == a_in)]
                    
                    if not match.empty:
                        st.session_state.logged_in = True
                        st.session_state.student_sno = int(match.iloc[0]['s_no'])
                        st.rerun()
                    else:
                        st.error("Invalid credentials. Please verify your UMIS and Aadhaar.")
        st.markdown('</div>', unsafe_allow_html=True)

    else:
        st.markdown("<h1>📝 Update Your Profile</h1>", unsafe_allow_html=True)
        
        if st.sidebar.button("🚪 Logout", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.student_sno = None
            st.rerun()
            
        student_row_df = pd.read_sql("SELECT * FROM students WHERE s_no = ?", conn, params=(st.session_state.student_sno,))
        
        if student_row_df.empty:
            st.error("Account not found.")
            st.session_state.logged_in = False
            st.rerun()
            
        student_data = student_row_df.iloc[0]
        
        st.info(f"Welcome back, **{student_data['student_name']}**! Please review your details below.")
        
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        with st.form("student_update_form"):
            st.subheader("Edit Profile Information")
            
            sc1, sc2, sc3 = st.columns(3)
            with sc1:
                mobile = st.text_input("Mobile Number", value=str(student_data.get('mobile_number', '')).strip() if pd.notnull(student_data.get('mobile_number', '')) and str(student_data.get('mobile_number', '')).strip() != 'nan' else "")
                dob = st.text_input("Date of Birth", value=str(student_data.get('dob', '')))
                gname = st.text_input("Guardian Name", value=str(student_data.get('guardian_name', '')))
            with sc2:
                email = st.text_input("Email ID", value=str(student_data.get('email', '')).strip() if pd.notnull(student_data.get('email', '')) and str(student_data.get('email', '')).strip() != 'nan' else "")
                gen_idx = ["Male", "Female", "Other", ""].index(student_data.get('gender', '')) if student_data.get('gender', '') in ["Male", "Female", "Other", ""] else 3
                gender = st.selectbox("Gender", ["Male", "Female", "Other", ""], index=gen_idx)
                gmobile = st.text_input("Guardian Mobile", value=str(student_data.get('guardian_mobile', '')))
            with sc3:
                aadhaar = st.text_input("Aadhaar Number", value=str(student_data.get('aadhaar', '')).strip() if pd.notnull(student_data.get('aadhaar', '')) and str(student_data.get('aadhaar', '')).strip() != 'nan' else "")
                blood = st.text_input("Blood Group", value=str(student_data.get('blood_group', '')))
                umis = st.text_input("UMIS Number", value=str(student_data.get('umis_number', '')).strip() if pd.notnull(student_data.get('umis_number', '')) and str(student_data.get('umis_number', '')).strip() != 'nan' else "")
            
            address = st.text_area("Permanent Address", value=str(student_data.get('address', '')))
            
            submitted = st.form_submit_button("Submit Updates", type="primary", use_container_width=True)
            
            if submitted:
                errors = []
                if mobile and not re.match(r"^\d{10}$", mobile):
                    errors.append("Mobile Number must be exactly 10 digits.")
                if aadhaar and not re.match(r"^\d{12}$", aadhaar.replace(" ", "")):
                    errors.append("Aadhaar Number must be 12 digits.")
                if email and "@" not in email:
                    errors.append("Please provide a valid Email ID.")
                
                if errors:
                    for err in errors:
                        st.error(err)
                else:
                    cursor = conn.cursor()
                    cursor.execute(
                        """
                        UPDATE students 
                        SET mobile_number = ?, email = ?, aadhaar = ?, umis_number = ?, 
                            dob = ?, gender = ?, blood_group = ?, guardian_name = ?, guardian_mobile = ?, address = ?, status = 'Updated'
                        WHERE s_no = ?
                        """,
                        (mobile, email, aadhaar, umis, dob, gender, blood, gname, gmobile, address, st.session_state.student_sno),
                    )
                    conn.commit()
                    st.balloons()
                    st.toast("✅ Profile updated successfully!", icon="🎉")
                    time.sleep(1.5)
                    st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Financial Summary Widget
        st.markdown("### 💸 Financial Summary")
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        try:
            col_fee_total = float(student_data.get('college_fee_total', 0.0) or 0.0)
            col_fee_paid = float(student_data.get('college_fee_paid', 0.0) or 0.0)
            exm_fee_total = float(student_data.get('exam_fee_total', 0.0) or 0.0)
            exm_fee_paid = float(student_data.get('exam_fee_paid', 0.0) or 0.0)
            
            c1, c2 = st.columns(2)
            c1.metric("College Fees Due", f"₹ {col_fee_total - col_fee_paid:,.2f}", f"Paid: ₹ {col_fee_paid:,.2f}", delta_color="inverse")
            c2.metric("Exam Fees Due", f"₹ {exm_fee_total - exm_fee_paid:,.2f}", f"Paid: ₹ {exm_fee_paid:,.2f}", delta_color="inverse")
        except Exception:
            st.info("Fee details are currently unavailable.")
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown("### 📄 Downloads")
        d1, d2 = st.columns(2)
        
        # Receipt PDF
        if d1.button("Generate Fee Receipt", use_container_width=True):
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(200, 10, text="College Fee Receipt", ln=True, align='C')
            pdf.set_font("Arial", size=12)
            pdf.ln(10)
            pdf.cell(200, 10, text=f"Student Name: {student_data.get('student_name', '')}", ln=True)
            pdf.cell(200, 10, text=f"UMIS Number: {student_data.get('umis_number', '')}", ln=True)
            pdf.cell(200, 10, text=f"Branch: {student_data.get('branch', '')}", ln=True)
            pdf.cell(200, 10, text=f"Total College Fee Paid: INR {col_fee_paid:,.2f}", ln=True)
            pdf.cell(200, 10, text=f"Total Exam Fee Paid: INR {exm_fee_paid:,.2f}", ln=True)
            pdf.ln(10)
            pdf.cell(200, 10, text="Thank you for your payment.", ln=True)
            
            pdf_bytes = bytes(pdf.output())
            d1.download_button(label="Download Receipt PDF", data=pdf_bytes, file_name=f"{student_data.get('student_name', 'Student')}_Receipt.pdf", mime="application/pdf", type="primary", use_container_width=True)
            
        # ID Card PDF
        if d2.button("Generate ID Card", use_container_width=True):
            pdf = FPDF(orientation='P', unit='mm', format=(53.98, 85.6)) # Standard ID card size portrait
            pdf.add_page()
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(0, 8, text="Student ID Card", ln=True, align='C')
            pdf.set_font("Arial", size=8)
            pdf.cell(0, 6, text=f"Name: {student_data.get('student_name', '')}", ln=True)
            pdf.cell(0, 6, text=f"UMIS: {student_data.get('umis_number', '')}", ln=True)
            pdf.cell(0, 6, text=f"Branch: {student_data.get('branch', '')}", ln=True)
            pdf.cell(0, 6, text=f"DOB: {student_data.get('dob', '')}", ln=True)
            pdf.cell(0, 6, text=f"Blood: {student_data.get('blood_group', '')}", ln=True)
            
            pdf_bytes = bytes(pdf.output())
            d2.download_button(label="Download ID Card PDF", data=pdf_bytes, file_name=f"{student_data.get('student_name', 'Student')}_ID_Card.pdf", mime="application/pdf", type="primary", use_container_width=True)

# --- ADVISOR DASHBOARD ---
elif role == "Advisor Dashboard":
    if not st.session_state.get('advisor_logged_in', False):
        st.markdown("<h1 style='text-align: center;'>🔒 Advisor Authentication</h1>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center; color: #888;'>Login with your secure credentials.</p>", unsafe_allow_html=True)
        
        st.markdown('<div class="glass-card">', unsafe_allow_html=True)
        with st.form("advisor_login_form"):
            user = st.text_input("Username")
            pwd = st.text_input("Password", type="password")
            if st.form_submit_button("Secure Login", type="primary", use_container_width=True):
                user_record = pd.read_sql("SELECT * FROM users WHERE username=?", conn, params=(user,))
                if not user_record.empty:
                    hashed_pw = user_record.iloc[0]['password_hash']
                    if check_password_hash(hashed_pw, pwd):
                        st.session_state.advisor_logged_in = True
                        st.session_state.advisor_role = user_record.iloc[0]['role']
                        st.session_state.advisor_username = user
                        st.rerun()
                    else:
                        st.error("Invalid password. Access denied.")
                else:
                    st.error("User not found.")
        st.markdown('</div>', unsafe_allow_html=True)
        st.stop()
        
    if st.sidebar.button("🚪 Logout", use_container_width=True):
        st.session_state.advisor_logged_in = False
        st.session_state.advisor_role = None
        st.session_state.advisor_username = None
        st.rerun()

    st.markdown("<h1>🔒 Advisor Master Dashboard</h1>", unsafe_allow_html=True)
    st.write("Real-time oversight of student data, updates, and fees.")
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Overview", "💰 Fees Management", "📂 Bulk Operations", "📈 Analytics", "📧 Automations"])
    
    df_all = pd.read_sql("SELECT * FROM students", conn)
    
    with tab1:
        if not df_all.empty:
            total_students = len(df_all)
            updated = len(df_all[df_all['status'] == 'Updated'])
            pending = total_students - updated
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Students", total_students)
            
            # Calculate percentage safely
            percentage = 0 if total_students == 0 else int((updated / total_students) * 100)
            c2.metric("Profiles Updated", updated, f"{percentage}%")
            c3.metric("Pending Updates", pending)
            
            st.markdown("<br>", unsafe_allow_html=True)
            # Defaulters & Alerts Section
            st.markdown("### ⚠️ Action Required (Defaulters & Alerts)")
            
            # Ensure columns are numeric for calculation
            for col in ['college_fee_total', 'college_fee_paid', 'exam_fee_total', 'exam_fee_paid', 'attendance_percent']:
                if col not in df_all.columns:
                    df_all[col] = 0.0
                df_all[col] = pd.to_numeric(df_all[col], errors='coerce').fillna(0.0)
                
            df_all['Total Due'] = (df_all['college_fee_total'] - df_all['college_fee_paid']) + (df_all['exam_fee_total'] - df_all['exam_fee_paid'])
            
            fee_defaulters = df_all[df_all['Total Due'] > 0]
            attendance_defaulters = df_all[df_all['attendance_percent'] < 75.0]
            
            d1, d2 = st.columns(2)
            with d1:
                st.error(f"💰 Fee Defaulters ({len(fee_defaulters)})")
                if not fee_defaulters.empty:
                    st.dataframe(fee_defaulters[['student_name', 'branch', 'Total Due']], hide_index=True, use_container_width=True)
                else:
                    st.success("No fee defaulters!")
            
            with d2:
                st.warning(f"📉 Low Attendance < 75% ({len(attendance_defaulters)})")
                if not attendance_defaulters.empty:
                    st.dataframe(attendance_defaulters[['student_name', 'branch', 'attendance_percent']], hide_index=True, use_container_width=True)
                else:
                    st.success("All students have satisfactory attendance!")
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            # Add New Student
            with st.expander("➕ Add New Student"):
                with st.form("add_student_form"):
                    st.write("Fill out the details to manually add a new student to the database. They can then log in via the Student Portal.")
                    
                    st.markdown("#### Academic Info")
                    c1, c2, c3 = st.columns(3)
                    new_name = c1.text_input("Full Name (Required)")
                    new_umis = c2.text_input("UMIS Number (Required)")
                    new_branch = c3.text_input("Branch")
                    
                    c4, c5 = st.columns(2)
                    new_cgpa = c4.number_input("CGPA", min_value=0.0, max_value=10.0, step=0.1)
                    new_attendance = c5.number_input("Attendance %", min_value=0.0, max_value=100.0, step=1.0)
                    
                    st.markdown("#### Demographics")
                    c6, c7, c8 = st.columns(3)
                    new_dob = c6.text_input("Date of Birth (YYYY-MM-DD)")
                    new_gender = c7.selectbox("Gender", ["Male", "Female", "Other", ""])
                    new_blood = c8.text_input("Blood Group")
                    
                    st.markdown("#### Contact Info")
                    c9, c10, c11 = st.columns(3)
                    new_mobile = c9.text_input("Mobile Number")
                    new_email = c10.text_input("Email ID")
                    new_aadhaar = c11.text_input("Aadhaar Number (Required)")
                    
                    c12, c13 = st.columns(2)
                    new_gname = c12.text_input("Guardian Name")
                    new_gmobile = c13.text_input("Guardian Mobile")
                    
                    new_address = st.text_area("Permanent Address")
                    
                    add_submit = st.form_submit_button("Add Student to Database", type="primary")
                    
                    if add_submit:
                        if not new_name or not new_umis or not new_aadhaar:
                            st.error("Name, UMIS, and Aadhaar are required fields!")
                        else:
                            cursor = conn.cursor()
                            max_sno = conn.execute("SELECT MAX(CAST(s_no AS INTEGER)) FROM students").fetchone()[0]
                            new_sno = 1 if max_sno is None else max_sno + 1
                            
                            cursor.execute("""
                                INSERT INTO students (
                                    s_no, student_name, branch, mobile_number, umis_number, aadhaar, email, status,
                                    dob, gender, blood_group, guardian_name, guardian_mobile, address, cgpa, attendance_percent
                                )
                                VALUES (?, ?, ?, ?, ?, ?, ?, 'Pending', ?, ?, ?, ?, ?, ?, ?, ?)
                            """, (
                                str(new_sno), new_name.strip(), new_branch.strip(), new_mobile.strip(), new_umis.strip(), 
                                new_aadhaar.strip(), new_email.strip(), new_dob, new_gender, new_blood, new_gname, 
                                new_gmobile, new_address, new_cgpa, new_attendance
                            ))
                            conn.commit()
                            st.balloons()
                            st.toast(f"✅ Added {new_name} successfully!", icon="🎈")
                            time.sleep(1.5)
                            st.rerun()
    
            # Edit / Delete Student
            with st.expander("✏️ Edit or Delete Student"):
                student_list = df_all['s_no'].astype(str) + " - " + df_all['student_name']
                selected_student_str = st.selectbox("Select a Student to Modify", options=student_list)
                
                if selected_student_str:
                    selected_sno_str = selected_student_str.split(" - ")[0]
                    # Cast dataframe column to string to ensure safe comparison
                    student_data = df_all[df_all['s_no'].astype(str) == selected_sno_str].iloc[0]
                    
                    with st.form("edit_student_form"):
                        st.markdown("#### Academic Info")
                        ec1, ec2, ec3 = st.columns(3)
                        mod_name = ec1.text_input("Full Name", value=student_data.get('student_name', ''))
                        mod_umis = ec2.text_input("UMIS Number", value=student_data.get('umis_number', ''))
                        mod_branch = ec3.text_input("Branch", value=student_data.get('branch', ''))
                        
                        ec4, ec5 = st.columns(2)
                        mod_cgpa = ec4.number_input("CGPA", value=float(student_data.get('cgpa', 0.0) or 0.0), min_value=0.0, max_value=10.0, step=0.1)
                        mod_attendance = ec5.number_input("Attendance %", value=float(student_data.get('attendance_percent', 0.0) or 0.0), min_value=0.0, max_value=100.0, step=1.0)
                        
                        st.markdown("#### Demographics")
                        ec6, ec7, ec8 = st.columns(3)
                        mod_dob = ec6.text_input("Date of Birth", value=student_data.get('dob', ''))
                        gen_idx = ["Male", "Female", "Other", ""].index(student_data.get('gender', '')) if student_data.get('gender', '') in ["Male", "Female", "Other", ""] else 3
                        mod_gender = ec7.selectbox("Gender", ["Male", "Female", "Other", ""], index=gen_idx)
                        mod_blood = ec8.text_input("Blood Group", value=student_data.get('blood_group', ''))
                        
                        st.markdown("#### Contact Info")
                        ec9, ec10, ec11 = st.columns(3)
                        mod_mobile = ec9.text_input("Mobile Number", value=student_data.get('mobile_number', ''))
                        mod_email = ec10.text_input("Email ID", value=student_data.get('email', ''))
                        mod_aadhaar = ec11.text_input("Aadhaar Number", value=student_data.get('aadhaar', ''))
                        
                        ec12, ec13 = st.columns(2)
                        mod_gname = ec12.text_input("Guardian Name", value=student_data.get('guardian_name', ''))
                        mod_gmobile = ec13.text_input("Guardian Mobile", value=student_data.get('guardian_mobile', ''))
                        
                        mod_address = st.text_area("Permanent Address", value=student_data.get('address', ''))
                        
                        mod_status = st.selectbox("Status", options=["Pending", "Updated"], index=0 if student_data.get('status', 'Pending') == "Pending" else 1)
                        
                        bc1, bc2 = st.columns(2)
                        update_submit = bc1.form_submit_button("Update Student", type="primary")
                        delete_submit = bc2.form_submit_button("Delete Student")
                        
                        if update_submit:
                            cursor = conn.cursor()
                            cursor.execute("""
                                UPDATE students 
                                SET student_name=?, branch=?, mobile_number=?, umis_number=?, aadhaar=?, email=?, status=?,
                                    dob=?, gender=?, blood_group=?, guardian_name=?, guardian_mobile=?, address=?, cgpa=?, attendance_percent=?
                                WHERE s_no=?
                            """, (
                                mod_name.strip(), mod_branch.strip(), mod_mobile.strip(), mod_umis.strip(), mod_aadhaar.strip(), 
                                mod_email.strip(), mod_status, mod_dob, mod_gender, mod_blood, mod_gname, mod_gmobile, 
                                mod_address, mod_cgpa, mod_attendance, selected_sno_str
                            ))
                            conn.commit()
                            if mod_status == "Updated":
                                st.balloons()
                            st.toast(f"✅ Updated {mod_name} successfully!", icon="💾")
                            time.sleep(1.5)
                            st.rerun()
                            
                        if delete_submit:
                            cursor = conn.cursor()
                            cursor.execute("DELETE FROM students WHERE s_no=?", (selected_sno_str,))
                            conn.commit()
                            st.toast(f"🗑️ Deleted {mod_name} successfully!", icon="🔥")
                            time.sleep(1.5)
                            st.rerun()
            
            st.markdown("<br>", unsafe_allow_html=True)
            # Interactive Filtering & Smart Search
            st.markdown("### 🔍 Filter Records & Smart Search")
            st.info("💡 Try Smart Search: 'defaulters in mca', 'attendance < 75', or 'updated profiles'")
            
            filter_col, search_col, smart_col = st.columns([1, 1, 2])
            status_filter = filter_col.selectbox("Filter by Status", ["All", "Pending", "Updated"])
            search_term = search_col.text_input("Search Name/Branch")
            smart_search = smart_col.text_input("✨ AI Smart Search", placeholder="e.g. mca students with low attendance")
            
            filtered_df = df_all.copy()
            if status_filter != "All":
                filtered_df = filtered_df[filtered_df['status'] == status_filter]
            
            if search_term:
                filtered_df = filtered_df[
                    filtered_df['student_name'].str.contains(search_term, case=False, na=False) |
                    filtered_df['branch'].str.contains(search_term, case=False, na=False)
                ]
                
            if smart_search:
                ss = smart_search.lower()
                if "defaulter" in ss or "pending fee" in ss or "due" in ss:
                    filtered_df = filtered_df[filtered_df['Total Due'] > 0]
                if "attendance <" in ss:
                    try:
                        val = int(re.search(r'attendance <\s*(\d+)', ss).group(1))
                        filtered_df = filtered_df[filtered_df['attendance_percent'] < val]
                    except:
                        pass
                elif "low attendance" in ss:
                    filtered_df = filtered_df[filtered_df['attendance_percent'] < 75.0]
                
                if "mca" in ss:
                    filtered_df = filtered_df[filtered_df['branch'].str.contains("MCA", case=False, na=False)]
                if "ai" in ss or "ds" in ss:
                    filtered_df = filtered_df[filtered_df['branch'].str.contains("AI", case=False, na=False) | filtered_df['branch'].str.contains("DS", case=False, na=False)]
                
                if "updated" in ss:
                    filtered_df = filtered_df[filtered_df['status'] == 'Updated']
                if "pending" in ss and "fee" not in ss:
                    filtered_df = filtered_df[filtered_df['status'] == 'Pending']
            
            st.markdown("### 📋 Student Directory")
            st.write("💡 *Tip: You can edit any cell directly, add new rows at the bottom, or select rows to delete.*")
            
            # Interactive Data Editor
            edited_df = st.data_editor(
                filtered_df, 
                width='stretch', 
                hide_index=True,
                num_rows="dynamic",
                column_config={
                    "status": st.column_config.SelectboxColumn(
                        "status",
                        options=["Pending", "Updated"],
                        required=True,
                    ),
                    "s_no": st.column_config.NumberColumn(
                        "s_no",
                        disabled=True
                    )
                },
                key="advanced_editor"
            )
            
            if st.button("💾 Sync Grid Changes to Database", type="primary"):
                cursor = conn.cursor()
                
                # Identify Deletes
                original_ids = set(filtered_df['s_no'].dropna())
                current_ids = set(edited_df['s_no'].dropna())
                deleted_ids = original_ids - current_ids
                for s_no in deleted_ids:
                    cursor.execute("DELETE FROM students WHERE s_no=?", (s_no,))
                    
                # Identify Inserts and Updates
                for _, row in edited_df.iterrows():
                    row_dict = row.to_dict()
                    s_no = row_dict.pop('s_no', None)
                    row_dict.pop('Total Due', None)
                    if pd.isna(s_no): # New Row
                        cols = list(row_dict.keys())
                        vals = list(row_dict.values())
                        placeholders = ",".join(["?"] * len(cols))
                        cursor.execute(f"INSERT INTO students ({','.join(cols)}) VALUES ({placeholders})", vals)
                    else:
                        cols = list(row_dict.keys())
                        vals = list(row_dict.values())
                        set_clause = ", ".join([f"{c}=?" for c in cols])
                        cursor.execute(f"UPDATE students SET {set_clause} WHERE s_no=?", vals + [s_no])
                        
                conn.commit()
                st.success("Database synchronized successfully!")
                time.sleep(1)
                st.rerun()
            
            # Charts section
            st.markdown("### 📊 Status Breakdown")
            status_counts = df_all['status'].value_counts().reset_index()
            status_counts.columns = ['Status', 'Count']
            st.bar_chart(status_counts.set_index('Status'), color=["#4b9fff"])
            
            st.markdown("### 💾 Export Filtered Data")
            c_exp1, c_exp2 = st.columns(2)
            csv = filtered_df.to_csv(index=False).encode("utf-8")
            c_exp1.download_button(
                label="📥 Download Filtered CSV",
                data=csv,
                file_name="Filtered_MCA_Students.csv",
                mime="text/csv",
                type="primary",
                use_container_width=True
            )
            
            excel_io = io.BytesIO()
            filtered_df.to_excel(excel_io, index=False, engine='openpyxl')
            c_exp2.download_button(
                label="📊 Download Excel Report",
                data=excel_io.getvalue(),
                file_name="Filtered_MCA_Students.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        else:
            st.warning("No student records found in the database. Please ensure the Excel file is populated.")
    
    with tab2:
        st.markdown("### 💰 Fees Management")
        st.write("Manage College and Exam fees for different departments.")
        
        if not df_all.empty:
            fc1, fc2 = st.columns(2)
            with fc1:
                fee_dept_filter = st.selectbox("Select Department", ["All", "MCA (1st Year)", "AI&DS (3rd Year)"], key="fee_dept_filter")
            with fc2:
                fee_search = st.text_input("Search Student by Name", key="fee_search")
                
            fee_df = df_all.copy()
            if fee_dept_filter == "MCA (1st Year)":
                fee_df = fee_df[fee_df['branch'].str.contains("MCA", case=False, na=False)]
            elif fee_dept_filter == "AI&DS (3rd Year)":
                fee_df = fee_df[fee_df['branch'].str.contains("AI", case=False, na=False) | fee_df['branch'].str.contains("DS", case=False, na=False)]
                
            if fee_search:
                fee_df = fee_df[fee_df['student_name'].str.contains(fee_search, case=False, na=False)]
                
            st.markdown("#### 💸 Fee Status Overview")
            
            # Show interactive grid for fees
            fee_display_cols = ['s_no', 'student_name', 'branch', 'college_fee_total', 'college_fee_paid', 'exam_fee_total', 'exam_fee_paid']
            
            # Ensure columns exist in dataframe (handle missing values)
            for col in ['college_fee_total', 'college_fee_paid', 'exam_fee_total', 'exam_fee_paid']:
                if col not in fee_df.columns:
                    fee_df[col] = 0.0
                fee_df[col] = pd.to_numeric(fee_df[col], errors='coerce').fillna(0.0)
                
            edited_fee_df = st.data_editor(
                fee_df[fee_display_cols],
                width='stretch',
                hide_index=True,
                disabled=['s_no', 'student_name', 'branch'],
                key="fee_editor"
            )
            
            # Detect changes and save
            fee_changed = False
            for idx in edited_fee_df.index:
                row_old = fee_df.loc[idx]
                row_new = edited_fee_df.loc[idx]
                if (row_old['college_fee_total'] != row_new['college_fee_total'] or
                    row_old['college_fee_paid'] != row_new['college_fee_paid'] or
                    row_old['exam_fee_total'] != row_new['exam_fee_total'] or
                    row_old['exam_fee_paid'] != row_new['exam_fee_paid']):
                    
                    cursor = conn.cursor()
                    cursor.execute(
                        "UPDATE students SET college_fee_total=?, college_fee_paid=?, exam_fee_total=?, exam_fee_paid=? WHERE s_no=?",
                        (row_new['college_fee_total'], row_new['college_fee_paid'], row_new['exam_fee_total'], row_new['exam_fee_paid'], row_new['s_no'])
                    )
                    fee_changed = True
                    
            if fee_changed:
                conn.commit()
                st.toast("✅ Fee details updated successfully!", icon="💰")
                time.sleep(1.0)
                st.rerun()
                
            # Summary metrics
            total_college = fee_df['college_fee_total'].sum()
            paid_college = fee_df['college_fee_paid'].sum()
            total_exam = fee_df['exam_fee_total'].sum()
            paid_exam = fee_df['exam_fee_paid'].sum()
            
            sc1, sc2, sc3 = st.columns(3)
            sc1.metric("College Fees Collected", f"₹ {paid_college:,.2f}", f"Due: ₹ {total_college - paid_college:,.2f}", delta_color="inverse")
            sc2.metric("Exam Fees Collected", f"₹ {paid_exam:,.2f}", f"Due: ₹ {total_exam - paid_exam:,.2f}", delta_color="inverse")
            sc3.metric("Total Outstanding", f"₹ {(total_college - paid_college) + (total_exam - paid_exam):,.2f}")
            
    with tab3:
        st.markdown("### 📂 Bulk Operations")
        
        st.markdown("#### 1. Bulk Upload Students (Excel)")
        st.write("Upload an Excel file containing a list of new students.")
        
        uploaded_file = st.file_uploader("Upload Student List (.xlsx)", type=['xlsx'])
        if uploaded_file is not None:
            try:
                upload_df = pd.read_excel(uploaded_file)
                st.write("Preview:")
                st.dataframe(upload_df.head())
                
                name_col = st.selectbox("Which column contains the Student Name?", options=upload_df.columns)
                target_dept = st.selectbox("Assign to Department", ["AI&DS (3rd Year)", "MCA (1st Year)"])
                
                if st.button("Process & Import", type="primary"):
                    branch_val = "AI&DS" if "AI&DS" in target_dept else "MCA"
                    year_val = "3" if "3rd" in target_dept else "1"
                    course_val = "UG" if "AI&DS" in target_dept else "PG"
                    
                    cursor = conn.cursor()
                    max_sno_res = cursor.execute("SELECT MAX(CAST(s_no AS INTEGER)) FROM students").fetchone()[0]
                    next_sno = 1 if max_sno_res is None else max_sno_res + 1
                    
                    added_count = 0
                    for _, row in upload_df.iterrows():
                        student_name = str(row[name_col]).strip()
                        if student_name and student_name != 'nan':
                            cursor.execute('''
                                INSERT INTO students (s_no, student_name, branch, year_of_study, course_type, status, college_fee_total, college_fee_paid, exam_fee_total, exam_fee_paid)
                                VALUES (?, ?, ?, ?, ?, 'Pending', 0.0, 0.0, 0.0, 0.0)
                            ''', (str(next_sno), student_name, branch_val, year_val, course_val))
                            next_sno += 1
                            added_count += 1
                            
                    conn.commit()
                    st.balloons()
                    st.success(f"✅ Successfully imported {added_count} students into {target_dept}!")
                    time.sleep(2)
                    st.rerun()
            except Exception as e:
                st.error(f"Error processing file: {e}")

        st.markdown("---")
        st.markdown("#### 2. Automated Bank Reconciliation (PDF)")
        st.write("Upload bank receipt PDFs. The system will automatically extract UMIS and Paid Amounts to reconcile records.")
        
        pdf_file = st.file_uploader("Upload Bank Receipt (.pdf)", type=['pdf'])
        if pdf_file is not None:
            try:
                with pdfplumber.open(pdf_file) as pdf:
                    text = ""
                    for page in pdf.pages:
                        text += page.extract_text() + "\n"
                        
                st.text_area("Extracted Text Preview", text, height=100)
                
                # Heuristic parsing
                # Looking for UMIS pattern (e.g. 6 digits) and Amount (INR / Rs / ₹ followed by numbers)
                umis_matches = re.findall(r'(?i)(?:umis|id|enrollment).*?(\d{5,8})', text)
                amount_matches = re.findall(r'(?i)(?:inr|rs\.?|amount|paid).*?([\d,]+\.?\d*)', text)
                
                if umis_matches and amount_matches:
                    detected_umis = umis_matches[0]
                    detected_amount = float(amount_matches[0].replace(',', ''))
                    
                    st.success(f"**AI Detected:** UMIS = `{detected_umis}`, Amount Paid = `₹{detected_amount}`")
                    
                    # Cross-reference with DB
                    student = df_all[df_all['umis_number'] == detected_umis]
                    if not student.empty:
                        student_name = student.iloc[0]['student_name']
                        s_no = student.iloc[0]['s_no']
                        current_paid = float(student.iloc[0]['college_fee_paid'] or 0)
                        
                        st.info(f"Student Found: **{student_name}**")
                        
                        if st.button(f"Reconcile ₹{detected_amount} for {student_name}", type="primary"):
                            cursor = conn.cursor()
                            new_paid = current_paid + detected_amount
                            cursor.execute("UPDATE students SET college_fee_paid=? WHERE s_no=?", (new_paid, s_no))
                            conn.commit()
                            st.balloons()
                            st.success(f"Successfully updated {student_name}'s fee record!")
                            time.sleep(1.5)
                            st.rerun()
                    else:
                        st.warning("Detected UMIS does not match any student in the database.")
                else:
                    st.warning("Could not clearly extract UMIS and Amount from the PDF. Please check the document format.")
            except Exception as e:
                st.error(f"Error parsing PDF: {e}")

    with tab4:
        st.markdown("### 📈 Analytics & Insights")
        if not df_all.empty:
            st.write("Visual breakdown of student data.")
            
            df_all['Fee Status'] = df_all['Total Due'].apply(lambda x: 'Defaulter' if x > 0 else 'Fully Paid')
            
            ac1, ac2 = st.columns(2)
            
            with ac1:
                st.subheader("Fee Payment Status")
                fee_chart = alt.Chart(df_all).mark_arc(innerRadius=50).encode(
                    theta=alt.Theta(field="Fee Status", aggregate="count"),
                    color=alt.Color(field="Fee Status", type="nominal", scale=alt.Scale(domain=['Fully Paid', 'Defaulter'], range=['#2e7d32', '#d32f2f'])),
                    tooltip=['Fee Status', 'count()']
                ).interactive()
                st.altair_chart(fee_chart, use_container_width=True)
                
            with ac2:
                st.subheader("Attendance Distribution")
                att_chart = alt.Chart(df_all).mark_bar().encode(
                    x=alt.X("attendance_percent:Q", bin=alt.Bin(maxbins=10), title="Attendance %"),
                    y=alt.Y("count()", title="Number of Students"),
                    color=alt.value("#4b9fff"),
                    tooltip=['count()']
                ).interactive()
                st.altair_chart(att_chart, use_container_width=True)
                
            st.subheader("CGPA vs Attendance")
            scatter_chart = alt.Chart(df_all).mark_circle(size=60).encode(
                x=alt.X("attendance_percent:Q", title="Attendance %"),
                y=alt.Y("cgpa:Q", title="CGPA"),
                color=alt.Color("branch:N", title="Branch"),
                tooltip=["student_name", "branch", "cgpa", "attendance_percent"]
            ).interactive()
            st.altair_chart(scatter_chart, use_container_width=True)
            
        else:
            st.info("No data available for analytics.")

    with tab5:
        st.markdown("### 📧 Email Automations")
        st.write("Send automated emails to students based on their status.")
        
        st.info("Emails will be sent using the SMTP credentials configured in Streamlit Secrets. If not found, it runs in 'Simulation Mode'.")
        
        target_group = st.radio("Select Target Group", ["Fee Defaulters", "Low Attendance (< 75%)"])
        
        if st.button(f"Prepare Emails for {target_group}", type="primary"):
            if target_group == "Fee Defaulters":
                targets = df_all[(df_all['Total Due'] > 0) & (df_all['email'].str.contains('@', na=False))]
                subject = "Action Required: Pending College Fees"
                body_template = "Dear {name},\n\nYou have pending fees of INR {due:,.2f}. Please clear your dues immediately."
            else:
                targets = df_all[(df_all['attendance_percent'] < 75.0) & (df_all['email'].str.contains('@', na=False))]
                subject = "Warning: Low Attendance"
                body_template = "Dear {name},\n\nYour attendance is currently {att}%, which is below the required 75%. Please meet your advisor."
                
            if targets.empty:
                st.success("No students found in this category with a valid email address.")
            else:
                st.write(f"**Found {len(targets)} students.**")
                import urllib.parse
                
                st.info("Since direct SMTP is disabled, click the buttons below to open your default email app (like Gmail or Outlook). The email will be automatically pre-filled for you!")
                
                for i, (_, row) in enumerate(targets.iterrows()):
                    body = body_template.format(
                        name=row['student_name'], 
                        due=row.get('Total Due', 0), 
                        att=row.get('attendance_percent', 0)
                    )
                    
                    # URL encode the subject and body for the mailto link
                    subject_enc = urllib.parse.quote(subject)
                    body_enc = urllib.parse.quote(body)
                    mailto_link = f"mailto:{row['email']}?subject={subject_enc}&body={body_enc}"
                    
                    col1, col2 = st.columns([3, 1])
                    col1.write(f"**{row['student_name']}** ({row['email']})")
                    col2.link_button(f"📧 Send Email", mailto_link, use_container_width=True)
                    
                st.success(f"Generated email links for {len(targets)} students. Click them to send!")

# Close the global connection at the end of the script run
if 'conn' in locals():
    try:
        conn.close()
    except Exception:
        pass
