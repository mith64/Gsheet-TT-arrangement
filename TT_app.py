import streamlit as st
import pandas as pd
import hashlib
import json
import os
from datetime import datetime
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

# ============================================
# GOOGLE SHEETS CONFIGURATION
# ============================================

# Try to import Google Sheets dependencies
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    GSHEETS_AVAILABLE = True
except ImportError:
    GSHEETS_AVAILABLE = False

# Google Sheets configuration
SHEET_NAME = "TimetableManagementSystem"
CREDENTIALS_FILE = "credentials.json"  # Download from Google Cloud Console

# Worksheet names
WORKSHEET_USERS = "users"
WORKSHEET_TIMETABLE = "timetable"
WORKSHEET_ARRANGEMENTS = "arrangements"

# ============================================
# SESSION STATE INITIALIZATION
# ============================================

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'username' not in st.session_state:
    st.session_state.username = None
if 'role' not in st.session_state:
    st.session_state.role = None
if 'timetable_df' not in st.session_state:
    st.session_state.timetable_df = None
if 'password_changed' not in st.session_state:
    st.session_state.password_changed = False
if 'show_password_change' not in st.session_state:
    st.session_state.show_password_change = False
if 'use_google_sheets' not in st.session_state:
    st.session_state.use_google_sheets = False

# ============================================
# GOOGLE SHEETS CONNECTION
# ============================================

@st.cache_resource
def get_gsheets_connection():
    """Establish and cache connection to Google Sheets"""
    if not GSHEETS_AVAILABLE:
        st.error("❌ gspread or oauth2client not installed!")
        st.info("Run: pip install gspread oauth2client")
        return None
    
    try:
        # Check if credentials file exists
        if not os.path.exists(CREDENTIALS_FILE):
            st.error(f"❌ Credentials file '{CREDENTIALS_FILE}' not found!")
            st.info("""
            Please download your service account JSON file from Google Cloud Console:
            1. Go to Google Cloud Console → APIs & Services → Credentials
            2. Create a Service Account and download the JSON key
            3. Save it as 'credentials.json' in your app directory
            """)
            return None
        
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
        
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)
        
        try:
            spreadsheet = client.open(SHEET_NAME)
        except gspread.SpreadsheetNotFound:
            spreadsheet = client.create(SHEET_NAME)
            st.success(f"✅ Created new spreadsheet: {SHEET_NAME}")
        
        return spreadsheet
    
    except Exception as e:
        st.error(f"Failed to connect to Google Sheets: {e}")
        return None

def get_or_create_worksheet(spreadsheet, worksheet_name, headers=None):
    """Get existing worksheet or create new one"""
    try:
        worksheet = spreadsheet.worksheet(worksheet_name)
        return worksheet
    except gspread.WorksheetNotFound:
        if headers:
            worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows=1000, cols=len(headers))
            worksheet.append_row(headers)
            return worksheet
        else:
            return spreadsheet.add_worksheet(title=worksheet_name, rows=1000, cols=10)

def load_from_gsheets(worksheet_name):
    """Load data from Google Sheets worksheet"""
    spreadsheet = get_gsheets_connection()
    if spreadsheet is None:
        return None
    
    try:
        worksheet = get_or_create_worksheet(spreadsheet, worksheet_name)
        data = worksheet.get_all_records()
        if data:
            return pd.DataFrame(data)
        else:
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Error loading from Google Sheets: {e}")
        return None

def save_to_gsheets(df, worksheet_name):
    """Save DataFrame to Google Sheets worksheet"""
    spreadsheet = get_gsheets_connection()
    if spreadsheet is None:
        return False
    
    try:
        # Get or create worksheet
        worksheet = get_or_create_worksheet(spreadsheet, worksheet_name)
        
        # Clear existing data
        worksheet.clear()
        
        # Update with new data
        if not df.empty:
            # Convert DataFrame to list of lists
            data = [df.columns.tolist()] + df.values.tolist()
            worksheet.update(data, value_input_option='USER_ENTERED')
        
        return True
    except Exception as e:
        st.error(f"Error saving to Google Sheets: {e}")
        return False
# Add this function to your code
def get_gsheets_connection_from_secrets():
    """Connect to Google Sheets using Streamlit secrets"""
    try:
        service_account_info = dict(st.secrets["connections"]["gsheets"])
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
        client = gspread.authorize(creds)
        return client.open(SHEET_NAME)
    except Exception as e:
        st.error(f"Failed to connect: {e}")
        return None
import streamlit as st
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Detect deployment environment
IS_STREAMLIT_CLOUD = os.environ.get('IS_STREAMLIT_CLOUD', False)

def get_gsheets_connection():
    """Connect to Google Sheets - works both locally and on Streamlit Cloud"""
    try:
        if IS_STREAMLIT_CLOUD or 'connections' in st.secrets:
            # Use secrets from Streamlit Cloud
            service_account_info = dict(st.secrets["connections"]["gsheets"])
            scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
            creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
            client = gspread.authorize(creds)
        else:
            # Local development - use credentials.json file
            scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
            creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
            client = gspread.authorize(creds)
        
        return client.open(SHEET_NAME)
    except Exception as e:
        st.error(f"Failed to connect to Google Sheets: {e}")
        return None
# ============================================
# DATA STORAGE FUNCTIONS (Google Sheets + Local)
# ============================================

def load_users():
    """Load users from Google Sheets or local JSON"""
    if st.session_state.use_google_sheets and GSHEETS_AVAILABLE:
        df = load_from_gsheets(WORKSHEET_USERS)
        if df is not None and not df.empty:
            # Convert DataFrame to dict
            users = {}
            for _, row in df.iterrows():
                users[row['username']] = {
                    'password': row['password'],
                    'name': row['name'],
                    'designation': row['designation'],
                    'role': row['role'],
                    'first_login': row.get('first_login', True),
                    'password_last_changed': row.get('password_last_changed', '')
                }
            return users
    
    # Fallback to local JSON
    try:
        if os.path.exists(USER_DB_FILE):
            with open(USER_DB_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            default_users = {
                "admin": {
                    "password": hash_password("admin123"),
                    "name": "Administrator",
                    "designation": "Admin",
                    "role": "admin",
                    "first_login": True,
                    "password_last_changed": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
            }
            save_users(default_users)
            return default_users
    except Exception as e:
        st.error(f"Error loading users: {e}")
        return {"admin": {"password": hash_password("admin123"), "name": "Admin", "designation": "Admin", "role": "admin", "first_login": True}}

def save_users(users):
    """Save users to Google Sheets or local JSON"""
    if st.session_state.use_google_sheets and GSHEETS_AVAILABLE:
        # Convert dict to DataFrame
        data = []
        for username, info in users.items():
            data.append({
                'username': username,
                'password': info.get('password', ''),
                'name': info.get('name', ''),
                'designation': info.get('designation', ''),
                'role': info.get('role', 'user'),
                'first_login': info.get('first_login', True),
                'password_last_changed': info.get('password_last_changed', '')
            })
        df = pd.DataFrame(data)
        return save_to_gsheets(df, WORKSHEET_USERS)
    
    # Fallback to local JSON
    try:
        with open(USER_DB_FILE, 'w', encoding='utf-8') as f:
            json.dump(users, f, indent=4)
        return True
    except Exception as e:
        st.error(f"Error saving users: {e}")
        return False

def load_timetable():
    """Load timetable from Google Sheets or Excel"""
    if st.session_state.use_google_sheets and GSHEETS_AVAILABLE:
        df = load_from_gsheets(WORKSHEET_TIMETABLE)
        if df is not None and not df.empty:
            st.session_state.timetable_df = df
            return df
    
    # Fallback to Excel
    try:
        if os.path.exists(TIMETABLE_FILE):
            df = pd.read_excel(TIMETABLE_FILE, engine='openpyxl')
            if not df.empty:
                st.session_state.timetable_df = df
                return df
    except Exception as e:
        st.warning(f"Could not read Excel file: {e}")
    
    # Create sample if no data
    return create_sample_timetable()

def save_timetable(df):
    """Save timetable to Google Sheets or Excel"""
    if st.session_state.use_google_sheets and GSHEETS_AVAILABLE:
        return save_to_gsheets(df, WORKSHEET_TIMETABLE)
    
    # Fallback to Excel
    try:
        df.to_excel(TIMETABLE_FILE, index=False, engine='openpyxl')
        st.session_state.timetable_df = df
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Error saving timetable: {e}")
        return False

def load_arrangements():
    """Load arrangements from Google Sheets or JSON"""
    if st.session_state.use_google_sheets and GSHEETS_AVAILABLE:
        df = load_from_gsheets(WORKSHEET_ARRANGEMENTS)
        if df is not None and not df.empty:
            # Convert DataFrame to dict
            arrangements = {}
            for _, row in df.iterrows():
                key = row.get('key', '')
                if key:
                    arrangements[key] = {
                        'date': row.get('date', ''),
                        'absent_teacher': row.get('absent_teacher', ''),
                        'day': row.get('day', ''),
                        'leave_type': row.get('leave_type', ''),
                        'reason': row.get('reason', ''),
                        'status': row.get('status', 'active'),
                        'total_periods': row.get('total_periods', 0),
                        'covered_periods': row.get('covered_periods', 0),
                        'uncovered_periods': row.get('uncovered_periods', 0)
                    }
            return arrangements
    
    # Fallback to local JSON
    try:
        if os.path.exists(ARRANGEMENT_FILE):
            with open(ARRANGEMENT_FILE, 'r', encoding='utf-8') as f:
                content = f.read()
                if content.strip():
                    return json.loads(content)
        return {}
    except Exception as e:
        st.error(f"Error loading arrangements: {e}")
        return {}

def save_arrangements(arrangements):
    """Save arrangements to Google Sheets or JSON"""
    if st.session_state.use_google_sheets and GSHEETS_AVAILABLE:
        # Convert dict to DataFrame
        data = []
        for key, info in arrangements.items():
            data.append({
                'key': key,
                'date': info.get('date', ''),
                'absent_teacher': info.get('absent_teacher', ''),
                'day': info.get('day', ''),
                'leave_type': info.get('leave_type', ''),
                'reason': info.get('reason', ''),
                'status': info.get('status', 'active'),
                'total_periods': info.get('total_periods', 0),
                'covered_periods': info.get('covered_periods', 0),
                'uncovered_periods': info.get('uncovered_periods', 0)
            })
        df = pd.DataFrame(data)
        return save_to_gsheets(df, WORKSHEET_ARRANGEMENTS)
    
    # Fallback to local JSON
    try:
        with open(ARRANGEMENT_FILE, 'w', encoding='utf-8') as f:
            json.dump(arrangements, f, indent=4)
        return True
    except Exception as e:
        st.error(f"Error saving arrangements: {e}")
        return False

# ============================================
# HASH FUNCTIONS
# ============================================

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

# ============================================
# TIMETABLE FUNCTIONS WITH DUPLICATE DETECTION
# ============================================

def detect_duplicates(df, new_entry):
    """Detect duplicate entries in timetable"""
    duplicates = []
    
    if df.empty:
        return duplicates
    
    # Check for exact duplicate (same day, time, teacher, subject, class)
    for idx, row in df.iterrows():
        if (row['Day'] == new_entry['Day'] and 
            row['Time'] == new_entry['Time'] and 
            row['Teacher'] == new_entry['Teacher'] and 
            row['Subject'] == new_entry['Subject'] and 
            row['Class'] == new_entry['Class']):
            duplicates.append({
                'row_index': idx,
                'existing_entry': row.to_dict()
            })
    
    return duplicates

def update_timetable_entry(df, row_index, new_data):
    """Update a specific row in the timetable"""
    for col in new_data.keys():
        if col in df.columns:
            df.at[row_index, col] = new_data[col]
    return df

def delete_timetable_entry(df, row_index):
    """Delete a specific row from the timetable"""
    return df.drop(row_index).reset_index(drop=True)

# ============================================
# EXISTING FUNCTIONS (Keep from previous code)
# ============================================

# [Keep all the existing helper functions from your code]
# create_directories, get_period_index, get_time_slots_for_half,
# is_teacher_available, find_available_teachers, suggest_teacher_for_period,
# suggest_half_day_arrangements, create_sample_timetable, etc.

# ============================================
# ADMIN PANEL WITH GOOGLE SHEETS AND DUPLICATE HANDLING
# ============================================

def admin_panel():
    st.header("👑 Admin Panel")
    
    # Google Sheets Toggle
    with st.expander("⚙️ Storage Settings", expanded=False):
        if GSHEETS_AVAILABLE:
            use_gsheets = st.checkbox("Use Google Sheets", value=st.session_state.use_google_sheets)
            if use_gsheets != st.session_state.use_google_sheets:
                st.session_state.use_google_sheets = use_gsheets
                st.cache_data.clear()
                st.rerun()
            
            if use_gsheets:
                if os.path.exists(CREDENTIALS_FILE):
                    st.success("✅ Google Sheets configured")
                else:
                    st.error("❌ credentials.json not found!")
            else:
                st.info("Using local storage (JSON/Excel)")
        else:
            st.warning("⚠️ Google Sheets libraries not installed")
            st.code("pip install gspread oauth2client")
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Create User", "Manage Users", "Upload Timetable", "Arrangement Management", "Security Settings"])
    
    # [Keep existing Create User, Manage Users tabs]
    # [Keep existing Upload Timetable with duplicate handling]
    
    # Upload Timetable Tab - With Duplicate Detection
    with tab3:
        st.subheader("Upload Timetable")
        
        if not OPENPYXL_AVAILABLE:
            st.error("❌ openpyxl is not installed!")
            st.code("pip install openpyxl", language="bash")
            return
        
        col1, col2 = st.columns(2)
        with col1:
            storage_type = "Google Sheets" if st.session_state.use_google_sheets else "Local Excel"
            st.info(f"📁 Storage: {storage_type}")
        
        with col2:
            if st.button("🗑️ Clear Timetable", type="secondary"):
                if st.session_state.use_google_sheets:
                    save_to_gsheets(pd.DataFrame(), WORKSHEET_TIMETABLE)
                else:
                    delete_timetable_file()
                st.cache_data.clear()
                st.success("Timetable cleared!")
                st.rerun()
        
        st.markdown("---")
        
        st.info("""
        **📋 Required Excel columns:**
        - Day (Monday, Tuesday, etc.)
        - Time (9:00-10:00 format)
        - Teacher (Teacher's name)
        - Subject (Subject name)
        - Class (Class name)
        - Designation (Math Teacher, etc.)
        """)
        
        uploaded_file = st.file_uploader(
            "Choose Excel file", 
            type=['xlsx', 'xls'],
            help="Upload a new timetable file"
        )
        
        if uploaded_file is not None:
            try:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
                
                required_cols = ['Day', 'Time', 'Teacher', 'Subject', 'Class', 'Designation']
                missing_cols = [col for col in required_cols if col not in df.columns]
                
                if missing_cols:
                    st.error(f"❌ Missing columns: {missing_cols}")
                else:
                    st.subheader("Preview of uploaded data:")
                    st.dataframe(df.head())
                    
                    # Check for duplicates within the uploaded file
                    duplicate_rows = []
                    for idx, row in df.iterrows():
                        entry = row.to_dict()
                        existing_duplicates = detect_duplicates(df.iloc[:idx], entry)
                        if existing_duplicates:
                            duplicate_rows.append(idx)
                    
                    if duplicate_rows:
                        st.warning(f"⚠️ Found {len(duplicate_rows)} duplicate rows in the uploaded file!")
                        with st.expander("View Duplicate Rows"):
                            st.dataframe(df.iloc[duplicate_rows])
                    
                    # Check for duplicates with existing timetable
                    existing_df = load_timetable()
                    duplicate_with_existing = []
                    
                    if not existing_df.empty:
                        for idx, row in df.iterrows():
                            entry = row.to_dict()
                            duplicates = detect_duplicates(existing_df, entry)
                            if duplicates:
                                duplicate_with_existing.append({
                                    'row_index': idx,
                                    'day': entry['Day'],
                                    'time': entry['Time'],
                                    'teacher': entry['Teacher'],
                                    'subject': entry['Subject'],
                                    'class': entry['Class']
                                })
                        
                        if duplicate_with_existing:
                            st.error(f"🚨 Found {len(duplicate_with_existing)} duplicates with existing timetable!")
                            
                            with st.expander("View Duplicates with Existing Timetable"):
                                dup_df = pd.DataFrame(duplicate_with_existing)
                                st.dataframe(dup_df)
                            
                            # Options for handling duplicates
                            st.subheader("How to handle duplicates?")
                            duplicate_action = st.radio(
                                "Select action:",
                                ["Skip duplicates (keep existing)", "Replace all with new data", "Show duplicates for manual review"],
                                horizontal=True
                            )
                            
                            if duplicate_action == "Skip duplicates (keep existing)":
                                # Filter out duplicates
                                df_filtered = df.copy()
                                for dup in duplicate_with_existing:
                                    df_filtered = df_filtered[~((df_filtered['Day'] == dup['day']) & 
                                                                  (df_filtered['Time'] == dup['time']) & 
                                                                  (df_filtered['Teacher'] == dup['teacher']) & 
                                                                  (df_filtered['Subject'] == dup['subject']) & 
                                                                  (df_filtered['Class'] == dup['class']))]
                                
                                st.info(f"Will upload {len(df_filtered)} new entries (skipped {len(duplicate_with_existing)} duplicates)")
                                
                                if st.button("✅ Upload New Entries Only", type="primary"):
                                    combined_df = pd.concat([existing_df, df_filtered], ignore_index=True)
                                    if save_timetable(combined_df):
                                        st.success(f"✅ Uploaded {len(df_filtered)} new entries successfully!")
                                        st.balloons()
                                        time.sleep(1)
                                        st.rerun()
                            
                            elif duplicate_action == "Replace all with new data":
                                if st.button("⚠️ Replace Entire Timetable", type="primary"):
                                    if save_timetable(df):
                                        st.success("✅ Timetable replaced successfully!")
                                        st.balloons()
                                        time.sleep(1)
                                        st.rerun()
                            
                            else:
                                # Manual review
                                st.info("Review duplicates below and choose which to keep")
                                
                                # Show comparison
                                for dup in duplicate_with_existing[:5]:
                                    st.write(f"**Duplicate found:**")
                                    col1, col2 = st.columns(2)
                                    with col1:
                                        st.write("🟢 **New Entry:**")
                                        st.write(f"Day: {dup['day']}")
                                        st.write(f"Time: {dup['time']}")
                                        st.write(f"Teacher: {dup['teacher']}")
                                        st.write(f"Subject: {dup['subject']}")
                                        st.write(f"Class: {dup['class']}")
                                    with col2:
                                        st.write("🔵 **Existing Entry:**")
                                        existing_row = existing_df[(existing_df['Day'] == dup['day']) & 
                                                                    (existing_df['Time'] == dup['time']) & 
                                                                    (existing_df['Teacher'] == dup['teacher']) & 
                                                                    (existing_df['Subject'] == dup['subject']) & 
                                                                    (existing_df['Class'] == dup['class'])]
                                        if not existing_row.empty:
                                            for col in existing_row.columns:
                                                st.write(f"{col}: {existing_row.iloc[0][col]}")
                                    st.markdown("---")
                    else:
                        if st.button("✅ Upload Timetable", type="primary"):
                            if save_timetable(df):
                                st.success("✅ Timetable uploaded successfully!")
                                st.balloons()
                                time.sleep(1)
                                st.rerun()
                            
            except Exception as e:
                st.error(f"Error reading file: {e}")
    
    # [Keep existing Arrangement Management and Security Settings tabs]

# ============================================
# MAIN APP
# ============================================

def main():
    st.set_page_config(
        page_title="Timetable Management System",
        page_icon="📚",
        layout="wide"
    )
    
    # Check Google Sheets availability
    if GSHEETS_AVAILABLE:
        st.sidebar.markdown("---")
        if os.path.exists(CREDENTIALS_FILE):
            st.sidebar.success("🔗 Google Sheets: Connected")
        else:
            st.sidebar.warning("🔗 Google Sheets: No credentials")
    
    st.title("📚 Timetable Management System with Smart Leave Management")
    
    if not st.session_state.logged_in:
        st.subheader("Login")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            with st.form("login_form"):
                username = st.text_input("Username")
                password = st.text_input("Password", type="password")
                submit = st.form_submit_button("Login")
                
                if submit:
                    if login(username, password):
                        st.success(f"Welcome {st.session_state.name}!")
                        st.rerun()
                    else:
                        st.error("Invalid username or password!")
            
            st.markdown("---")
            st.caption("Demo Credentials:")
            st.caption("Admin: admin / admin123")
            st.caption("*Note: Admin will be prompted to change password on first login*")
    else:
        # Check if password change is required
        if st.session_state.show_password_change and not st.session_state.password_changed:
            password_change_form()
        else:
            with st.sidebar:
                st.write(f"**Logged in as:** {st.session_state.name}")
                st.write(f"**Username:** {st.session_state.username}")
                st.write(f"**Role:** {st.session_state.role}")
                st.markdown("---")
                
                if st.button("🚪 Logout"):
                    logout()
                
                st.markdown("---")
                st.caption(f"Login Time: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
            
            if st.session_state.role == 'admin':
                admin_panel()
                user_dashboard()
            else:
                user_dashboard()

if __name__ == "__main__":
    main()
