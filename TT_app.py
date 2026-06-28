import streamlit as st
import pandas as pd
import hashlib
import json
import os
import shutil
import tempfile
from datetime import datetime
import time
import gc
from collections import defaultdict, Counter
import numpy as np

# Try importing optional packages
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    st.error("openpyxl not installed. Please add it to requirements.txt")

try:
    import xlsxwriter
    XLSXWRITER_AVAILABLE = True
except ImportError:
    XLSXWRITER_AVAILABLE = False

# File paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USER_DB_FILE = os.path.join(BASE_DIR, "users.json")
TIMETABLE_FILE = os.path.join(BASE_DIR, "timetable.xlsx")
ARRANGEMENT_FILE = os.path.join(BASE_DIR, "arrangements.json")
BACKUP_FOLDER = os.path.join(BASE_DIR, "backups")
CLASSROOM_FILE = os.path.join(BASE_DIR, "classrooms.json")

# Initialize session state
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
if 'editing_mode' not in st.session_state:
    st.session_state.editing_mode = False
if 'edit_df' not in st.session_state:
    st.session_state.edit_df = None
if 'school_id' not in st.session_state:
    st.session_state.school_id = None

# Create necessary directories
def create_directories():
    """Create necessary directories if they don't exist"""
    try:
        if not os.path.exists(BACKUP_FOLDER):
            os.makedirs(BACKUP_FOLDER, exist_ok=True)
    except Exception as e:
        st.error(f"Error creating directories: {e}")

create_directories()

# Hash password function
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

# Load users from JSON file
def load_users():
    """Load users with proper error handling"""
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
                    "password_last_changed": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "school_id": "admin123"
                }
            }
            save_users(default_users)
            return default_users
    except json.JSONDecodeError:
        st.error("Users file is corrupted. Creating new one...")
        default_users = {
            "admin": {
                "password": hash_password("admin123"),
                "name": "Administrator",
                "designation": "Admin",
                "role": "admin",
                "first_login": True,
                "password_last_changed": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "school_id": "admin123"
            }
        }
        save_users(default_users)
        return default_users
    except Exception as e:
        st.error(f"Error loading users: {e}")
        return {"admin": {"password": hash_password("admin123"), "name": "Admin", "designation": "Admin", "role": "admin", "first_login": True, "school_id": "admin123"}}

# Save users to JSON file
def save_users(users):
    """Save users with error handling"""
    try:
        with open(USER_DB_FILE, 'w', encoding='utf-8') as f:
            json.dump(users, f, indent=4)
        return True
    except Exception as e:
        st.error(f"Error saving users: {e}")
        return False

# Auto-create users from teachers in timetable
def auto_create_teacher_users():
    """Automatically create user accounts for teachers from timetable"""
    df = load_timetable()
    users = load_users()
    
    if df is not None and not df.empty:
        teachers = df['Teacher'].unique()
        created_count = 0
        
        for teacher in teachers:
            teacher_data = df[df['Teacher'] == teacher].iloc[0] if len(df[df['Teacher'] == teacher]) > 0 else None
            designation = teacher_data['Designation'] if teacher_data is not None else "Teacher"
            
            username = teacher.lower().replace(" ", "_").replace(".", "").replace("dr_", "dr").replace("prof_", "prof")
            
            if username not in users and teacher != "New Teacher":
                default_password = teacher.lower().replace(" ", "_")
                users[username] = {
                    "password": hash_password(default_password),
                    "name": teacher,
                    "designation": designation,
                    "role": "user",
                    "first_login": True,
                    "password_last_changed": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "teacher_name": teacher,
                    "school_id": default_password
                }
                created_count += 1
        
        if created_count > 0:
            save_users(users)
            return created_count
    return 0

# Load classrooms
def load_classrooms():
    """Load classroom data"""
    try:
        if os.path.exists(CLASSROOM_FILE):
            with open(CLASSROOM_FILE, 'r', encoding='utf-8') as f:
                content = f.read()
                if content.strip():
                    return json.loads(content)
                else:
                    return {}
        else:
            save_classrooms({})
            return {}
    except json.JSONDecodeError:
        save_classrooms({})
        return {}
    except Exception as e:
        st.error(f"Error loading classrooms: {e}")
        return {}

def save_classrooms(classrooms):
    """Save classroom data"""
    try:
        if classrooms is None:
            classrooms = {}
        with open(CLASSROOM_FILE, 'w', encoding='utf-8') as f:
            json.dump(classrooms, f, indent=4)
        return True
    except Exception as e:
        st.error(f"Error saving classrooms: {e}")
        return False

# Load arrangements
def load_arrangements():
    """Load arrangements with proper error handling"""
    try:
        if os.path.exists(ARRANGEMENT_FILE):
            with open(ARRANGEMENT_FILE, 'r', encoding='utf-8') as f:
                content = f.read()
                if content.strip():
                    return json.loads(content)
                else:
                    return {}
        else:
            save_arrangements({})
            return {}
    except json.JSONDecodeError:
        st.warning("Arrangements file was corrupted. Creating new one...")
        save_arrangements({})
        return {}
    except Exception as e:
        st.error(f"Error loading arrangements: {e}")
        return {}

# Save arrangements
def save_arrangements(arrangements):
    """Save arrangements with error handling"""
    try:
        if arrangements is None:
            arrangements = {}
        with open(ARRANGEMENT_FILE, 'w', encoding='utf-8') as f:
            json.dump(arrangements, f, indent=4)
        return True
    except Exception as e:
        st.error(f"Error saving arrangements: {e}")
        return False

# Load timetable
@st.cache_data(ttl=300)
def load_timetable():
    """Load timetable with multiple fallback methods"""
    
    if not OPENPYXL_AVAILABLE:
        st.error("openpyxl package is not installed!")
        return create_sample_timetable()
    
    try:
        if not os.path.exists(TIMETABLE_FILE):
            st.warning("No timetable file found. Creating sample data...")
            return create_sample_timetable()
        
        try:
            df = pd.read_excel(TIMETABLE_FILE, engine='openpyxl')
            if not df.empty:
                st.session_state.timetable_df = df
                return df
        except Exception as e:
            st.warning(f"Could not read with openpyxl: {e}")
            
        try:
            df = pd.read_excel(TIMETABLE_FILE)
            if not df.empty:
                st.session_state.timetable_df = df
                return df
        except Exception as e:
            st.warning(f"Could not read with default engine: {e}")
            
        return create_sample_timetable()
            
    except Exception as e:
        st.error(f"Error loading timetable: {e}")
        if st.session_state.timetable_df is not None:
            return st.session_state.timetable_df
        return create_sample_timetable()

def create_sample_timetable():
    """Create sample timetable data"""
    sample_data = {
        'Day': ['Monday', 'Monday', 'Monday', 'Monday', 'Monday', 'Monday', 'Monday', 'Monday',
                'Tuesday', 'Tuesday', 'Tuesday', 'Tuesday', 'Tuesday', 'Tuesday', 'Tuesday', 'Tuesday'],
        'Time': ['9:00-10:00', '10:00-11:00', '11:00-12:00', '12:00-1:00', '1:00-2:00', '2:00-3:00', '3:00-4:00', '4:00-5:00'] * 2,
        'Teacher': ['Dr. Smith', 'Prof. Johnson', 'Dr. Smith', 'Prof. Brown', 'Dr. Smith', 'Prof. Johnson', 'Dr. Smith', 'Prof. Brown'] * 2,
        'Subject': ['Mathematics', 'Physics', 'Mathematics', 'Chemistry', 'Mathematics', 'Physics', 'Mathematics', 'Chemistry'] * 2,
        'Class': ['10A', '10A', '10B', '10B', '10C', '10C', '10A', '10A'] * 2,
        'Designation': ['Math Teacher', 'Physics Teacher', 'Math Teacher', 'Chemistry Teacher', 
                       'Math Teacher', 'Physics Teacher', 'Math Teacher', 'Chemistry Teacher'] * 2,
        'Room': ['101', '102', '101', '103', '104', '102', '101', '103'] * 2
    }
    df = pd.DataFrame(sample_data)
    save_timetable(df)
    return df

def save_timetable(df):
    """Save timetable with simplified approach"""
    
    if not OPENPYXL_AVAILABLE:
        st.error("Cannot save: openpyxl not installed")
        return False
    
    try:
        gc.collect()
        time.sleep(0.5)
        
        try:
            df.to_excel(TIMETABLE_FILE, index=False, engine='openpyxl')
            st.session_state.timetable_df = df
            st.cache_data.clear()
            auto_create_teacher_users()
            return True
        except Exception as e1:
            st.warning(f"Direct save failed: {e1}")
            
            try:
                temp_file = tempfile.NamedTemporaryFile(
                    delete=False, 
                    suffix='.xlsx',
                    mode='wb'
                )
                temp_file.close()
                
                df.to_excel(temp_file.name, index=False, engine='openpyxl')
                
                if os.path.exists(TIMETABLE_FILE):
                    os.remove(TIMETABLE_FILE)
                shutil.copy2(temp_file.name, TIMETABLE_FILE)
                os.unlink(temp_file.name)
                
                st.session_state.timetable_df = df
                st.cache_data.clear()
                auto_create_teacher_users()
                return True
            except Exception as e3:
                st.error(f"All save methods failed. Last error: {e3}")
                return False
                    
    except Exception as e:
        st.error(f"Unexpected error saving: {e}")
        return False

# ============ DETECT COLLISIONS ============

def detect_collisions(df):
    """Detect teacher collisions (same teacher at same time on same day)"""
    collisions = []
    
    if df is None or df.empty:
        return collisions
    
    grouped = df.groupby(['Day', 'Time', 'Teacher']).size().reset_index(name='count')
    collisions_df = grouped[grouped['count'] > 1]
    
    if not collisions_df.empty:
        for _, row in collisions_df.iterrows():
            classes = df[(df['Day'] == row['Day']) & 
                        (df['Time'] == row['Time']) & 
                        (df['Teacher'] == row['Teacher'])]['Class'].tolist()
            collisions.append({
                'day': row['Day'],
                'time': row['Time'],
                'teacher': row['Teacher'],
                'classes': classes,
                'count': row['count']
            })
    
    return collisions

def highlight_collisions(df):
    """Create a styled dataframe with collision highlights"""
    if df is None or df.empty:
        return df
    
    collisions = detect_collisions(df)
    
    if not collisions:
        return df
    
    styled_df = df.copy()
    styled_df['_collision'] = False
    
    for collision in collisions:
        mask = (styled_df['Day'] == collision['day']) & \
               (styled_df['Time'] == collision['time']) & \
               (styled_df['Teacher'] == collision['teacher'])
        styled_df.loc[mask, '_collision'] = True
    
    return styled_df

# ============ ARRANGEMENT SUMMARY TABLE ============

def display_arrangement_summary():
    """Display arrangement summary in tabular format"""
    st.subheader("📊 Daily Arrangement Summary")
    
    df = load_timetable()
    if df.empty:
        st.warning("No timetable data available")
        return
    
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
    selected_day = st.selectbox("Select Day", days, key="summary_day")
    
    arrangements = load_arrangements()
    day_arrangements = {k: v for k, v in arrangements.items() if v.get('day') == selected_day and v.get('status') != 'completed'}
    
    if not day_arrangements:
        st.info(f"No active arrangements for {selected_day}")
        return
    
    day_times = df[df['Day'] == selected_day]['Time'].unique()
    day_times = sorted(day_times)
    
    period_labels = []
    for i in range(len(day_times)):
        if i == 0:
            period_labels.append(f"{i+1}st Period")
        elif i == 1:
            period_labels.append(f"{i+1}nd Period")
        elif i == 2:
            period_labels.append(f"{i+1}rd Period")
        else:
            period_labels.append(f"{i+1}th Period")
    
    absent_teachers_data = {}
    for key, value in day_arrangements.items():
        absent = value.get('absent_teacher')
        time_slot = value.get('time')
        replacement = value.get('replacement_teacher')
        class_name = value.get('class')
        
        if absent not in absent_teachers_data:
            absent_teachers_data[absent] = {}
        
        absent_teachers_data[absent][time_slot] = {
            'replacement': replacement,
            'class': class_name
        }
    
    table_data = []
    for absent, slots in absent_teachers_data.items():
        row = {'Absent Teacher': absent}
        for i, time_slot in enumerate(day_times):
            if time_slot in slots:
                row[period_labels[i]] = f"{slots[time_slot]['replacement']}\n({slots[time_slot]['class']})"
            else:
                row[period_labels[i]] = "—"
        table_data.append(row)
    
    if table_data:
        summary_df = pd.DataFrame(table_data)
        st.dataframe(summary_df, use_container_width=True, height=400)
        
        csv = summary_df.to_csv(index=False)
        st.download_button(
            label="📥 Download Summary as CSV",
            data=csv,
            file_name=f"arrangement_summary_{selected_day}.csv",
            mime="text/csv"
        )
    else:
        st.info(f"No arrangement data for {selected_day}")

# ============ PREDICTION SYSTEM ============

def get_teacher_availability(df, day, time_slot, exclude_teacher=None):
    """Get available teachers for a given time slot"""
    busy_teachers = df[(df['Day'] == day) & (df['Time'] == time_slot)]['Teacher'].tolist()
    all_teachers = df['Teacher'].unique()
    
    available = []
    for teacher in all_teachers:
        if teacher not in busy_teachers and teacher != exclude_teacher:
            teacher_duties = df[(df['Day'] == day) & (df['Time'] == time_slot) & (df['Teacher'] == teacher)]
            if teacher_duties.empty:
                available.append(teacher)
    
    return available

def calculate_teacher_load(df, teacher):
    """Calculate weekly load for a teacher"""
    return len(df[df['Teacher'] == teacher])

def get_teacher_vacant_periods(df, teacher, day):
    """Get all vacant periods for a teacher on a specific day"""
    teacher_schedule = df[(df['Teacher'] == teacher) & (df['Day'] == day)]
    teacher_times = set(teacher_schedule['Time'].tolist())
    all_times = set(df[df['Day'] == day]['Time'].unique())
    
    vacant_periods = all_times - teacher_times
    return list(vacant_periods)

def get_teacher_periods(df, teacher, day):
    """Get all periods for a teacher on a specific day"""
    teacher_schedule = df[(df['Teacher'] == teacher) & (df['Day'] == day)]
    return sorted(teacher_schedule['Time'].tolist())

def predict_best_replacement(df, absent_teacher, day, time_slot, class_name, subject):
    """PREDICTION ALGORITHM: Find best replacement teacher based on multiple criteria"""
    
    subject_teachers = df[df['Subject'] == subject]['Teacher'].unique()
    busy_teachers = df[(df['Day'] == day) & (df['Time'] == time_slot)]['Teacher'].tolist()
    on_duty = df[(df['Day'] == day) & (df['Time'] == time_slot) & (df['Designation'].str.contains('Duty', case=False, na=False))]['Teacher'].tolist()
    
    candidates = []
    
    for teacher in subject_teachers:
        if teacher == absent_teacher:
            continue
        if teacher in busy_teachers:
            continue
        if teacher in on_duty:
            continue
        
        score = 0
        load = calculate_teacher_load(df, teacher)
        score += (100 - load)
        
        vacant_periods = get_teacher_vacant_periods(df, teacher, day)
        if len(vacant_periods) >= 2:
            score += 30
        
        if len(df[(df['Teacher'] == teacher) & (df['Class'] == class_name)]) > 0:
            score += 20
        
        candidates.append((teacher, score))
    
    candidates.sort(key=lambda x: x[1], reverse=True)
    
    if candidates:
        return candidates[0][0]
    else:
        all_available = get_teacher_availability(df, day, time_slot, absent_teacher)
        return all_available[0] if all_available else None

def check_crisis_mode(df, arrangements):
    """Check if absent teacher count exceeds 40%"""
    total_teachers = len(df['Teacher'].unique())
    
    recent_absences = set()
    for key, value in arrangements.items():
        if 'status' in value and value['status'] == 'pending':
            recent_absences.add(value.get('absent_teacher'))
    
    absent_count = len(recent_absences)
    
    if total_teachers > 0 and (absent_count / total_teachers) >= 0.4:
        return True, absent_count, total_teachers
    return False, absent_count, total_teachers

# ============ ONLINE TIMETABLE EDITOR ============

def edit_timetable_online():
    """Inline editor for timetable with collision detection"""
    st.subheader("✏️ Online Timetable Editor")
    
    df = load_timetable()
    
    if df.empty:
        st.warning("No timetable data available. Please upload or create a new one.")
        return
    
    col1, col2, col3, col4 = st.columns([2, 1, 1, 1])
    with col1:
        if st.button("✏️ Enable Edit Mode", type="primary" if not st.session_state.editing_mode else "secondary"):
            st.session_state.editing_mode = True
            st.session_state.edit_df = df.copy()
            st.rerun()
    
    with col2:
        if st.button("📥 Export Excel"):
            if save_timetable(df):
                st.success("Timetable exported successfully!")
    
    with col3:
        if st.button("🔄 Refresh"):
            st.cache_data.clear()
            st.rerun()
    
    with col4:
        collisions = detect_collisions(df)
        if collisions:
            st.warning(f"⚠️ {len(collisions)} collision(s) detected!")
        else:
            st.success("✅ No collisions detected")
    
    if not st.session_state.editing_mode:
        collisions = detect_collisions(df)
        if collisions:
            st.error("🚨 **Teacher Collisions Detected!**")
            for coll in collisions:
                st.warning(f"👤 {coll['teacher']} is assigned to multiple classes on {coll['day']} at {coll['time']}: {', '.join(coll['classes'])}")
    
    if st.session_state.editing_mode:
        st.info("📝 Edit Mode Active - Click on any cell to edit, then click 'Save Changes'")
        
        if st.session_state.edit_df is not None:
            collisions = detect_collisions(st.session_state.edit_df)
            if collisions:
                st.error(f"⚠️ {len(collisions)} collision(s) in current edit data!")
                for coll in collisions:
                    st.warning(f"👤 {coll['teacher']} has {coll['count']} classes on {coll['day']} at {coll['time']}: {', '.join(coll['classes'])}")
        
        edited_df = st.data_editor(
            st.session_state.edit_df,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "Day": st.column_config.SelectboxColumn(
                    "Day",
                    options=["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
                    required=True
                ),
                "Time": st.column_config.TextColumn("Time", required=True),
                "Teacher": st.column_config.TextColumn("Teacher", required=True),
                "Subject": st.column_config.TextColumn("Subject", required=True),
                "Class": st.column_config.TextColumn("Class", required=True),
                "Designation": st.column_config.TextColumn("Designation", required=True),
                "Room": st.column_config.TextColumn("Room", required=False)
            }
        )
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            if st.button("💾 Save Changes", type="primary"):
                collisions = detect_collisions(edited_df)
                if collisions:
                    st.error("❌ Cannot save: Teacher collisions detected!")
                    for coll in collisions:
                        st.warning(f"👤 {coll['teacher']} has {coll['count']} classes on {coll['day']} at {coll['time']}")
                    if st.button("Force Save Anyway"):
                        if save_timetable(edited_df):
                            st.success("Timetable saved successfully!")
                            st.session_state.editing_mode = False
                            st.session_state.edit_df = None
                            st.cache_data.clear()
                            time.sleep(1)
                            st.rerun()
                else:
                    if save_timetable(edited_df):
                        st.success("Timetable saved successfully!")
                        st.session_state.editing_mode = False
                        st.session_state.edit_df = None
                        st.cache_data.clear()
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error("Failed to save timetable")
        
        with col2:
            if st.button("❌ Cancel Editing"):
                st.session_state.editing_mode = False
                st.session_state.edit_df = None
                st.rerun()
        
        with col3:
            if st.button("➕ Add New Row"):
                new_row = pd.DataFrame([{
                    'Day': 'Monday',
                    'Time': '11:00-12:00',
                    'Teacher': 'New Teacher',
                    'Subject': 'New Subject',
                    'Class': 'New Class',
                    'Designation': 'New Designation',
                    'Room': ''
                }])
                st.session_state.edit_df = pd.concat([st.session_state.edit_df, new_row], ignore_index=True)
                st.rerun()
        
        with col4:
            if st.button("🗑️ Delete Last Row"):
                if len(st.session_state.edit_df) > 0:
                    st.session_state.edit_df = st.session_state.edit_df.iloc[:-1]
                    st.rerun()
    
    else:
        st.subheader("📊 Current Timetable")
        
        if 'Room' in df.columns:
            display_cols = ['Day', 'Time', 'Teacher', 'Subject', 'Class', 'Designation', 'Room']
        else:
            display_cols = ['Day', 'Time', 'Teacher', 'Subject', 'Class', 'Designation']
        
        st.dataframe(df[display_cols], use_container_width=True)
        
        st.subheader("📊 Timetable Summary")
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("Total Classes", len(df))
        with col2:
            st.metric("Unique Teachers", len(df['Teacher'].unique()))
        with col3:
            st.metric("Unique Subjects", len(df['Subject'].unique()))
        with col4:
            st.metric("Unique Classes", len(df['Class'].unique()))
        with col5:
            collisions = detect_collisions(df)
            st.metric("Collisions", len(collisions), delta="⚠️" if collisions else "✅")

# ============ CLASSROOM MANAGEMENT UI ============

def classroom_management():
    """Complete classroom management interface"""
    st.subheader("🏫 Classroom Management")
    
    classrooms = load_classrooms()
    
    tab1, tab2, tab3 = st.tabs(["📋 View Classrooms", "➕ Add/Edit Classroom", "🗑️ Delete Classroom"])
    
    with tab1:
        if classrooms:
            for room_id, room_data in classrooms.items():
                with st.container():
                    col1, col2, col3 = st.columns([2, 2, 1])
                    with col1:
                        st.markdown(f"**🏠 {room_data.get('name', room_id)}**")
                    with col2:
                        st.markdown(f"Capacity: {room_data.get('capacity', 'N/A')} students")
                    with col3:
                        st.markdown(f"Floor: {room_data.get('floor', 'N/A')}")
                    
                    if room_data.get('equipment'):
                        st.caption(f"🛠️ Equipment: {', '.join(room_data['equipment'])}")
                    
                    if room_data.get('current_class'):
                        st.info(f"📚 Current Class: {room_data['current_class']}")
                    
                    st.markdown("---")
        else:
            st.info("No classrooms added yet. Use 'Add Classroom' tab to add.")
    
    with tab2:
        st.subheader("Add/Edit Classroom")
        
        operation = st.radio("Select Operation", ["Add New Classroom", "Edit Existing Classroom"])
        
        if operation == "Edit Existing Classroom" and classrooms:
            selected_room = st.selectbox("Select Classroom to Edit", list(classrooms.keys()))
            room_data = classrooms.get(selected_room, {})
        else:
            selected_room = None
            room_data = {}
        
        with st.form("classroom_form"):
            room_name = st.text_input("Classroom Name/Room Number", value=room_data.get('name', '') if room_data else '')
            capacity = st.number_input("Capacity (Number of Students)", min_value=1, value=room_data.get('capacity', 30) if room_data else 30)
            floor = st.number_input("Floor Level", min_value=0, max_value=10, value=room_data.get('floor', 1) if room_data else 1)
            
            equipment = st.multiselect(
                "Equipment Available",
                options=["Projector", "Smart Board", "AC", "Computers", "WiFi", "Whiteboard", "Speakers", "Microphone"],
                default=room_data.get('equipment', []) if room_data else []
            )
            
            current_class = st.text_input("Currently Assigned Class (Optional)", value=room_data.get('current_class', '') if room_data else '')
            
            submitted = st.form_submit_button("Save Classroom")
            
            if submitted:
                if room_name:
                    room_id = room_name.replace(" ", "_").lower()
                    classrooms[room_id] = {
                        "name": room_name,
                        "capacity": capacity,
                        "floor": floor,
                        "equipment": equipment,
                        "current_class": current_class,
                        "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    save_classrooms(classrooms)
                    st.success(f"Classroom '{room_name}' saved successfully!")
                    st.rerun()
                else:
                    st.error("Please enter classroom name")
    
    with tab3:
        if classrooms:
            st.subheader("Delete Classroom")
            room_to_delete = st.selectbox("Select Classroom to Delete", list(classrooms.keys()))
            
            if st.button("🗑️ Delete Classroom", type="secondary"):
                confirm = st.checkbox("I confirm deletion of this classroom")
                if confirm:
                    del classrooms[room_to_delete]
                    save_classrooms(classrooms)
                    st.success(f"Classroom deleted successfully!")
                    st.rerun()
        else:
            st.info("No classrooms to delete")

# ============ IMPROVED ARRANGEMENT SYSTEM ============

def arrangement_management():
    """Enhanced arrangement management with full day and half day options"""
    st.subheader("📋 Teacher Absence & Intelligent Arrangement System")
    
    df = load_timetable()
    if df.empty:
        st.warning("Please upload timetable first")
        return
    
    arrangements = load_arrangements()
    if arrangements is None:
        arrangements = {}
    with tab3:
        st.subheader("Upload Timetable (Excel)")
    uploaded_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            
            # Clean data immediately after upload
            required_cols = ['Day', 'Time', 'Teacher', 'Subject', 'Class', 'Designation']
            
            # Check required columns
            if not all(col in df.columns for col in required_cols):
                st.error(f"Missing columns. Required: {required_cols}")
                st.write("Your columns:", df.columns.tolist())
            else:
                # Convert all columns to string and clean
                for col in df.columns:
                    df[col] = df[col].fillna('').astype(str).str.strip()
                
                # Add Room column if missing
                if 'Room' not in df.columns:
                    df['Room'] = ''
                
                # Remove empty rows
                df = df[df['Teacher'] != '']
                df = df[df['Teacher'] != 'nan']
                
                if save_timetable(df):
                    st.success("Timetable uploaded successfully!")
                    st.rerun()
                else:
                    st.error("Failed to save timetable")
        except Exception as e:
            st.error(f"Error reading file: {e}")
    display_arrangement_summary()
    
    st.markdown("---")
    
    crisis_mode, absent_count, total_teachers = check_crisis_mode(df, arrangements)
    
    if crisis_mode:
        st.error(f"⚠️ **CRISIS MODE ACTIVATED!** {absent_count}/{total_teachers} teachers absent ({int(absent_count/total_teachers*100)}%)")
        st.warning("Using Serial Class Prediction System - Teachers will be assigned 3-4 consecutive classes")
    
    days = df['Day'].unique() if not df.empty else []
    time_periods = df['Time'].unique() if not df.empty else []
    teachers = df['Teacher'].unique() if not df.empty else []
    
    st.subheader("1️⃣ Report Teacher Absence")
    with st.form("absence_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            absent_teacher = st.selectbox("Absent Teacher", teachers.tolist())
        with col2:
            absence_day = st.selectbox("Day of Absence", days.tolist())
        with col3:
            absence_type = st.selectbox("Absence Type", ["Single Period", "Full Day", "First Half", "Second Half"])
        
        if absence_type == "Single Period":
            absence_time = st.selectbox("Time Period", time_periods.tolist())
            periods_to_cover = [absence_time]
        elif absence_type == "Full Day":
            teacher_periods = get_teacher_periods(df, absent_teacher, absence_day)
            periods_to_cover = teacher_periods
            st.info(f"📚 This teacher has {len(periods_to_cover)} periods on this day")
            if periods_to_cover:
                st.write("Periods to cover:", ", ".join(periods_to_cover))
            absence_time = "Full Day"
        elif absence_type == "First Half":
            all_periods = sorted(df[df['Day'] == absence_day]['Time'].unique())
            mid_point = len(all_periods) // 2
            periods_to_cover = all_periods[:mid_point]
            st.info(f"📚 First half has {len(periods_to_cover)} periods")
            if periods_to_cover:
                st.write("Periods to cover:", ", ".join(periods_to_cover))
            absence_time = "First Half"
        else:
            all_periods = sorted(df[df['Day'] == absence_day]['Time'].unique())
            mid_point = len(all_periods) // 2
            periods_to_cover = all_periods[mid_point:]
            st.info(f"📚 Second half has {len(periods_to_cover)} periods")
            if periods_to_cover:
                st.write("Periods to cover:", ", ".join(periods_to_cover))
            absence_time = "Second Half"
        
        reason = st.text_area("Reason for Absence")
        
        if st.form_submit_button("Report Absence & Get Predictions"):
            if periods_to_cover:
                st.success(f"✅ Processing {len(periods_to_cover)} periods for {absent_teacher}")
                
                for period in periods_to_cover:
                    absent_class = df[(df['Day'] == absence_day) & 
                                     (df['Time'] == period) & 
                                     (df['Teacher'] == absent_teacher)]
                    
                    if not absent_class.empty:
                        subject = absent_class.iloc[0]['Subject']
                        class_name = absent_class.iloc[0]['Class']
                        
                        best_replacement = predict_best_replacement(df, absent_teacher, absence_day, period, class_name, subject)
                        
                        if best_replacement:
                            key = f"{absence_day}_{period}_{class_name}"
                            arrangements[key] = {
                                "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "absent_teacher": absent_teacher,
                                "replacement_teacher": best_replacement,
                                "class": class_name,
                                "subject": subject,
                                "day": absence_day,
                                "time": period,
                                "reason": reason,
                                "status": "assigned",
                                "crisis_mode": crisis_mode,
                                "absence_type": absence_type
                            }
                            save_arrangements(arrangements)
                        else:
                            st.warning(f"⚠️ No suitable replacement found for {period}")
                
                st.success("All arrangements processed successfully!")
                st.rerun()
            else:
                st.error("No periods to cover for this teacher on this day")
    
    st.subheader("2️⃣ Individual Arrangements")
    
    st.subheader("3️⃣ Teachers Needing Arrangements")
    all_teachers = set(df['Teacher'].unique())
    assigned_teachers = set()
    for key, value in arrangements.items():
        if value.get('status') != 'completed':
            assigned_teachers.add(value.get('absent_teacher'))
    
    remaining_teachers = all_teachers - assigned_teachers
    if remaining_teachers:
        st.warning(f"⚠️ {len(remaining_teachers)} teachers still need arrangements:")
        for teacher in remaining_teachers:
            st.write(f"- {teacher}")
    else:
        st.success("✅ All teachers have arrangements assigned")
    
    st.markdown("---")
    
    if arrangements and len(arrangements) > 0:
        for key, value in list(arrangements.items()):
            if value.get('status') != 'completed':
                with st.expander(f"📅 {value.get('day', 'N/A')} - {value.get('time', 'N/A')} - {value.get('class', 'N/A')}"):
                    st.write(f"**Absent Teacher:** {value.get('absent_teacher', 'N/A')}")
                    st.write(f"**Replacement:** {value.get('replacement_teacher', 'N/A')}")
                    st.write(f"**Subject:** {value.get('subject', 'N/A')}")
                    st.write(f"**Status:** {value.get('status', 'N/A')}")
                    if value.get('absence_type'):
                        st.write(f"**Absence Type:** {value.get('absence_type')}")
                    if value.get('crisis_mode'):
                        st.warning("⚠️ Crisis Mode Assignment")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button(f"✓ Mark Complete", key=f"complete_{key}"):
                            value['status'] = 'completed'
                            save_arrangements(arrangements)
                            st.rerun()
                    with col2:
                        if st.button(f"🗑️ Delete", key=f"del_{key}"):
                            del arrangements[key]
                            save_arrangements(arrangements)
                            st.rerun()
    else:
        st.info("No pending arrangements")

# ============ AUTHENTICATION FUNCTIONS ============

def change_password(username, old_password, new_password, confirm_password):
    """Change user password with validation"""
    users = load_users()
    
    if username not in users:
        return False, "User not found"
    
    if users[username]['password'] != hash_password(old_password):
        return False, "Current password is incorrect"
    
    if old_password == new_password:
        return False, "New password cannot be the same as current password"
    
    if len(new_password) < 6:
        return False, "New password must be at least 6 characters long"
    
    if new_password != confirm_password:
        return False, "New passwords do not match"
    
    users[username]['password'] = hash_password(new_password)
    users[username]['first_login'] = False
    users[username]['password_last_changed'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    save_users(users)
    return True, "Password changed successfully!"

def reset_user_password(username, new_password):
    """Admin function to reset user password"""
    users = load_users()
    
    if username not in users:
        return False, "User not found"
    
    if len(new_password) < 6:
        return False, "Password must be at least 6 characters long"
    
    users[username]['password'] = hash_password(new_password)
    users[username]['first_login'] = False
    users[username]['password_last_changed'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    save_users(users)
    return True, f"Password reset for {username} successfully!"

def password_change_form():
    """Display password change form"""
    st.markdown("---")
    st.subheader("🔐 Change Password")
    st.warning("⚠️ For security reasons, please change your default password")
    
    with st.form("change_password_form"):
        old_password = st.text_input("Current Password", type="password")
        new_password = st.text_input("New Password", type="password", 
                                     help="Password must be at least 6 characters long")
        confirm_password = st.text_input("Confirm New Password", type="password")
        
        col1, col2 = st.columns(2)
        with col1:
            submit = st.form_submit_button("Change Password", type="primary")
        with col2:
            skip = st.form_submit_button("Remind Me Later")
    
    if submit:
        if old_password and new_password and confirm_password:
            success, message = change_password(
                st.session_state.username, 
                old_password, 
                new_password, 
                confirm_password
            )
            if success:
                st.success(message)
                st.session_state.password_changed = True
                st.session_state.show_password_change = False
                st.balloons()
                time.sleep(1)
                st.rerun()
            else:
                st.error(message)
        else:
            st.error("Please fill all fields")
    
    if skip:
        st.session_state.show_password_change = False
        st.rerun()

def login(username, password, school_id):
    users = load_users()
    if username in users:
        if users[username]['password'] == hash_password(school_id) or users[username]['password'] == hash_password(password):
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.role = users[username]['role']
            st.session_state.name = users[username]['name']
            st.session_state.designation = users[username]['designation']
            st.session_state.school_id = school_id
            
            if users[username].get('first_login', False) and username == 'admin':
                st.session_state.show_password_change = True
            else:
                st.session_state.show_password_change = False
            
            return True
    return False

def logout():
    st.session_state.logged_in = False
    st.session_state.username = None
    st.session_state.role = None
    st.session_state.name = None
    st.session_state.designation = None
    st.session_state.show_password_change = False
    st.session_state.password_changed = False
    st.session_state.editing_mode = False
    st.session_state.school_id = None
    st.rerun()

def admin_panel():
    st.header("👑 Admin Panel")
    
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "Create User", "Manage Users", "Upload Timetable", 
        "Online Editor", "Classroom Management", "Arrangements"
    ])
    
    with tab1:
        st.subheader("Create New User")
        with st.form("create_user_form"):
            new_username = st.text_input("Username")
            new_school_id = st.text_input("School ID (will be used as default password)", value="teacher123")
            new_name = st.text_input("Full Name")
            new_designation = st.text_input("Designation")
            new_role = st.selectbox("Role", ["user", "admin"])
            
            st.info(f"💡 Default password for new users: {new_school_id}")
            
            if st.form_submit_button("Create User"):
                if new_username and new_school_id and new_name and new_designation:
                    if len(new_school_id) < 4:
                        st.error("School ID must be at least 4 characters long!")
                    else:
                        users = load_users()
                        if new_username not in users:
                            users[new_username] = {
                                "password": hash_password(new_school_id),
                                "name": new_name,
                                "designation": new_designation,
                                "role": new_role,
                                "first_login": True,
                                "password_last_changed": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "school_id": new_school_id
                            }
                            save_users(users)
                            st.success(f"User {new_username} created successfully! School ID: {new_school_id}")
                            st.rerun()
                        else:
                            st.error("Username already exists!")
                else:
                    st.error("Please fill all fields!")
    
    with tab2:
        st.subheader("Manage Users")
        users = load_users()
        
        st.info("💡 Teacher accounts are automatically created from timetable data with School ID as password")
        
        user_list = [u for u in users.keys() if u != 'admin']
        
        if user_list:
            for username in user_list:
                with st.expander(f"User: {username}"):
                    st.write(f"**Name:** {users[username]['name']}")
                    st.write(f"**Designation:** {users[username]['designation']}")
                    st.write(f"**Role:** {users[username]['role']}")
                    st.write(f"**School ID:** {users[username].get('school_id', 'N/A')}")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button(f"Delete {username}", key=f"del_{username}"):
                            del users[username]
                            save_users(users)
                            st.success(f"User {username} deleted!")
                            st.rerun()
                    with col2:
                        with st.popover(f"Reset Password for {username}"):
                            new_pass = st.text_input(f"New password for {username}", type="password", key=f"reset_pass_{username}", value=users[username].get('school_id', 'teacher123'))
                            if st.button(f"Confirm Reset", key=f"confirm_reset_{username}"):
                                if new_pass and len(new_pass) >= 4:
                                    success, message = reset_user_password(username, new_pass)
                                    if success:
                                        st.success(message)
                                        st.rerun()
                                    else:
                                        st.error(message)
                                else:
                                    st.error("Password must be at least 4 characters")
        else:
            st.info("No users found except admin")
    
    with tab3:
        st.subheader("Upload Timetable (Excel)")
        uploaded_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'])
        
        if uploaded_file:
            try:
                df = pd.read_excel(uploaded_file)
                required_cols = ['Day', 'Time', 'Teacher', 'Subject', 'Class', 'Designation']
                if all(col in df.columns for col in required_cols):
                    if 'Room' not in df.columns:
                        df['Room'] = ''
                    if save_timetable(df):
                        st.success("Timetable uploaded successfully! Teacher accounts created automatically.")
                        st.rerun()
                else:
                    st.error(f"Missing columns. Required: {required_cols}")
            except Exception as e:
                st.error(f"Error: {e}")
    
    with tab4:
        edit_timetable_online()
    
    with tab5:
        classroom_management()
    
    with tab6:
        arrangement_management()

def user_dashboard():
    st.header(f"👋 Welcome, {st.session_state.name}!")
    st.write(f"**Designation:** {st.session_state.designation}")
    st.write(f"**School ID:** {st.session_state.school_id}")
    
    df = load_timetable()
    
    if df.empty:
        st.warning("No timetable available. Please contact admin.")
        return
    
    st.subheader("📅 Your Timetable")
    
    user_timetable = df[df['Teacher'] == st.session_state.name]
    
    if user_timetable.empty:
        user_timetable = df[df['Designation'].str.lower() == st.session_state.designation.lower()]
    
    if not user_timetable.empty:
        display_cols = ['Day', 'Time', 'Subject', 'Class']
        if 'Room' in user_timetable.columns:
            display_cols.append('Room')
        display_df = user_timetable[display_cols].sort_values(['Day', 'Time'])
        st.dataframe(display_df, use_container_width=True)
        
        st.subheader("📊 Your Weekly Summary")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Classes", len(user_timetable))
        with col2:
            st.metric("Unique Subjects", len(user_timetable['Subject'].unique()))
        with col3:
            st.metric("Unique Classes", len(user_timetable['Class'].unique()))
        with col4:
            if 'Room' in user_timetable.columns:
                rooms = user_timetable['Room'].unique()
                st.metric("Rooms Assigned", len(rooms))
    else:
        st.info(f"No timetable entries found for {st.session_state.name}")
    
    st.subheader("🏫 Your Classroom Assignments")
    classrooms = load_classrooms()
    if classrooms and 'Room' in user_timetable.columns:
        assigned_rooms = user_timetable['Room'].unique()
        assigned_rooms = [r for r in assigned_rooms if r and r != '']
        
        if assigned_rooms:
            for room in assigned_rooms:
                if room in classrooms:
                    room_data = classrooms[room]
                    st.info(f"**Room {room}** - Capacity: {room_data.get('capacity', 'N/A')} | Floor: {room_data.get('floor', 'N/A')}")
                else:
                    st.info(f"**Room {room}**")
        else:
            st.info("No classrooms assigned")
    else:
        st.info("No classroom data available")
    
    st.subheader("🔄 Your Arrangement Assignments")
    arrangements = load_arrangements()
    
    if arrangements and len(arrangements) > 0:
        user_arrangements = []
        for key, value in arrangements.items():
            if value.get('replacement_teacher') == st.session_state.name:
                if value.get('status') != 'completed':
                    user_arrangements.append(value)
        
        if user_arrangements:
            for arr in user_arrangements:
                with st.expander(f"📌 {arr.get('day')} - {arr.get('time')}"):
                    st.write(f"**Class:** {arr.get('class')}")
                    st.write(f"**Subject:** {arr.get('subject')}")
                    st.write(f"**Covering for:** {arr.get('absent_teacher')}")
                    st.write(f"**Reason:** {arr.get('reason', 'Not specified')}")
                    if arr.get('absence_type'):
                        st.write(f"**Absence Type:** {arr.get('absence_type')}")
                    if arr.get('crisis_mode'):
                        st.warning("⚠️ Crisis Mode - Multiple classes may be assigned")
        else:
            st.info("No active arrangements assigned to you")
    else:
        st.info("No arrangements found")

def main():
    st.set_page_config(
        page_title="Timetable Management System with AI Prediction",
        page_icon="📚",
        layout="wide"
    )
    
    # Custom CSS
    st.markdown("""
    <style>
    .stApp {
        background-color: #f0f2f6;
    }
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
    }
    .login-container {
        background: white;
        padding: 3rem;
        border-radius: 15px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        max-width: 400px;
        margin: 0 auto;
    }
    .collision-warning {
        background-color: #ffcccc;
        padding: 10px;
        border-radius: 5px;
        border-left: 5px solid #ff0000;
    }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class="main-header">
        <h1>📚 Intelligent Timetable Management System</h1>
        <p>Powered by AI Prediction & Classroom Management</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Auto-create teacher users from timetable on startup
    created = auto_create_teacher_users()
    if created > 0 and st.session_state.get('logged_in'):
        st.success(f"✅ Auto-created {created} teacher accounts")
    
    if not st.session_state.logged_in:
        st.subheader("🔐 Login")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            with st.container():
                st.markdown('<div class="login-container">', unsafe_allow_html=True)
                
                with st.form("login_form"):
                    username = st.text_input("Username", placeholder="Enter your username")
                    school_id = st.text_input("School ID", placeholder="Enter your School ID", type="password")
                    password = st.text_input("Password (Optional)", placeholder="Enter password if different from School ID", type="password")
                    submit = st.form_submit_button("🔑 Login", use_container_width=True)
                    
                    if submit:
                        if login(username, password if password else school_id, school_id):
                            st.success(f"Welcome {st.session_state.name}!")
                            st.rerun()
                        else:
                            st.error("Invalid username, School ID, or password!")
                
                st.markdown("---")
                st.caption("💡 **Demo Credentials:**")
                st.caption("Admin: admin / School ID: admin123")
                st.caption("Teachers: Use teacher name as username (e.g., dr_smith) with School ID as password")
                st.caption("Default School ID for teachers: teacher name (e.g., dr_smith)")
                
                st.markdown('</div>', unsafe_allow_html=True)
    else:
        if st.session_state.show_password_change and not st.session_state.password_changed:
            password_change_form()
        else:
            with st.sidebar:
                st.write(f"**👤 Logged in as:** {st.session_state.name}")
                st.write(f"**Username:** {st.session_state.username}")
                st.write(f"**Role:** {st.session_state.role}")
                st.write(f"**School ID:** {st.session_state.school_id}")
                st.markdown("---")
                
                if st.button("🚪 Logout", use_container_width=True):
                    logout()
                
                st.markdown("---")
                st.caption(f"🕐 Login Time: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
            
            if st.session_state.role == 'admin':
                admin_panel()
                st.markdown("---")
                st.subheader("📊 Your Dashboard (Admin View)")
                user_dashboard()
            else:
                user_dashboard()

if __name__ == "__main__":
    main()
