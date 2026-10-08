"""
STUDENT MANAGEMENT SYSTEM
Backend Flask Application
Version: 4.0.0 (Improved & Optimized)
"""

from flask import (
    Flask, render_template, session, redirect, request, jsonify,
    url_for, flash, make_response, g, send_from_directory
)
import sqlite3
import os
import secrets
from functools import wraps
from datetime import datetime, date
from werkzeug.security import generate_password_hash, check_password_hash
import logging
from logging.handlers import RotatingFileHandler
from ai_attendance import analyse_student_course

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from io import BytesIO
from ai_features import predict_final_grade, calculate_dropout_risk, generate_study_recommendations
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from werkzeug.utils import secure_filename
from flask_socketio import SocketIO, emit, join_room
from ai_questions import get_questions_for_course
import json

from dotenv import load_dotenv
load_dotenv()

# ============================================================================
# APP INITIALIZATION
# ============================================================================

app = Flask(__name__, template_folder='templates', static_folder='static')
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', secrets.token_hex(32))
app.config['DATABASE'] = os.environ.get('DATABASE_URI', 'database.db')
app.config['ITEMS_PER_PAGE'] = 20
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max
ALLOWED_EXTENSIONS = {'pdf', 'doc', 'docx', 'txt', 'png', 'jpg', 'jpeg', 'zip', 'pptx', 'xlsx'}

# Create uploads folder if not exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(os.path.join(app.config['UPLOAD_FOLDER'], 'assignments'), exist_ok=True)
os.makedirs(os.path.join(app.config['UPLOAD_FOLDER'], 'submissions'), exist_ok=True)
app.config['DEBUG'] = True
socketio = SocketIO(app, cors_allowed_origins="*", async_mode='eventlet')

# Setup logging
if not os.path.exists('logs'):
    os.makedirs('logs')

file_handler = RotatingFileHandler('logs/app.log', maxBytes=10240, backupCount=10)
file_handler.setFormatter(logging.Formatter(
    '%(asctime)s %(levelname)s: %(message)s [in %(pathname)s:%(lineno)d]'
))
file_handler.setLevel(logging.INFO)
app.logger.addHandler(file_handler)
app.logger.setLevel(logging.INFO)
app.logger.info('Student Management System startup')

# ============================================================================
# DATABASE HELPERS
# ============================================================================

def get_db():
    """Get database connection"""
    if 'db' not in g:
        g.db = sqlite3.connect(app.config['DATABASE'])
        g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(error):
    """Close database connection"""
    db = g.pop('db', None)
    if db is not None:
        db.close()

def init_db():
    """Initialize database with all tables and default users"""
    conn = sqlite3.connect(app.config['DATABASE'])
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    
    # Create tables
    tables = [
        '''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL,
            role TEXT DEFAULT 'student',
            is_active BOOLEAN DEFAULT 1,
            last_login DATETIME,
            profile_image TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS departments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            dept_code TEXT UNIQUE NOT NULL,
            dept_name TEXT UNIQUE NOT NULL,
            description TEXT,
            head_of_department TEXT,
            office_location TEXT,
            phone TEXT,
            email TEXT,
            established_date DATE,
            budget DECIMAL(10,2),
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS instructors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER UNIQUE,
            instructor_id TEXT UNIQUE NOT NULL,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            phone TEXT,
            department_id INTEGER,
            qualification TEXT,
            specialization TEXT,
            joining_date DATE,
            salary DECIMAL(10,2),
            status TEXT DEFAULT 'active',
            office_hours TEXT,
            profile_image TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id),
            FOREIGN KEY (department_id) REFERENCES departments(id)
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER UNIQUE,
            student_id TEXT UNIQUE NOT NULL,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL,
            date_of_birth DATE,
            gender TEXT,
            email TEXT UNIQUE NOT NULL,
            phone TEXT,
            address TEXT,
            city TEXT,
            state TEXT,
            zip_code TEXT,
            country TEXT DEFAULT 'USA',
            department_id INTEGER,
            enrollment_date DATE DEFAULT CURRENT_DATE,
            graduation_date DATE,
            status TEXT DEFAULT 'active',
            guardian_name TEXT,
            guardian_phone TEXT,
            guardian_email TEXT,
            emergency_contact TEXT,
            medical_info TEXT,
            profile_image TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id),
            FOREIGN KEY (department_id) REFERENCES departments(id)
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS courses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            course_code TEXT UNIQUE NOT NULL,
            course_name TEXT NOT NULL,
            description TEXT,
            credits INTEGER DEFAULT 3,
            department_id INTEGER,
            instructor_id INTEGER,
            semester TEXT,
            academic_year TEXT,
            max_students INTEGER DEFAULT 30,
            enrolled_students INTEGER DEFAULT 0,
            status TEXT DEFAULT 'active',
            start_date DATE,
            end_date DATE,
            schedule TEXT,
            classroom TEXT,
            prerequisites TEXT,
            syllabus TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (department_id) REFERENCES departments(id),
            FOREIGN KEY (instructor_id) REFERENCES instructors(id)
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS enrollments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            course_id INTEGER NOT NULL,
            enrollment_date DATETIME DEFAULT CURRENT_TIMESTAMP,
            status TEXT DEFAULT 'enrolled',
            grade TEXT,
            progress INTEGER DEFAULT 0,
            last_accessed DATETIME,
            completed_date DATETIME,
            notes TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (student_id) REFERENCES students(id),
            FOREIGN KEY (course_id) REFERENCES courses(id),
            UNIQUE(student_id, course_id)
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            course_id INTEGER NOT NULL,
            date DATE NOT NULL,
            status TEXT CHECK(status IN ('present', 'absent', 'late', 'excused')),
            check_in_time TIME,
            check_out_time TIME,
            notes TEXT,
            marked_by INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (student_id) REFERENCES students(id),
            FOREIGN KEY (course_id) REFERENCES courses(id),
            FOREIGN KEY (marked_by) REFERENCES users(id),
            UNIQUE(student_id, course_id, date)
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS grades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            enrollment_id INTEGER NOT NULL,
            assessment_type TEXT NOT NULL,
            assessment_name TEXT NOT NULL,
            score DECIMAL(5,2),
            max_score DECIMAL(5,2) DEFAULT 100,
            weight DECIMAL(5,2) DEFAULT 1,
            comments TEXT,
            graded_by INTEGER,
            graded_date DATETIME,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (enrollment_id) REFERENCES enrollments(id),
            FOREIGN KEY (graded_by) REFERENCES users(id)
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            payment_date DATETIME DEFAULT CURRENT_TIMESTAMP,
            amount DECIMAL(10,2) NOT NULL,
            payment_type TEXT,
            payment_method TEXT,
            transaction_id TEXT UNIQUE,
            status TEXT DEFAULT 'completed',
            description TEXT,
            receipt_number TEXT UNIQUE,
            processed_by INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (student_id) REFERENCES students(id),
            FOREIGN KEY (processed_by) REFERENCES users(id)
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS notifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            message TEXT NOT NULL,
            type TEXT DEFAULT 'info',
            is_read BOOLEAN DEFAULT 0,
            link TEXT,
            action_url TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id)
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            action TEXT NOT NULL,
            table_name TEXT,
            record_id INTEGER,
            old_values TEXT,
            new_values TEXT,
            ip_address TEXT,
            user_agent TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id)
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS classroom_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            course_id INTEGER NOT NULL,
            instructor_id INTEGER NOT NULL,
            session_title TEXT,
            status TEXT DEFAULT 'waiting',
            started_at DATETIME,
            ended_at DATETIME,
            questions_data TEXT,
            notes TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (course_id) REFERENCES courses(id),
            FOREIGN KEY (instructor_id) REFERENCES instructors(id)
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS classroom_attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER NOT NULL,
            student_id INTEGER NOT NULL,
            joined_at DATETIME,
            left_at DATETIME,
            status TEXT DEFAULT 'present',
            was_kicked BOOLEAN DEFAULT 0,
            FOREIGN KEY (session_id) REFERENCES classroom_sessions(id),
            FOREIGN KEY (student_id) REFERENCES students(id),
            UNIQUE(session_id, student_id)
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS proctor_responses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER NOT NULL,
            student_id INTEGER NOT NULL,
            question_id INTEGER NOT NULL,
            question_text TEXT,
            answer_given INTEGER,
            correct_answer INTEGER,
            is_correct BOOLEAN,
            time_taken_seconds INTEGER,
            answered_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (session_id) REFERENCES classroom_sessions(id),
            FOREIGN KEY (student_id) REFERENCES students(id)
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS assignments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            course_id INTEGER NOT NULL,
            title TEXT NOT NULL,
            description TEXT,
            due_date DATE,
            max_marks INTEGER DEFAULT 100,
            file_path TEXT,
            created_by INTEGER,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (course_id) REFERENCES courses(id)
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            assignment_id INTEGER NOT NULL,
            student_id INTEGER NOT NULL,
            file_path TEXT,
            submitted_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            marks INTEGER,
            feedback TEXT,
            status TEXT DEFAULT 'submitted',
            FOREIGN KEY (assignment_id) REFERENCES assignments(id),
            FOREIGN KEY (student_id) REFERENCES students(id),
            UNIQUE(assignment_id, student_id)
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS email_settings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            smtp_email TEXT,
            smtp_password TEXT,
            smtp_server TEXT DEFAULT 'smtp.gmail.com',
            smtp_port INTEGER DEFAULT 587,
            enabled BOOLEAN DEFAULT 0
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS settings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            setting_key TEXT UNIQUE NOT NULL,
            setting_value TEXT,
            setting_type TEXT DEFAULT 'string',
            description TEXT,
            updated_by INTEGER,
            updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS course_applications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            course_id INTEGER NOT NULL,
            reason TEXT,
            status TEXT DEFAULT 'pending',
            reviewed_by INTEGER,
            reviewed_at DATETIME,
            admin_note TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (student_id) REFERENCES students(id),
            FOREIGN KEY (course_id) REFERENCES courses(id),
            FOREIGN KEY (reviewed_by) REFERENCES users(id),
            UNIQUE(student_id, course_id)
        )
        ''',
        '''
        CREATE TABLE IF NOT EXISTS student_badges (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            badge_name TEXT NOT NULL,
            icon TEXT,
            description TEXT,
            earned_at DATETIME DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (student_id) REFERENCES students(id)
        )
        '''
    ]
    
    for table_sql in tables:
        c.execute(table_sql)
    
    # Insert default admin user
    c.execute("SELECT * FROM users WHERE username = 'admin'")
    if not c.fetchone():
        admin_password = generate_password_hash('admin123')
        c.execute(
            """INSERT INTO users (username, email, password_hash, first_name, last_name, role) VALUES (?, ?, ?, ?, ?, ?)""",
            ('admin', 'admin@system.com', admin_password, 'System', 'Administrator', 'admin'))
        print("[OK] Admin user created")
    
    # Insert default student user for testing
    c.execute("SELECT * FROM users WHERE username = 'student'")
    if not c.fetchone():
        try:
            student_password = generate_password_hash('student123')
            c.execute(
                """INSERT INTO users (username, email, password_hash, first_name, last_name, role) VALUES (?, ?, ?, ?, ?, ?)""",
                ('student', 'student@school.edu', student_password, 'John', 'Student', 'student'))
            print("[OK] Default student user created")
            student_user_id = c.lastrowid
            c.execute(
                """INSERT OR IGNORE INTO students (user_id, student_id, first_name, last_name, email, enrollment_date, status) VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (student_user_id, 'STU001', 'John', 'Student', 'student@school.edu', datetime.now().strftime('%Y-%m-%d'), 'active'))
            print("[OK] Student record created")
        except Exception as e:
            print(f"[WARN] Default student setup skipped: {e}")

    # Insert default faculty user for testing
    c.execute("SELECT * FROM users WHERE username = 'faculty'")
    if not c.fetchone():
        try:
            faculty_password = generate_password_hash('faculty123')
            c.execute(
                """INSERT INTO users (username, email, password_hash, first_name, last_name, role) VALUES (?, ?, ?, ?, ?, ?)""",
                ('faculty', 'faculty@school.edu', faculty_password, 'Jane', 'Teacher', 'teacher'))
            print("[OK] Default faculty user created")
            faculty_user_id = c.lastrowid
            c.execute(
                """INSERT OR IGNORE INTO instructors (user_id, instructor_id, first_name, last_name, email, joining_date, status) VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (faculty_user_id, 'FAC001', 'Jane', 'Teacher', 'faculty@school.edu', datetime.now().strftime('%Y-%m-%d'), 'active'))
            print("[OK] Faculty record created")
        except Exception as e:
            print(f"[WARN] Default faculty setup skipped: {e}")
    
    # Insert default departments
    departments_data = [
        ('CS', 'Computer Science', 'Department of Computer Science', 'Dr. Smith', 'Building A, Room 101', '+1-555-0101', 'cs@school.edu', '2000-01-15', 500000.00),
        ('MATH', 'Mathematics', 'Department of Mathematics', 'Dr. Johnson', 'Building B, Room 201', '+1-555-0102', 'math@school.edu', '2000-01-15', 450000.00),
        ('PHY', 'Physics', 'Department of Physics', 'Dr. Williams', 'Building C, Room 301', '+1-555-0103', 'physics@school.edu', '2001-01-15', 475000.00),
        ('CHEM', 'Chemistry', 'Department of Chemistry', 'Dr. Brown', 'Building D, Room 401', '+1-555-0104', 'chem@school.edu', '2001-01-15', 460000.00),
        ('ENG', 'Engineering', 'Department of Engineering', 'Dr. Jones', 'Building E, Room 501', '+1-555-0105', 'eng@school.edu', '2002-01-15', 600000.00)
    ]
    
    for dept in departments_data:
        try:
            c.execute(
                """INSERT OR IGNORE INTO departments (dept_code, dept_name, description, head_of_department, office_location, phone, email, established_date, budget) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                dept)
        except Exception as e:
            print(f"[WARN] Department insert warning: {e}")
    
    # Insert default settings
    settings_data = [
        ('gemini_api_key', '', 'string', 'Google Gemini API Key for AI Features'),
        ('school_name', 'Student Management System', 'string', 'School name'),
        ('academic_year', '2026-2027', 'string', 'Current academic year'),
        ('semester', 'Spring', 'string', 'Current semester'),
        ('enable_registration', 'true', 'boolean', 'Enable registration'),
        ('max_courses_per_student', '6', 'integer', 'Maximum courses'),
        ('tuition_per_credit', '500', 'integer', 'Tuition per credit'),
        ('attendance_warning_threshold', '55', 'integer', 'Probability % below which alert is sent'),
        ('total_semester_classes', '90', 'integer', 'Total planned classes per course in a semester'),
    ]
    
    for setting in settings_data:
        try:
            c.execute(
                """INSERT OR IGNORE INTO settings (setting_key, setting_value, setting_type, description) VALUES (?, ?, ?, ?)""",
                setting)
        except Exception as e:
            print(f"[WARN] Settings insert warning: {e}")
    
    conn.commit()
    conn.close()
    print("=" * 50)
    print("[OK] Database initialized successfully!")
    print("[OK] Default Users Created:")
    print("   - Admin: admin / admin123")
    print("   - Student: student / student123")
    print("   - Faculty: faculty / faculty123")
    print("=" * 50)

# Initialize database
init_db()

# ============================================================================
# DECORATORS
# ============================================================================

def login_required(f):
    """Decorator to require login"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash('Please login to access this page', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    """Decorator to require admin role"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash('Please login to access this page', 'warning')
            return redirect(url_for('login'))
        if session.get('role') != 'admin':
            flash('You do not have permission to access this page', 'error')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated_function

def student_required(f):
    """Decorator to require student role"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash('Please login to access this page', 'warning')
            return redirect(url_for('student_login'))
        if session.get('role') != 'student':
            flash('You do not have permission to access this page', 'error')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated_function

def faculty_required(f):
    """Decorator to require faculty (teacher) role"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash('Please login to access this page', 'warning')
            return redirect(url_for('faculty_login'))
        if session.get('role') != 'teacher':
            flash('You do not have permission to access this page', 'error')
            return redirect(url_for('dashboard'))
        return f(*args, **kwargs)
    return decorated_function

# ============================================================================
# AUTHENTICATION ROUTES
# ============================================================================

@app.route('/')
@app.route('/login', methods=['GET', 'POST'])
def login():
    """Combined login page"""
    if request.method == 'POST':
        username = request.form.get('username', '')
        password = request.form.get('password', '')
        
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE username = ? OR email = ?", (username, username))
        user = c.fetchone()
        conn.close()
        
        if user and check_password_hash(user['password_hash'], password):
            session.clear()
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['first_name'] = user['first_name']
            session['last_name'] = user['last_name']
            session['role'] = user['role']
            session['logged_in'] = True
            
            flash(f'Welcome back, {user["first_name"]}!', 'success')
            
            if user['role'] == 'admin':
                return redirect(url_for('admin_dashboard'))
            elif user['role'] == 'teacher':
                return redirect(url_for('teacher_dashboard'))
            else:
                return redirect(url_for('student_dashboard'))
        else:
            flash('Invalid username or password', 'error')
    
    return render_template('login.html')

@app.route('/student/login', methods=['GET', 'POST'])
def student_login():
    """Student-specific login page"""
    if request.method == 'POST':
        username = request.form.get('username', '')
        password = request.form.get('password', '')
        
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE (username = ? OR email = ?) AND role = 'student'", (username, username))
        user = c.fetchone()
        conn.close()
        
        if user and check_password_hash(user['password_hash'], password):
            session.clear()
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['first_name'] = user['first_name']
            session['last_name'] = user['last_name']
            session['role'] = user['role']
            session['logged_in'] = True
            
            flash(f'Welcome back, {user["first_name"]}!', 'success')
            return redirect(url_for('student_dashboard'))
        else:
            flash('Invalid student credentials', 'error')
    
    return render_template('student_login.html')

@app.route('/faculty/login', methods=['GET', 'POST'])
def faculty_login():
    """Faculty (teacher) specific login page"""
    if request.method == 'POST':
        username = request.form.get('username', '')
        password = request.form.get('password', '')
        
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE (username = ? OR email = ?) AND role = 'teacher'", (username, username))
        user = c.fetchone()
        conn.close()
        
        if user and check_password_hash(user['password_hash'], password):
            session.clear()
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['first_name'] = user['first_name']
            session['last_name'] = user['last_name']
            session['role'] = user['role']
            session['logged_in'] = True
            
            flash(f'Welcome back, {user["first_name"]}!', 'success')
            return redirect(url_for('teacher_dashboard'))
        else:
            flash('Invalid faculty credentials', 'error')
    
    return render_template('faculty_login.html')

@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    """Admin-specific login page"""
    if request.method == 'POST':
        username = request.form.get('username', '')
        password = request.form.get('password', '')
        
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM users WHERE (username = ? OR email = ?) AND role = 'admin'", (username, username))
        user = c.fetchone()
        conn.close()
        
        if user and check_password_hash(user['password_hash'], password):
            session.clear()
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['first_name'] = user['first_name']
            session['last_name'] = user['last_name']
            session['role'] = user['role']
            session['logged_in'] = True
            
            flash(f'Welcome back, {user["first_name"]}!', 'success')
            return redirect(url_for('admin_dashboard'))
        else:
            flash('Invalid admin credentials', 'error')
    
    return render_template('admin_login.html')

@app.route('/logout')
def logout():
    """Logout user"""
    session.clear()
    flash('You have been logged out', 'success')
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    """Redirect to role-specific dashboard"""
    if session['role'] == 'admin':
        return redirect(url_for('admin_dashboard'))
    elif session['role'] == 'teacher':
        return redirect(url_for('teacher_dashboard'))
    else:
        return redirect(url_for('student_dashboard'))

@app.route('/register', methods=['GET', 'POST'])
def register():
    """Register page"""
    if request.method == 'POST':
        username = request.form.get('username')
        email = request.form.get('email')
        password = request.form.get('password')
        first_name = request.form.get('first_name')
        last_name = request.form.get('last_name')
        role = request.form.get('role', 'student')
        
        # Validate role
        if role not in ['student', 'teacher', 'admin']:
            role = 'student'
            
        conn = get_db()
        c = conn.cursor()
        
        # Check if user exists
        c.execute("SELECT id FROM users WHERE username = ? OR email = ?", (username, email))
        if c.fetchone():
            flash('Username or email already exists', 'error')
            conn.close()
            return render_template('register.html')
        
        # Create user
        password_hash = generate_password_hash(password)
        c.execute(
            """INSERT INTO users (username, email, password_hash, first_name, last_name, role) VALUES (?, ?, ?, ?, ?, ?)""",
            (username, email, password_hash, first_name, last_name, role))
        
        # Get the user_id of the newly created user
        user_id = c.lastrowid
        year = datetime.now().strftime('%Y')
        
        if role == 'student':
            # Create student record
            c.execute("SELECT COUNT(*) as count FROM students")
            count = c.fetchone()['count'] + 1
            student_id = f'{year}{count:04d}'
            
            c.execute(
                """INSERT INTO students (user_id, student_id, first_name, last_name, email, enrollment_date, status) VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (user_id, student_id, first_name, last_name, email, datetime.now().strftime('%Y-%m-%d'), 'active'))
        
        elif role == 'teacher':
            # Create instructor record
            c.execute("SELECT COUNT(*) as count FROM instructors")
            count = c.fetchone()['count'] + 1
            instructor_id = f'FAC{year}{count:03d}'
            
            c.execute(
                """INSERT INTO instructors (user_id, instructor_id, first_name, last_name, email, joining_date, status) VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (user_id, instructor_id, first_name, last_name, email, datetime.now().strftime('%Y-%m-%d'), 'active'))
        
        conn.commit()
        conn.close()
        
        flash('Registration successful! Please login.', 'success')
        
        if role == 'teacher':
            return redirect(url_for('faculty_login'))
        elif role == 'admin':
            return redirect(url_for('admin_login'))
        else:
            return redirect(url_for('student_login'))
    
    return render_template('register.html')

@app.route('/forgot-password', methods=['GET', 'POST'])
def forgot_password():
    """Forgot password page"""
    if request.method == 'POST':
        flash('Password reset link sent to your email', 'info')
        return redirect(url_for('login'))
    return render_template('forgot_password.html')

# ============================================================================
# AI & GLOBAL ROUTES
# ============================================================================

@app.route('/api/teacher/generate_quiz/<int:course_id>', methods=['POST'])
@faculty_required
def generate_quiz(course_id):
    api_key = os.environ.get('GEMINI_API_KEY')
    if not api_key:
        return jsonify({'error': 'AI not configured (missing API Key).'})
        
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM courses WHERE id = ?", (course_id,))
    course = c.fetchone()
    if not course:
        conn.close()
        return jsonify({'error': 'Course not found.'})
        
    try:
        import google.generativeai as genai
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-2.5-flash')
        prompt = f"Create a 3-question multiple-choice quiz covering: '{course['course_name']} - {course['description']}'. Output ONLY JSON: [{{'question': '...', 'options': ['A', 'B', 'C', 'D'], 'answer': 'A'}}, ...]"
        response = model.generate_content(prompt)
        text = response.text.replace('```json', '').replace('```', '').strip()
        quiz_data = json.loads(text)
        
        from datetime import date, timedelta
        due = (date.today() + timedelta(days=7)).isoformat()
        
        c.execute(
            "INSERT INTO assignments (course_id, title, description, due_date, max_marks, created_by) VALUES (?, ?, ?, ?, ?, ?)",
            (course_id, f"Auto-Quiz: {course['course_name']}", "AI Generated Quiz:\n" + json.dumps(quiz_data, indent=2), due, 30, session['user_id'])
        )
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'Quiz generated successfully!'})
    except Exception as e:
        error_str = str(e)
        if "429" in error_str or "quota" in error_str.lower():
            # Fallback mock quiz if quota exceeded
            quiz_data = [
                {
                    "question": f"What is a core fundamental concept in {course['course_name']}?",
                    "options": ["Option A", "Option B", "Option C", "Option D"],
                    "answer": "Option A"
                },
                {
                    "question": "Which of the following describes the main goal of this unit?",
                    "options": ["Incorrect Answer 1", "Correct Answer", "Incorrect Answer 2", "Incorrect Answer 3"],
                    "answer": "Correct Answer"
                },
                {
                    "question": "How do you apply these principles in practice?",
                    "options": ["Not defined", "Via theoretical analysis", "By practical application", "Skip practice entirely"],
                    "answer": "By practical application"
                }
            ]
            from datetime import date, timedelta
            due = (date.today() + timedelta(days=7)).isoformat()
            desc = "AI Quota Exceeded. Mock Quiz Inserted:\\n" + json.dumps(quiz_data, indent=2)
            c.execute(
                "INSERT INTO assignments (course_id, title, description, due_date, max_marks, created_by) VALUES (?, ?, ?, ?, ?, ?)",
                (course_id, f"Mock Auto-Quiz: {course['course_name']}", desc, due, 30, session['user_id'])
            )
            conn.commit()
            conn.close()
            return jsonify({'success': True, 'message': 'Quiz generated successfully! (Using Mock Fallback due to Gemini Quota)'})
            
        conn.close()
        return jsonify({'error': str(e)})

# ============================================================================
# STUDENT ROUTES
# ============================================================================

@app.route('/api/student/calendar_events')
@student_required
def student_calendar_events():
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id FROM students WHERE user_id = ?", (session['user_id'],))
    student = c.fetchone()
    
    events = []
    if student:
        # Get assignments
        c.execute('''
            SELECT a.id, a.title, a.due_date, c.course_code 
            FROM assignments a
            JOIN enrollments e ON a.course_id = e.course_id
            JOIN courses c ON a.course_id = c.id
            WHERE e.student_id = ?
        ''', (student['id'],))
        for row in c.fetchall():
            if row['due_date']:
                events.append({
                    'id': f"assign_{row['id']}",
                    'title': f"{row['course_code']}: {row['title']} (Due)",
                    'start': row['due_date'],
                    'color': '#ff9f43',
                    'url': url_for('student_assignments')
                })
                
        # Get active live classroom sessions
        c.execute('''
            SELECT cs.id, cs.session_title, cs.started_at, c.course_code
            FROM classroom_sessions cs
            JOIN enrollments e ON cs.course_id = e.course_id
            JOIN courses c ON cs.course_id = c.id
            WHERE e.student_id = ? AND cs.status = 'active'
        ''', (student['id'],))
        for row in c.fetchall():
            if row['started_at']:
                events.append({
                    'id': f"sess_{row['id']}",
                    'title': f"🔴 LIVE: {row['course_code']} {row['session_title']}",
                    'start': row['started_at'],
                    'color': '#ea5455',
                    'url': url_for('classroom_respond', session_id=row['id'], action='join')
                })
    conn.close()
    return jsonify(events)


@app.route('/student/dashboard')
@student_required
def student_dashboard():
    """Student dashboard with AI attendance intelligence"""
    conn = get_db()
    c = conn.cursor()

    c.execute(
        """SELECT s.*, d.dept_name FROM students s LEFT JOIN departments d ON s.department_id = d.id WHERE s.user_id = ?""",
        (session['user_id'],))
    student = c.fetchone()

    if not student:
        conn.close()
        flash('Student profile not found', 'error')
        return redirect(url_for('dashboard'))

    c.execute(
        """SELECT c.*, e.grade, e.progress, e.status as enrollment_status, i.first_name as instructor_first, i.last_name as instructor_last FROM courses c JOIN enrollments e ON c.id = e.course_id LEFT JOIN instructors i ON c.instructor_id = i.id WHERE e.student_id = ? ORDER BY c.course_code""",
        (student['id'],))
    courses = c.fetchall()

    c.execute(
        """SELECT a.*, c.course_name FROM attendance a JOIN courses c ON a.course_id = c.id WHERE a.student_id = ? ORDER BY a.date DESC LIMIT 10""",
        (student['id'],))
    attendance = c.fetchall()

    c.execute(
        """SELECT AVG( CASE grade WHEN 'A'  THEN 4.0 WHEN 'A-' THEN 3.7 WHEN 'B+' THEN 3.3 WHEN 'B'  THEN 3.0 WHEN 'B-' THEN 2.7 WHEN 'C+' THEN 2.3 WHEN 'C'  THEN 2.0 WHEN 'C-' THEN 1.7 WHEN 'D+' THEN 1.3 WHEN 'D'  THEN 1.0 ELSE 0.0 END ) as gpa FROM enrollments WHERE student_id = ? AND grade IS NOT NULL""",
        (student['id'],))
    gpa_result = c.fetchone()
    gpa = round(gpa_result['gpa'], 2) if gpa_result and gpa_result['gpa'] else 0.0

    # AI: fetch total semester classes setting
    c.execute("SELECT setting_value FROM settings WHERE setting_key = 'total_semester_classes'")
    row = c.fetchone()
    total_semester_classes = int(row['setting_value']) if row else 90

    # AI: build per-course attendance analysis
    attendance_analysis = {}
    for course in courses:
        c.execute(
            """SELECT COUNT(CASE WHEN status IN ('present','late') THEN 1 END) as present_count,
                COUNT(*) as total_held
            FROM attendance
            WHERE student_id = ? AND course_id = ?""",
            (student['id'], course['id']))
        att = c.fetchone()
        present_count = att['present_count'] if att else 0
        total_held = att['total_held'] if att else 0
        analysis = analyse_student_course(present_count, total_held, total_semester_classes)
        attendance_analysis[course['id']] = analysis

    # Get live classes for enrolled courses
    course_ids = [c_['id'] for c_ in courses]
    live_classes = []
    if course_ids:
        placeholders = ','.join(['?' for _ in course_ids])
        sql = (
            "SELECT cs.id as session_id, cs.session_title, cs.status, cs.started_at,"
            " co.course_code, co.course_name,"
            " i.first_name || ' ' || i.last_name as instructor_name"
            " FROM classroom_sessions cs"
            " JOIN courses co ON cs.course_id = co.id"
            " LEFT JOIN instructors i ON co.instructor_id = i.id"
            " WHERE cs.course_id IN (" + placeholders + ")"
            " AND cs.status = 'active'"
            " ORDER BY cs.started_at DESC"
        )
        c.execute(sql, course_ids)
        live_classes = c.fetchall()

    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    c.execute("SELECT * FROM student_badges WHERE student_id = ? ORDER BY earned_at DESC", (student['id'],))
    student_badges = c.fetchall()

    conn.close()

    return render_template('student_dashboard.html',
                           student=student,
                           courses=courses,
                           attendance=attendance,
                           gpa=gpa,
                           unread_count=unread_count,
                           attendance_analysis=attendance_analysis,
                           live_classes=live_classes,
                           student_badges=student_badges)

@app.route('/student/courses')
@student_required
def student_courses():
    """View my courses (student view)"""
    conn = get_db()
    c = conn.cursor()
    
    c.execute("SELECT id FROM students WHERE user_id = ?", (session['user_id'],))
    student = c.fetchone()
    
    if not student:
        conn.close()
        flash('Student profile not found', 'error')
        return redirect(url_for('dashboard'))
    
    c.execute(
        """SELECT c.*, e.grade, e.progress, e.status as enrollment_status, i.first_name as instructor_first, i.last_name as instructor_last, d.dept_name FROM courses c JOIN enrollments e ON c.id = e.course_id LEFT JOIN instructors i ON c.instructor_id = i.id LEFT JOIN departments d ON c.department_id = d.id WHERE e.student_id = ? ORDER BY c.course_code""",
        (student['id'],))
    courses = c.fetchall()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('student_courses.html', courses=courses, unread_count=unread_count)

@app.route('/student/attendance')
@student_required
def student_attendance():
    """View my attendance (student view)"""
    conn = get_db()
    c = conn.cursor()
    
    c.execute("SELECT id FROM students WHERE user_id = ?", (session['user_id'],))
    student = c.fetchone()
    
    if not student:
        conn.close()
        flash('Student profile not found', 'error')
        return redirect(url_for('dashboard'))
    
    c.execute(
        """SELECT a.*, c.course_name, c.course_code FROM attendance a JOIN courses c ON a.course_id = c.id WHERE a.student_id = ? ORDER BY a.date DESC LIMIT 50""",
        (student['id'],))
    attendance = c.fetchall()
    
    c.execute(
        """SELECT COUNT(CASE WHEN status = 'present' THEN 1 END) as present_count, COUNT(CASE WHEN status = 'absent' THEN 1 END) as absent_count, COUNT(CASE WHEN status = 'late' THEN 1 END) as late_count, COUNT(CASE WHEN status = 'excused' THEN 1 END) as excused_count, COUNT(*) as total FROM attendance WHERE student_id = ?""",
        (student['id'],))
    stats = c.fetchone()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('student_attendance.html', 
                         attendance=attendance, 
                         stats=stats,
                         unread_count=unread_count)

@app.route('/student/grades')
@student_required
def student_grades():
    """View my grades (student view)"""
    conn = get_db()
    c = conn.cursor()
    
    c.execute("SELECT id FROM students WHERE user_id = ?", (session['user_id'],))
    student = c.fetchone()
    
    if not student:
        conn.close()
        flash('Student profile not found', 'error')
        return redirect(url_for('dashboard'))
    
    c.execute(
        """SELECT c.course_code, c.course_name, c.credits, e.grade, e.progress FROM enrollments e JOIN courses c ON e.course_id = c.id WHERE e.student_id = ? AND e.grade IS NOT NULL ORDER BY c.course_code""",
        (student['id'],))
    grades = c.fetchall()
    
    c.execute(
        """SELECT AVG( CASE grade WHEN 'A' THEN 4.0 WHEN 'A-' THEN 3.7 WHEN 'B+' THEN 3.3 WHEN 'B' THEN 3.0 WHEN 'B-' THEN 2.7 WHEN 'C+' THEN 2.3 WHEN 'C' THEN 2.0 WHEN 'C-' THEN 1.7 WHEN 'D+' THEN 1.3 WHEN 'D' THEN 1.0 ELSE 0.0 END ) as gpa FROM enrollments WHERE student_id = ? AND grade IS NOT NULL""",
        (student['id'],))
    gpa_result = c.fetchone()
    gpa = round(gpa_result['gpa'], 2) if gpa_result and gpa_result['gpa'] else 0.0
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('student_grades.html', grades=grades, gpa=gpa, unread_count=unread_count)

@app.route('/student/flashcards/<int:course_id>')
@student_required
def student_flashcards(course_id):
    """Generate AI Study Flashcards for a specific course"""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id FROM students WHERE user_id = ?", (session['user_id'],))
    student = c.fetchone()
    
    if not student:
        conn.close()
        flash('Student profile not found', 'error')
        return redirect(url_for('dashboard'))

    c.execute("SELECT * FROM courses WHERE id = ?", (course_id,))
    course = c.fetchone()
    conn.close()

    flashcards = []
    error = None

    try:
        import google.generativeai as genai
        import os
        import json
        api_key = os.environ.get('GEMINI_API_KEY')
        if not api_key:
            raise Exception("Gemini API key not configured")
            
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-2.5-flash')
        prompt = f"Generate 6 interactive study flashcards for a university course named '{course['course_name']}' with description '{course['description']}'. Return ONLY a raw JSON array of objects with keys 'front' (the term/concept, max 4 words) and 'back' (the definition/explanation, max 2 sentences). Example: [{{\"front\": \"Algorithm\", \"back\": \"A step-by-step procedure for solving a problem.\"}}]"
        
        response = model.generate_content(prompt)
        text = response.text.replace('```json', '').replace('```', '').strip()
        flashcards = json.loads(text)
    except Exception as e:
        print("[WARN] Flashcard Generation Error:", e)
        error = f"Failed to generate flashcards: {str(e)}"
        # Fallback Mock Data
        flashcards = [
            {"front": "Core Principle 1", "back": f"The primary foundation of {course['course_name']} involving structural analysis."},
            {"front": "Key Terminology", "back": "A frequently tested vocabulary word essential to understanding the syllabus."},
            {"front": "Practical Application", "back": "How this concept is utilized in real-world scenarios to solve complex issues."}
        ]

    return render_template('student_flashcards.html', course=course, flashcards=flashcards, error=error)

# ============================================================================
# STUDENT COURSE APPLICATION ROUTES
# ============================================================================

@app.route('/student/apply-courses')
@student_required
def student_apply_courses():
    """Browse available courses and apply"""
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT id FROM students WHERE user_id = ?", (session['user_id'],))
    student = c.fetchone()

    if not student:
        conn.close()
        flash('Student profile not found', 'error')
        return redirect(url_for('dashboard'))

    # Get all active courses with enrollment info
    c.execute(
        """SELECT c.*, d.dept_name,
            i.first_name as instructor_first, i.last_name as instructor_last,
            (SELECT COUNT(*) FROM enrollments WHERE course_id = c.id) as enrolled_count
        FROM courses c
        LEFT JOIN departments d ON c.department_id = d.id
        LEFT JOIN instructors i ON c.instructor_id = i.id
        WHERE c.status = 'active'
        ORDER BY c.course_code""")
    all_courses = c.fetchall()

    # Get already enrolled course IDs
    c.execute("SELECT course_id FROM enrollments WHERE student_id = ?", (student['id'],))
    enrolled_ids = set(row['course_id'] for row in c.fetchall())

    # Get already applied course IDs (pending or approved)
    c.execute(
        "SELECT course_id, status FROM course_applications WHERE student_id = ?",
        (student['id'],))
    applied = {row['course_id']: row['status'] for row in c.fetchall()}

    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']

    conn.close()

    return render_template('student_apply_courses.html',
                           courses=all_courses,
                           enrolled_ids=enrolled_ids,
                           applied=applied,
                           unread_count=unread_count)

@app.route('/student/apply-courses/<int:course_id>', methods=['POST'])
@student_required
def student_apply_course(course_id):
    """Submit a course application"""
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT id FROM students WHERE user_id = ?", (session['user_id'],))
    student = c.fetchone()

    if not student:
        conn.close()
        flash('Student profile not found', 'error')
        return redirect(url_for('dashboard'))

    reason = request.form.get('reason', '')

    try:
        # Check not already enrolled
        c.execute("SELECT id FROM enrollments WHERE student_id = ? AND course_id = ?",
                  (student['id'], course_id))
        if c.fetchone():
            flash('You are already enrolled in this course', 'warning')
            conn.close()
            return redirect(url_for('student_apply_courses'))

        # Check course capacity
        c.execute(
            """SELECT max_students, (SELECT COUNT(*) FROM enrollments WHERE course_id = ?) as enrolled
            FROM courses WHERE id = ?""",
            (course_id, course_id))
        course = c.fetchone()

        if course and course['enrolled'] >= course['max_students']:
            flash('This course has reached maximum capacity', 'error')
            conn.close()
            return redirect(url_for('student_apply_courses'))

        # Insert application (UNIQUE constraint handles duplicates)
        c.execute(
            """INSERT INTO course_applications (student_id, course_id, reason, status)
            VALUES (?, ?, ?, 'pending')""",
            (student['id'], course_id, reason))

        conn.commit()
        flash('Application submitted successfully! The admin will review it.', 'success')

    except sqlite3.IntegrityError:
        flash('You have already applied for this course', 'warning')
    except Exception as e:
        flash(f'Error submitting application: {str(e)}', 'error')

    conn.close()
    return redirect(url_for('student_apply_courses'))

@app.route('/student/my-applications')
@student_required
def student_my_applications():
    """View my course applications"""
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT id FROM students WHERE user_id = ?", (session['user_id'],))
    student = c.fetchone()

    if not student:
        conn.close()
        flash('Student profile not found', 'error')
        return redirect(url_for('dashboard'))

    c.execute(
        """SELECT ca.*, c.course_code, c.course_name, c.credits, d.dept_name,
            i.first_name as instructor_first, i.last_name as instructor_last
        FROM course_applications ca
        JOIN courses c ON ca.course_id = c.id
        LEFT JOIN departments d ON c.department_id = d.id
        LEFT JOIN instructors i ON c.instructor_id = i.id
        WHERE ca.student_id = ?
        ORDER BY ca.created_at DESC""",
        (student['id'],))
    applications = c.fetchall()

    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']

    conn.close()

    return render_template('student_my_applications.html',
                           applications=applications,
                           unread_count=unread_count)

# ============================================================================
# ADMIN COURSE APPLICATION ROUTES
# ============================================================================

@app.route('/admin/applications')
@admin_required
def admin_applications():
    """View and manage course applications"""
    status_filter = request.args.get('status', 'pending')

    conn = get_db()
    c = conn.cursor()

    if status_filter == 'all':
        c.execute(
            """SELECT ca.*, s.student_id as student_code, s.first_name as student_first,
                s.last_name as student_last, c.course_code, c.course_name, c.credits,
                d.dept_name
            FROM course_applications ca
            JOIN students s ON ca.student_id = s.id
            JOIN courses c ON ca.course_id = c.id
            LEFT JOIN departments d ON c.department_id = d.id
            ORDER BY ca.created_at DESC""")
    else:
        c.execute(
            """SELECT ca.*, s.student_id as student_code, s.first_name as student_first,
                s.last_name as student_last, c.course_code, c.course_name, c.credits,
                d.dept_name
            FROM course_applications ca
            JOIN students s ON ca.student_id = s.id
            JOIN courses c ON ca.course_id = c.id
            LEFT JOIN departments d ON c.department_id = d.id
            WHERE ca.status = ?
            ORDER BY ca.created_at DESC""",
            (status_filter,))
    applications = c.fetchall()

    # Count by status
    c.execute("SELECT status, COUNT(*) as count FROM course_applications GROUP BY status")
    status_counts = {row['status']: row['count'] for row in c.fetchall()}

    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']

    conn.close()

    return render_template('admin_applications.html',
                           applications=applications,
                           status_filter=status_filter,
                           status_counts=status_counts,
                           unread_count=unread_count)

@app.route('/admin/applications/<int:app_id>/approve', methods=['POST'])
@admin_required
def admin_application_approve(app_id):
    """Approve a course application and auto-enroll the student"""
    conn = get_db()
    c = conn.cursor()

    try:
        # Get the application
        c.execute(
            """SELECT ca.*, s.user_id as student_user_id, c.course_name, c.course_code,
                c.max_students,
                (SELECT COUNT(*) FROM enrollments WHERE course_id = ca.course_id) as enrolled
            FROM course_applications ca
            JOIN students s ON ca.student_id = s.id
            JOIN courses c ON ca.course_id = c.id
            WHERE ca.id = ? AND ca.status = 'pending'""",
            (app_id,))
        application = c.fetchone()

        if not application:
            flash('Application not found or already processed', 'warning')
            conn.close()
            return redirect(url_for('admin_applications'))

        # Check capacity
        if application['enrolled'] >= application['max_students']:
            flash('Course has reached maximum capacity. Cannot approve.', 'error')
            conn.close()
            return redirect(url_for('admin_applications'))

        admin_note = request.form.get('admin_note', '')

        # Update application status
        c.execute(
            """UPDATE course_applications
            SET status = 'approved', reviewed_by = ?, reviewed_at = CURRENT_TIMESTAMP, admin_note = ?
            WHERE id = ?""",
            (session['user_id'], admin_note, app_id))

        # Auto-enroll the student
        c.execute(
            """INSERT INTO enrollments (student_id, course_id, status, enrollment_date)
            VALUES (?, ?, 'enrolled', ?)""",
            (application['student_id'], application['course_id'],
             datetime.now().strftime('%Y-%m-%d %H:%M:%S')))

        # Send notification to student
        c.execute(
            """INSERT INTO notifications (user_id, title, message, type, link)
            VALUES (?, ?, ?, 'success', ?)""",
            (application['student_user_id'],
             'Application Approved!',
             f'Your application for {application["course_code"]} - {application["course_name"]} has been approved. You are now enrolled!',
             url_for('student_courses')))

        # Audit log
        c.execute(
            """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent)
            VALUES (?, ?, ?, ?, ?, ?)""",
            (session['user_id'], 'APPROVE_APPLICATION', 'course_applications', app_id,
             request.remote_addr, request.user_agent.string))

        conn.commit()
        flash(f'Application approved! Student enrolled in {application["course_code"]}.', 'success')

    except sqlite3.IntegrityError:
        flash('Student is already enrolled in this course', 'warning')
    except Exception as e:
        flash(f'Error approving application: {str(e)}', 'error')

    conn.close()
    return redirect(url_for('admin_applications'))

@app.route('/admin/applications/<int:app_id>/reject', methods=['POST'])
@admin_required
def admin_application_reject(app_id):
    """Reject a course application"""
    conn = get_db()
    c = conn.cursor()

    try:
        # Get the application
        c.execute(
            """SELECT ca.*, s.user_id as student_user_id, c.course_name, c.course_code
            FROM course_applications ca
            JOIN students s ON ca.student_id = s.id
            JOIN courses c ON ca.course_id = c.id
            WHERE ca.id = ? AND ca.status = 'pending'""",
            (app_id,))
        application = c.fetchone()

        if not application:
            flash('Application not found or already processed', 'warning')
            conn.close()
            return redirect(url_for('admin_applications'))

        admin_note = request.form.get('admin_note', '')

        # Update application status
        c.execute(
            """UPDATE course_applications
            SET status = 'rejected', reviewed_by = ?, reviewed_at = CURRENT_TIMESTAMP, admin_note = ?
            WHERE id = ?""",
            (session['user_id'], admin_note, app_id))

        # Send notification to student
        reason_text = f' Reason: {admin_note}' if admin_note else ''
        c.execute(
            """INSERT INTO notifications (user_id, title, message, type, link)
            VALUES (?, ?, ?, 'warning', ?)""",
            (application['student_user_id'],
             'Application Rejected',
             f'Your application for {application["course_code"]} - {application["course_name"]} has been rejected.{reason_text}',
             url_for('student_my_applications')))

        # Audit log
        c.execute(
            """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent)
            VALUES (?, ?, ?, ?, ?, ?)""",
            (session['user_id'], 'REJECT_APPLICATION', 'course_applications', app_id,
             request.remote_addr, request.user_agent.string))

        conn.commit()
        flash('Application rejected.', 'success')

    except Exception as e:
        flash(f'Error rejecting application: {str(e)}', 'error')

    conn.close()
    return redirect(url_for('admin_applications'))

# ============================================================================
# FACULTY (TEACHER) ROUTES
# ============================================================================

@app.route('/faculty/dashboard')
@faculty_required
def faculty_dashboard():
    """Faculty dashboard (alias for teacher_dashboard)"""
    return redirect(url_for('teacher_dashboard'))

@app.route('/teacher/dashboard')
@faculty_required
def teacher_dashboard():
    """Teacher dashboard with AI attendance alerts"""
    conn = get_db()
    c = conn.cursor()

    c.execute(
        """SELECT i.*, d.dept_name FROM instructors i LEFT JOIN departments d ON i.department_id = d.id WHERE i.user_id = ?""",
        (session['user_id'],))
    instructor = c.fetchone()

    if not instructor:
        conn.close()
        flash('Instructor profile not found', 'error')
        return redirect(url_for('dashboard'))

    c.execute(
        """SELECT c.*, d.dept_name, COUNT(DISTINCT e.student_id) as enrolled_count FROM courses c LEFT JOIN departments d ON c.department_id = d.id LEFT JOIN enrollments e ON c.id = e.course_id WHERE c.instructor_id = ? GROUP BY c.id ORDER BY c.academic_year DESC, c.semester""",
        (instructor['id'],))
    courses = c.fetchall()

    today = datetime.now().strftime('%Y-%m-%d')
    c.execute(
        """SELECT c.course_name, c.course_code, c.schedule, c.classroom, COUNT(a.id) as attendance_marked FROM courses c LEFT JOIN attendance a ON c.id = a.course_id AND a.date = ? WHERE c.instructor_id = ? GROUP BY c.id""",
        (today, instructor['id']))
    today_schedule = c.fetchall()

    # AI: fetch total semester classes setting
    c.execute("SELECT setting_value FROM settings WHERE setting_key = 'total_semester_classes'")
    row = c.fetchone()
    total_semester_classes = int(row['setting_value']) if row else 90

    # AI: build at-risk student report
    at_risk_students = []
    for course in courses:
        c.execute(
            """SELECT s.id, s.student_id, s.first_name, s.last_name FROM students s JOIN enrollments e ON s.id = e.student_id WHERE e.course_id = ?""",
            (course['id'],))
        students_in_course = c.fetchall()

        for stu in students_in_course:
            c.execute(
                """SELECT COUNT(CASE WHEN status IN ('present','late') THEN 1 END) as present_count, COUNT(*) as total_held FROM attendance WHERE student_id = ? AND course_id = ?""",
                (stu['id'], course['id']))
            att = c.fetchone()
            present_count = att['present_count'] if att else 0
            total_held = att['total_held'] if att else 0

            analysis = analyse_student_course(present_count, total_held, total_semester_classes)

            if analysis['send_alert']:
                at_risk_students.append({
                    'student_id':     stu['student_id'],
                    'name':           f"{stu['first_name']} {stu['last_name']}",
                    'course_code':    course['course_code'],
                    'course_name':    course['course_name'],
                    'attendance_pct': analysis['attendance_pct'],
                    'probability':    analysis['probability_of_80'],
                    'days_needed':    analysis['days_needed_to_recover'],
                    'risk_badge':     analysis['risk_badge'],
                    'badge_color':    analysis['badge_color'],
                    'message':        analysis['message'],
                    'user_id':        stu['id'],
                })

    # AI: send in-app notifications to at-risk students (once per day)
    for risk in at_risk_students:
        c.execute(
            """SELECT id FROM notifications WHERE user_id = (SELECT user_id FROM students WHERE id = ?) AND title LIKE ? AND DATE(created_at) = DATE('now')""",
            (risk['user_id'], f"%{risk['course_code']}%"))
        already_sent = c.fetchone()

        if not already_sent:
            badge = risk['risk_badge']
            prob  = risk['probability']
            days  = risk['days_needed']
            msg = (
                f"[{badge}] {risk['course_code']}: Your attendance is {risk['attendance_pct']}%. "
                f"Probability of reaching 80% is {prob}%. "
                f"You need to attend the next {days} consecutive class(es) to recover."
            )
            c.execute(
                """INSERT INTO notifications (user_id, title, message, type) VALUES ( (SELECT user_id FROM students WHERE id = ?), ?, ?, 'warning' )""",
                (risk['user_id'], f"Low Attendance Alert - {risk['course_code']}", msg))

    conn.commit()

    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']

    conn.close()

    return render_template('teacher_dashboard.html',
                           instructor=instructor,
                           courses=courses,
                           today_schedule=today_schedule,
                           unread_count=unread_count,
                           at_risk_students=at_risk_students)

@app.route('/faculty/courses')
@faculty_required
def faculty_courses():
    """View my courses (faculty view)"""
    conn = get_db()
    c = conn.cursor()
    
    c.execute("SELECT id FROM instructors WHERE user_id = ?", (session['user_id'],))
    instructor = c.fetchone()
    
    if not instructor:
        conn.close()
        flash('Instructor profile not found', 'error')
        return redirect(url_for('dashboard'))
    
    c.execute(
        """SELECT c.*, d.dept_name, COUNT(DISTINCT e.student_id) as enrolled_count, c.max_students FROM courses c LEFT JOIN departments d ON c.department_id = d.id LEFT JOIN enrollments e ON c.id = e.course_id WHERE c.instructor_id = ? GROUP BY c.id ORDER BY c.course_code""",
        (instructor['id'],))
    courses = c.fetchall()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('faculty_courses.html', courses=courses, unread_count=unread_count)

@app.route('/faculty/course/<int:course_id>/students')
@faculty_required
def faculty_course_students(course_id):
    """View students in a course (faculty view)"""
    conn = get_db()
    c = conn.cursor()
    
    c.execute(
        """SELECT c.* FROM courses c JOIN instructors i ON c.instructor_id = i.id WHERE c.id = ? AND i.user_id = ?""",
        (course_id, session['user_id']))
    course = c.fetchone()
    
    if not course:
        conn.close()
        flash('Course not found or access denied', 'error')
        return redirect(url_for('faculty_courses'))
    
    c.execute(
        """SELECT s.*, e.grade, e.progress, e.enrollment_date FROM students s JOIN enrollments e ON s.id = e.student_id WHERE e.course_id = ? ORDER BY s.last_name, s.first_name""",
        (course_id,))
    students = c.fetchall()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('faculty_course_students.html',
                         course=course,
                         students=students,
                         unread_count=unread_count)

@app.route('/faculty/attendance/mark', methods=['GET', 'POST'])
@faculty_required
def faculty_mark_attendance():
    """Mark attendance for a course"""
    if request.method == 'POST':
        course_id = request.form.get('course_id')
        date = request.form.get('date', datetime.now().strftime('%Y-%m-%d'))
        attendance_data = request.form.getlist('attendance[]')
        
        try:
            conn = get_db()
            c = conn.cursor()
            
            c.execute(
                """SELECT c.id FROM courses c JOIN instructors i ON c.instructor_id = i.id WHERE c.id = ? AND i.user_id = ?""",
                (course_id, session['user_id']))
            if not c.fetchone():
                flash('Access denied', 'error')
                return redirect(url_for('faculty_courses'))
            
            for item in attendance_data:
                student_id, status = item.split(':')
                
                c.execute(
                    """INSERT OR REPLACE INTO attendance (student_id, course_id, date, status, marked_by) VALUES (?, ?, ?, ?, ?)""",
                    (student_id, course_id, date, status, session['user_id']))
            
            conn.commit()
            conn.close()
            
            flash('Attendance marked successfully', 'success')
            return redirect(url_for('faculty_course_students', course_id=course_id))
            
        except Exception as e:
            flash(f'Error marking attendance: {str(e)}', 'error')
    
    course_id = request.args.get('course_id')
    date = request.args.get('date', datetime.now().strftime('%Y-%m-%d'))
    
    conn = get_db()
    c = conn.cursor()
    
    c.execute(
        """SELECT c.* FROM courses c JOIN instructors i ON c.instructor_id = i.id WHERE c.id = ? AND i.user_id = ?""",
        (course_id, session['user_id']))
    course = c.fetchone()
    
    if not course:
        conn.close()
        flash('Course not found or access denied', 'error')
        return redirect(url_for('faculty_courses'))
    
    c.execute(
        """SELECT s.id, s.student_id, s.first_name, s.last_name, a.status, a.check_in_time FROM students s JOIN enrollments e ON s.id = e.student_id LEFT JOIN attendance a ON s.id = a.student_id AND a.course_id = ? AND a.date = ? WHERE e.course_id = ? ORDER BY s.last_name, s.first_name""",
        (course_id, date, course_id))
    students = c.fetchall()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('faculty_mark_attendance.html',
                         course=course,
                         students=students,
                         date=date,
                         unread_count=unread_count)

@app.route('/faculty/grades/update', methods=['POST'])
@faculty_required
def faculty_update_grade():
    """Update student grade"""
    enrollment_id = request.form.get('enrollment_id')
    grade = request.form.get('grade')
    course_id = request.form.get('course_id')
    
    try:
        conn = get_db()
        c = conn.cursor()
        
        c.execute(
            """SELECT e.id FROM enrollments e JOIN courses c ON e.course_id = c.id JOIN instructors i ON c.instructor_id = i.id WHERE e.id = ? AND i.user_id = ?""",
            (enrollment_id, session['user_id']))
        
        if not c.fetchone():
            flash('Access denied', 'error')
            return redirect(url_for('faculty_courses'))
        
        c.execute(
            """UPDATE enrollments SET grade = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?""",
            (grade, enrollment_id))
        
        conn.commit()
        conn.close()
        
        flash('Grade updated successfully', 'success')
        
    except Exception as e:
        flash(f'Error updating grade: {str(e)}', 'error')
    
    return redirect(request.referrer or url_for('faculty_courses'))

# ============================================================================
# ADMIN ROUTES
# ============================================================================

@app.route('/admin/dashboard')
@admin_required
def admin_dashboard():
    """Admin dashboard"""
    conn = get_db()
    c = conn.cursor()
    
    stats = {}
    c.execute("SELECT COUNT(*) as count FROM students WHERE status = 'active'")
    stats['total_students'] = c.fetchone()['count']
    
    c.execute("SELECT COUNT(*) as count FROM courses WHERE status = 'active'")
    stats['total_courses'] = c.fetchone()['count']
    
    c.execute("SELECT COUNT(*) as count FROM instructors WHERE status = 'active'")
    stats['total_teachers'] = c.fetchone()['count']
    
    c.execute("SELECT COUNT(*) as count FROM departments")
    stats['total_departments'] = c.fetchone()['count']
    
    c.execute(
        """SELECT a.*, u.username FROM audit_log a JOIN users u ON a.user_id = u.id ORDER BY a.created_at DESC LIMIT 10""",
        )
    recent_activities = c.fetchall()
    
    c.execute(
        """SELECT d.dept_name, COUNT(DISTINCT s.id) as student_count, COUNT(DISTINCT c.id) as course_count FROM departments d LEFT JOIN students s ON d.id = s.department_id LEFT JOIN courses c ON d.id = c.department_id GROUP BY d.id""",
        )
    department_stats = c.fetchall()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']

    # Get pending applications count
    c.execute("SELECT COUNT(*) as count FROM course_applications WHERE status = 'pending'")
    pending_applications = c.fetchone()['count']
    
    conn.close()
    
    return render_template('admin_dashboard.html',
                         stats=stats,
                         recent_activities=recent_activities,
                         department_stats=department_stats,
                         unread_count=unread_count,
                         pending_applications=pending_applications)

@app.route('/admin/students')
@admin_required
def student_list():
    """List all students"""
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '')
    per_page = app.config['ITEMS_PER_PAGE']
    offset = (page - 1) * per_page
    
    conn = get_db()
    c = conn.cursor()
    
    if search:
        c.execute(
            """SELECT s.*, d.dept_name FROM students s LEFT JOIN departments d ON s.department_id = d.id WHERE s.first_name LIKE ? OR s.last_name LIKE ? OR s.student_id LIKE ? OR s.email LIKE ? ORDER BY s.last_name, s.first_name LIMIT ? OFFSET ?""",
            (f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%', per_page, offset))
        students = c.fetchall()
        
        c.execute(
            """SELECT COUNT(*) as count FROM students WHERE first_name LIKE ? OR last_name LIKE ? OR student_id LIKE ? OR email LIKE ?""",
            (f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%'))
        total = c.fetchone()['count']
    else:
        c.execute(
            """SELECT s.*, d.dept_name FROM students s LEFT JOIN departments d ON s.department_id = d.id ORDER BY s.last_name, s.first_name LIMIT ? OFFSET ?""",
            (per_page, offset))
        students = c.fetchall()
        
        c.execute('SELECT COUNT(*) as count FROM students')
        total = c.fetchone()['count']
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    total_pages = (total + per_page - 1) // per_page
    
    return render_template('student_list.html',
                         students=students,
                         page=page,
                         total_pages=total_pages,
                         search=search,
                         unread_count=unread_count)

@app.route('/admin/students/add', methods=['GET', 'POST'])
@admin_required
def student_add():
    """Add new student"""
    if request.method == 'POST':
        try:
            conn = get_db()
            c = conn.cursor()
            
            year = datetime.now().strftime('%Y')
            c.execute("SELECT COUNT(*) as count FROM students WHERE student_id LIKE ?", (f'{year}%',))
            count = c.fetchone()['count'] + 1
            student_id = f'{year}{count:04d}'
            
            password_hash = generate_password_hash('welcome123')
            c.execute(
                """INSERT INTO users (username, email, password_hash, first_name, last_name, role) VALUES (?, ?, ?, ?, ?, ?)""",
                ( request.form.get('email').split('@')[0], request.form.get('email'), password_hash, request.form.get('first_name'), request.form.get('last_name'), 'student' ))
            
            user_id = c.lastrowid
            
            c.execute(
                """INSERT INTO students ( user_id, student_id, first_name, last_name, date_of_birth, gender, email, phone, address, city, state, zip_code, country, department_id, enrollment_date, guardian_name, guardian_phone, guardian_email, emergency_contact, medical_info, status ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                ( user_id, student_id, request.form.get('first_name'), request.form.get('last_name'), request.form.get('date_of_birth'), request.form.get('gender'), request.form.get('email'), request.form.get('phone'), request.form.get('address'), request.form.get('city'), request.form.get('state'), request.form.get('zip_code'), request.form.get('country', 'USA'), request.form.get('department_id'), request.form.get('enrollment_date', datetime.now().strftime('%Y-%m-%d')), request.form.get('guardian_name'), request.form.get('guardian_phone'), request.form.get('guardian_email'), request.form.get('emergency_contact'), request.form.get('medical_info'), 'active' ))
            
            c.execute(
                """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?)""",
                ( session['user_id'], 'INSERT', 'students', c.lastrowid, request.remote_addr, request.user_agent.string ))
            
            conn.commit()
            conn.close()
            
            flash(f'Student added successfully! Student ID: {student_id}', 'success')
            return redirect(url_for('student_list'))
            
        except Exception as e:
            flash(f'Error adding student: {str(e)}', 'error')
            return redirect(url_for('student_add'))
    
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM departments ORDER BY dept_name")
    departments = c.fetchall()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('student_form.html', 
                         departments=departments, 
                         student=None,
                         form_title='Add New Student',
                         unread_count=unread_count)

@app.route('/admin/students/edit/<int:id>', methods=['GET', 'POST'])
@admin_required
def student_edit(id):
    """Edit student"""
    conn = get_db()
    c = conn.cursor()
    
    if request.method == 'POST':
        try:
            c.execute(
                """UPDATE students SET first_name = ?, last_name = ?, date_of_birth = ?, gender = ?, email = ?, phone = ?, address = ?, city = ?, state = ?, zip_code = ?, country = ?, department_id = ?, guardian_name = ?, guardian_phone = ?, guardian_email = ?, emergency_contact = ?, medical_info = ?, status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?""",
                ( request.form.get('first_name'), request.form.get('last_name'), request.form.get('date_of_birth'), request.form.get('gender'), request.form.get('email'), request.form.get('phone'), request.form.get('address'), request.form.get('city'), request.form.get('state'), request.form.get('zip_code'), request.form.get('country'), request.form.get('department_id'), request.form.get('guardian_name'), request.form.get('guardian_phone'), request.form.get('guardian_email'), request.form.get('emergency_contact'), request.form.get('medical_info'), request.form.get('status'), id ))
            
            c.execute(
                """UPDATE users SET email = ?, first_name = ?, last_name = ?, updated_at = CURRENT_TIMESTAMP WHERE id = (SELECT user_id FROM students WHERE id = ?)""",
                ( request.form.get('email'), request.form.get('first_name'), request.form.get('last_name'), id ))
            
            c.execute(
                """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?)""",
                ( session['user_id'], 'UPDATE', 'students', id, request.remote_addr, request.user_agent.string ))
            
            conn.commit()
            conn.close()
            
            flash('Student updated successfully', 'success')
            return redirect(url_for('student_list'))
            
        except Exception as e:
            flash(f'Error updating student: {str(e)}', 'error')
    
    c.execute(
        """SELECT s.*, u.username FROM students s JOIN users u ON s.user_id = u.id WHERE s.id = ?""",
        (id,))
    student = c.fetchone()
    
    c.execute("SELECT * FROM departments ORDER BY dept_name")
    departments = c.fetchall()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('student_form.html',
                         student=student,
                         departments=departments,
                         form_title='Edit Student',
                         unread_count=unread_count)

@app.route('/admin/students/delete/<int:id>', methods=['POST'])
@admin_required
def student_delete(id):
    """Delete student"""
    try:
        conn = get_db()
        c = conn.cursor()
        
        c.execute('SELECT user_id FROM students WHERE id = ?', (id,))
        student = c.fetchone()
        
        if student:
            c.execute('DELETE FROM students WHERE id = ?', (id,))
            c.execute('DELETE FROM users WHERE id = ?', (student['user_id'],))
            
            c.execute(
                """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?)""",
                ( session['user_id'], 'DELETE', 'students', id, request.remote_addr, request.user_agent.string ))
            
            conn.commit()
            flash('Student deleted successfully', 'success')
        
        conn.close()
        
    except Exception as e:
        flash(f'Error deleting student: {str(e)}', 'error')
    
    return redirect(url_for('student_list'))

@app.route('/admin/courses')
@admin_required
def course_list():
    """List all courses"""
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '')
    per_page = app.config['ITEMS_PER_PAGE']
    offset = (page - 1) * per_page
    
    conn = get_db()
    c = conn.cursor()
    
    if search:
        c.execute(
            """SELECT c.*, d.dept_name, i.first_name as instructor_first, i.last_name as instructor_last FROM courses c LEFT JOIN departments d ON c.department_id = d.id LEFT JOIN instructors i ON c.instructor_id = i.id WHERE c.course_code LIKE ? OR c.course_name LIKE ? OR d.dept_name LIKE ? ORDER BY c.course_code LIMIT ? OFFSET ?""",
            (f'%{search}%', f'%{search}%', f'%{search}%', per_page, offset))
        courses = c.fetchall()
        
        c.execute(
            """SELECT COUNT(*) as count FROM courses c LEFT JOIN departments d ON c.department_id = d.id WHERE c.course_code LIKE ? OR c.course_name LIKE ? OR d.dept_name LIKE ?""",
            (f'%{search}%', f'%{search}%', f'%{search}%'))
        total = c.fetchone()['count']
    else:
        c.execute(
            """SELECT c.*, d.dept_name, i.first_name as instructor_first, i.last_name as instructor_last FROM courses c LEFT JOIN departments d ON c.department_id = d.id LEFT JOIN instructors i ON c.instructor_id = i.id ORDER BY c.course_code LIMIT ? OFFSET ?""",
            (per_page, offset))
        courses = c.fetchall()
        
        c.execute('SELECT COUNT(*) as count FROM courses')
        total = c.fetchone()['count']
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    total_pages = (total + per_page - 1) // per_page
    
    return render_template('course_list.html',
                         courses=courses,
                         page=page,
                         total_pages=total_pages,
                         search=search,
                         unread_count=unread_count)

@app.route('/admin/courses/add', methods=['GET', 'POST'])
@admin_required
def course_add():
    """Add new course"""
    if request.method == 'POST':
        try:
            conn = get_db()
            c = conn.cursor()
            
            c.execute(
                """INSERT INTO courses ( course_code, course_name, description, credits, department_id, instructor_id, semester, academic_year, max_students, status, start_date, end_date, schedule, classroom, prerequisites, syllabus ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                ( request.form.get('course_code'), request.form.get('course_name'), request.form.get('description'), request.form.get('credits', 3), request.form.get('department_id'), request.form.get('instructor_id'), request.form.get('semester'), request.form.get('academic_year'), request.form.get('max_students', 30), request.form.get('status', 'active'), request.form.get('start_date'), request.form.get('end_date'), request.form.get('schedule'), request.form.get('classroom'), request.form.get('prerequisites'), request.form.get('syllabus')
            ))
            
            c.execute(
                """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?)""",
                ( session['user_id'], 'INSERT', 'courses', c.lastrowid, request.remote_addr, request.user_agent.string ))
            
            conn.commit()
            conn.close()
            
            flash('Course added successfully', 'success')
            return redirect(url_for('course_list'))
            
        except Exception as e:
            flash(f'Error adding course: {str(e)}', 'error')
    
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM departments ORDER BY dept_name")
    departments = c.fetchall()
    c.execute("SELECT * FROM instructors WHERE status = 'active' ORDER BY last_name, first_name")
    instructors = c.fetchall()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('course_form.html', 
                         departments=departments,
                         instructors=instructors,
                         course=None,
                         form_title='Add New Course',
                         unread_count=unread_count)

@app.route('/admin/courses/edit/<int:id>', methods=['GET', 'POST'])
@admin_required
def course_edit(id):
    """Edit course"""
    conn = get_db()
    c = conn.cursor()
    
    if request.method == 'POST':
        try:
            c.execute(
                """UPDATE courses SET course_code = ?, course_name = ?, description = ?, credits = ?, department_id = ?, instructor_id = ?, semester = ?, academic_year = ?, max_students = ?, status = ?, start_date = ?, end_date = ?, schedule = ?, classroom = ?, prerequisites = ?, syllabus = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?""",
                ( request.form.get('course_code'), request.form.get('course_name'), request.form.get('description'), request.form.get('credits'), request.form.get('department_id'), request.form.get('instructor_id'), request.form.get('semester'), request.form.get('academic_year'), request.form.get('max_students'), request.form.get('status'), request.form.get('start_date'), request.form.get('end_date'), request.form.get('schedule'), request.form.get('classroom'), request.form.get('prerequisites'), request.form.get('syllabus'), id ))
            
            c.execute(
                """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?)""",
                ( session['user_id'], 'UPDATE', 'courses', id, request.remote_addr, request.user_agent.string ))
            
            conn.commit()
            conn.close()
            
            flash('Course updated successfully', 'success')
            return redirect(url_for('course_list'))
            
        except Exception as e:
            flash(f'Error updating course: {str(e)}', 'error')
    
    c.execute('SELECT * FROM courses WHERE id = ?', (id,))
    course = c.fetchone()
    
    c.execute("SELECT * FROM departments ORDER BY dept_name")
    departments = c.fetchall()
    
    c.execute("SELECT * FROM instructors WHERE status = 'active' ORDER BY last_name, first_name")
    instructors = c.fetchall()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('course_form.html',
                         course=course,
                         departments=departments,
                         instructors=instructors,
                         form_title='Edit Course',
                         unread_count=unread_count)

@app.route('/admin/courses/delete/<int:id>', methods=['POST'])
@admin_required
def course_delete(id):
    """Delete course"""
    try:
        conn = get_db()
        c = conn.cursor()
        
        c.execute('DELETE FROM courses WHERE id = ?', (id,))
        
        c.execute(
            """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?)""",
            ( session['user_id'], 'DELETE', 'courses', id, request.remote_addr, request.user_agent.string ))
        
        conn.commit()
        conn.close()
        
        flash('Course deleted successfully', 'success')
        
    except Exception as e:
        flash(f'Error deleting course: {str(e)}', 'error')
    
    return redirect(url_for('course_list'))

@app.route('/admin/instructors')
@admin_required
def instructor_list():
    """List all instructors"""
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '')
    per_page = app.config['ITEMS_PER_PAGE']
    offset = (page - 1) * per_page
    
    conn = get_db()
    c = conn.cursor()
    
    if search:
        c.execute(
            """SELECT i.*, d.dept_name FROM instructors i LEFT JOIN departments d ON i.department_id = d.id WHERE i.first_name LIKE ? OR i.last_name LIKE ? OR i.instructor_id LIKE ? OR i.email LIKE ? ORDER BY i.last_name, i.first_name LIMIT ? OFFSET ?""",
            (f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%', per_page, offset))
        instructors = c.fetchall()
        
        c.execute(
            """SELECT COUNT(*) as count FROM instructors WHERE first_name LIKE ? OR last_name LIKE ? OR instructor_id LIKE ? OR email LIKE ?""",
            (f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%'))
        total = c.fetchone()['count']
    else:
        c.execute(
            """SELECT i.*, d.dept_name FROM instructors i LEFT JOIN departments d ON i.department_id = d.id ORDER BY i.last_name, i.first_name LIMIT ? OFFSET ?""",
            (per_page, offset))
        instructors = c.fetchall()
        
        c.execute('SELECT COUNT(*) as count FROM instructors')
        total = c.fetchone()['count']
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    total_pages = (total + per_page - 1) // per_page
    
    return render_template('instructor_list.html',
                         instructors=instructors,
                         page=page,
                         total_pages=total_pages,
                         search=search,
                         unread_count=unread_count)

@app.route('/admin/instructors/add', methods=['GET', 'POST'])
@admin_required
def instructor_add():
    """Add new instructor"""
    if request.method == 'POST':
        try:
            conn = get_db()
            c = conn.cursor()
            
            year = datetime.now().strftime('%Y')
            c.execute("SELECT COUNT(*) as count FROM instructors WHERE instructor_id LIKE ?", (f'FAC%',))
            count = c.fetchone()['count'] + 1
            instructor_id = f'FAC{count:03d}'
            
            password_hash = generate_password_hash('welcome123')
            c.execute(
                """INSERT INTO users (username, email, password_hash, first_name, last_name, role) VALUES (?, ?, ?, ?, ?, ?)""",
                ( request.form.get('email').split('@')[0], request.form.get('email'), password_hash, request.form.get('first_name'), request.form.get('last_name'), 'teacher' ))
            
            user_id = c.lastrowid
            
            c.execute(
                """INSERT INTO instructors ( user_id, instructor_id, first_name, last_name, email, phone, department_id, qualification, specialization, joining_date, salary, status, office_hours ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                ( user_id, instructor_id, request.form.get('first_name'), request.form.get('last_name'), request.form.get('email'), request.form.get('phone'), request.form.get('department_id'), request.form.get('qualification'), request.form.get('specialization'), request.form.get('joining_date', datetime.now().strftime('%Y-%m-%d')), request.form.get('salary'), request.form.get('status', 'active'), request.form.get('office_hours')
            ))
            
            c.execute(
                """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?)""",
                ( session['user_id'], 'INSERT', 'instructors', c.lastrowid, request.remote_addr, request.user_agent.string ))
            
            conn.commit()
            conn.close()
            
            flash(f'Instructor added successfully! Instructor ID: {instructor_id}', 'success')
            return redirect(url_for('instructor_list'))
            
        except Exception as e:
            flash(f'Error adding instructor: {str(e)}', 'error')
            return redirect(url_for('instructor_add'))
    
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM departments ORDER BY dept_name")
    departments = c.fetchall()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('instructor_form.html', 
                         departments=departments, 
                         instructor=None,
                         form_title='Add New Instructor',
                         unread_count=unread_count)

@app.route('/admin/instructors/edit/<int:id>', methods=['GET', 'POST'])
@admin_required
def instructor_edit(id):
    """Edit instructor"""
    conn = get_db()
    c = conn.cursor()
    
    if request.method == 'POST':
        try:
            c.execute(
                """UPDATE instructors SET first_name = ?, last_name = ?, email = ?, phone = ?, department_id = ?, qualification = ?, specialization = ?, salary = ?, status = ?, office_hours = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?""",
                ( request.form.get('first_name'), request.form.get('last_name'), request.form.get('email'), request.form.get('phone'), request.form.get('department_id'), request.form.get('qualification'), request.form.get('specialization'), request.form.get('salary'), request.form.get('status'), request.form.get('office_hours'), id ))
            
            c.execute(
                """UPDATE users SET email = ?, first_name = ?, last_name = ?, updated_at = CURRENT_TIMESTAMP WHERE id = (SELECT user_id FROM instructors WHERE id = ?)""",
                ( request.form.get('email'), request.form.get('first_name'), request.form.get('last_name'), id ))
            
            c.execute(
                """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?)""",
                ( session['user_id'], 'UPDATE', 'instructors', id, request.remote_addr, request.user_agent.string ))
            
            conn.commit()
            conn.close()
            
            flash('Instructor updated successfully', 'success')
            return redirect(url_for('instructor_list'))
            
        except Exception as e:
            flash(f'Error updating instructor: {str(e)}', 'error')
    
    c.execute(
        """SELECT i.*, u.username FROM instructors i JOIN users u ON i.user_id = u.id WHERE i.id = ?""",
        (id,))
    instructor = c.fetchone()
    
    c.execute("SELECT * FROM departments ORDER BY dept_name")
    departments = c.fetchall()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('instructor_form.html',
                         instructor=instructor,
                         departments=departments,
                         form_title='Edit Instructor',
                         unread_count=unread_count)

@app.route('/admin/instructors/delete/<int:id>', methods=['POST'])
@admin_required
def instructor_delete(id):
    """Delete instructor"""
    try:
        conn = get_db()
        c = conn.cursor()
        
        c.execute('SELECT user_id FROM instructors WHERE id = ?', (id,))
        instructor = c.fetchone()
        
        if instructor:
            c.execute('DELETE FROM instructors WHERE id = ?', (id,))
            c.execute('DELETE FROM users WHERE id = ?', (instructor['user_id'],))
            
            c.execute(
                """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?)""",
                ( session['user_id'], 'DELETE', 'instructors', id, request.remote_addr, request.user_agent.string ))
            
            conn.commit()
            flash('Instructor deleted successfully', 'success')
        
        conn.close()
        
    except Exception as e:
        flash(f'Error deleting instructor: {str(e)}', 'error')
    
    return redirect(url_for('instructor_list'))

@app.route('/admin/departments')
@admin_required
def department_list():
    """List all departments"""
    conn = get_db()
    c = conn.cursor()
    
    c.execute(
        """SELECT d.*, COUNT(DISTINCT s.id) as student_count, COUNT(DISTINCT i.id) as instructor_count, COUNT(DISTINCT c.id) as course_count FROM departments d LEFT JOIN students s ON d.id = s.department_id LEFT JOIN instructors i ON d.id = i.department_id LEFT JOIN courses c ON d.id = c.department_id GROUP BY d.id ORDER BY d.dept_code""",
        )
    departments = c.fetchall()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('department_list.html', 
                         departments=departments,
                         unread_count=unread_count)

@app.route('/admin/departments/add', methods=['GET', 'POST'])
@admin_required
def department_add():
    """Add new department"""
    if request.method == 'POST':
        try:
            conn = get_db()
            c = conn.cursor()
            
            c.execute(
                """INSERT INTO departments ( dept_code, dept_name, description, head_of_department, office_location, phone, email, established_date, budget ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                ( request.form.get('dept_code'), request.form.get('dept_name'), request.form.get('description'), request.form.get('head_of_department'), request.form.get('office_location'), request.form.get('phone'), request.form.get('email'), request.form.get('established_date'), request.form.get('budget')
            ))
            
            c.execute(
                """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?)""",
                ( session['user_id'], 'INSERT', 'departments', c.lastrowid, request.remote_addr, request.user_agent.string ))
            
            conn.commit()
            conn.close()
            
            flash('Department added successfully', 'success')
            return redirect(url_for('department_list'))
            
        except Exception as e:
            flash(f'Error adding department: {str(e)}', 'error')
    
    # Get unread notifications count
    conn = get_db()
    c = conn.cursor()
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    conn.close()
    
    return render_template('department_form.html', 
                         department=None,
                         form_title='Add New Department',
                         unread_count=unread_count)

@app.route('/admin/departments/edit/<int:id>', methods=['GET', 'POST'])
@admin_required
def department_edit(id):
    """Edit department"""
    conn = get_db()
    c = conn.cursor()
    
    if request.method == 'POST':
        try:
            c.execute(
                """UPDATE departments SET dept_code = ?, dept_name = ?, description = ?, head_of_department = ?, office_location = ?, phone = ?, email = ?, established_date = ?, budget = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?""",
                ( request.form.get('dept_code'), request.form.get('dept_name'), request.form.get('description'), request.form.get('head_of_department'), request.form.get('office_location'), request.form.get('phone'), request.form.get('email'), request.form.get('established_date'), request.form.get('budget'), id ))
            
            c.execute(
                """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?)""",
                ( session['user_id'], 'UPDATE', 'departments', id, request.remote_addr, request.user_agent.string ))
            
            conn.commit()
            conn.close()
            
            flash('Department updated successfully', 'success')
            return redirect(url_for('department_list'))
            
        except Exception as e:
            flash(f'Error updating department: {str(e)}', 'error')
    
    c.execute('SELECT * FROM departments WHERE id = ?', (id,))
    department = c.fetchone()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('department_form.html',
                         department=department,
                         form_title='Edit Department',
                         unread_count=unread_count)

@app.route('/admin/departments/delete/<int:id>', methods=['POST'])
@admin_required
def department_delete(id):
    """Delete department"""
    try:
        conn = get_db()
        c = conn.cursor()
        
        c.execute('DELETE FROM departments WHERE id = ?', (id,))
        
        c.execute(
            """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?)""",
            ( session['user_id'], 'DELETE', 'departments', id, request.remote_addr, request.user_agent.string ))
        
        conn.commit()
        conn.close()
        
        flash('Department deleted successfully', 'success')
        
    except Exception as e:
        flash(f'Error deleting department: {str(e)}', 'error')
    
    return redirect(url_for('department_list'))

@app.route('/admin/enrollments')
@admin_required
def enrollment_list():
    """List all enrollments"""
    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '')
    per_page = app.config['ITEMS_PER_PAGE']
    offset = (page - 1) * per_page
    
    conn = get_db()
    c = conn.cursor()
    
    if search:
        c.execute(
            """SELECT e.*, s.first_name as student_first, s.last_name as student_last, s.student_id, c.course_name, c.course_code FROM enrollments e JOIN students s ON e.student_id = s.id JOIN courses c ON e.course_id = c.id WHERE s.first_name LIKE ? OR s.last_name LIKE ? OR s.student_id LIKE ? OR c.course_name LIKE ? OR c.course_code LIKE ? ORDER BY e.enrollment_date DESC LIMIT ? OFFSET ?""",
            (f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%', per_page, offset))
        enrollments = c.fetchall()
        
        c.execute(
            """SELECT COUNT(*) as count FROM enrollments e JOIN students s ON e.student_id = s.id JOIN courses c ON e.course_id = c.id WHERE s.first_name LIKE ? OR s.last_name LIKE ? OR s.student_id LIKE ? OR c.course_name LIKE ? OR c.course_code LIKE ?""",
            (f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%', f'%{search}%'))
        total = c.fetchone()['count']
    else:
        c.execute(
            """SELECT e.*, s.first_name as student_first, s.last_name as student_last, s.student_id, c.course_name, c.course_code FROM enrollments e JOIN students s ON e.student_id = s.id JOIN courses c ON e.course_id = c.id ORDER BY e.enrollment_date DESC LIMIT ? OFFSET ?""",
            (per_page, offset))
        enrollments = c.fetchall()
        
        c.execute('SELECT COUNT(*) as count FROM enrollments')
        total = c.fetchone()['count']
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    total_pages = (total + per_page - 1) // per_page
    
    return render_template('enrollment_list.html',
                         enrollments=enrollments,
                         page=page,
                         total_pages=total_pages,
                         search=search,
                         unread_count=unread_count)

@app.route('/admin/enrollments/add', methods=['GET', 'POST'])
@admin_required
def enrollment_add():
    """Add new enrollment"""
    if request.method == 'POST':
        try:
            conn = get_db()
            c = conn.cursor()
            
            student_id = request.form.get('student_id')
            course_id = request.form.get('course_id')
            
            # Check if already enrolled
            c.execute('SELECT id FROM enrollments WHERE student_id = ? AND course_id = ?',
                     (student_id, course_id))
            if c.fetchone():
                flash('Student already enrolled in this course', 'error')
                return redirect(url_for('enrollment_add'))
            
            # Check course capacity
            c.execute(
                """SELECT max_students, (SELECT COUNT(*) FROM enrollments WHERE course_id = ?) as enrolled FROM courses WHERE id = ?""",
                (course_id, course_id))
            course = c.fetchone()
            
            if course and course['enrolled'] >= course['max_students']:
                flash('Course has reached maximum capacity', 'error')
                return redirect(url_for('enrollment_add'))
            
            c.execute(
                """INSERT INTO enrollments (student_id, course_id, status, enrollment_date) VALUES (?, ?, ?, ?)""",
                ( student_id, course_id, request.form.get('status', 'enrolled'), request.form.get('enrollment_date', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            ))
            
            c.execute(
                """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?)""",
                ( session['user_id'], 'INSERT', 'enrollments', c.lastrowid, request.remote_addr, request.user_agent.string ))
            
            conn.commit()
            conn.close()
            
            flash('Enrollment added successfully', 'success')
            return redirect(url_for('enrollment_list'))
            
        except Exception as e:
            flash(f'Error adding enrollment: {str(e)}', 'error')
    
    conn = get_db()
    c = conn.cursor()
    
    c.execute('SELECT id, student_id, first_name, last_name FROM students WHERE status = "active" ORDER BY last_name, first_name')
    students = c.fetchall()
    
    c.execute('SELECT id, course_code, course_name FROM courses WHERE status = "active" ORDER BY course_code')
    courses = c.fetchall()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('enrollment_form.html',
                         students=students,
                         courses=courses,
                         enrollment=None,
                         form_title='Add New Enrollment',
                         unread_count=unread_count)

@app.route('/admin/enrollments/edit/<int:id>', methods=['GET', 'POST'])
@admin_required
def enrollment_edit(id):
    """Edit enrollment"""
    conn = get_db()
    c = conn.cursor()
    
    if request.method == 'POST':
        try:
            c.execute(
                """UPDATE enrollments SET status = ?, grade = ?, progress = ?, notes = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?""",
                ( request.form.get('status'), request.form.get('grade'), request.form.get('progress', 0), request.form.get('notes'), id ))
            
            c.execute(
                """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?)""",
                ( session['user_id'], 'UPDATE', 'enrollments', id, request.remote_addr, request.user_agent.string ))
            
            conn.commit()
            conn.close()
            
            flash('Enrollment updated successfully', 'success')
            return redirect(url_for('enrollment_list'))
            
        except Exception as e:
            flash(f'Error updating enrollment: {str(e)}', 'error')
    
    c.execute(
        """SELECT e.*, s.student_id, s.first_name as student_first, s.last_name as student_last, c.course_code, c.course_name FROM enrollments e JOIN students s ON e.student_id = s.id JOIN courses c ON e.course_id = c.id WHERE e.id = ?""",
        (id,))
    enrollment = c.fetchone()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('enrollment_form.html',
                         enrollment=enrollment,
                         form_title='Edit Enrollment',
                         unread_count=unread_count)

@app.route('/admin/enrollments/delete/<int:id>', methods=['POST'])
@admin_required
def enrollment_delete(id):
    """Delete enrollment"""
    try:
        conn = get_db()
        c = conn.cursor()
        
        c.execute('DELETE FROM enrollments WHERE id = ?', (id,))
        
        c.execute(
            """INSERT INTO audit_log (user_id, action, table_name, record_id, ip_address, user_agent) VALUES (?, ?, ?, ?, ?, ?)""",
            ( session['user_id'], 'DELETE', 'enrollments', id, request.remote_addr, request.user_agent.string ))
        
        conn.commit()
        conn.close()
        
        flash('Enrollment deleted successfully', 'success')
        
    except Exception as e:
        flash(f'Error deleting enrollment: {str(e)}', 'error')
    
    return redirect(url_for('enrollment_list'))

@app.route('/admin/settings', methods=['GET', 'POST'])
@admin_required
def settings():
    """Settings page"""
    if request.method == 'POST':
        try:
            conn = get_db()
            c = conn.cursor()
            
            for key in request.form:
                if key.startswith('setting_'):
                    setting_key = key[8:]  # Remove 'setting_' prefix
                    setting_value = request.form.get(key)
                    
                    c.execute(
                        """UPDATE settings SET setting_value = ?, updated_by = ?, updated_at = CURRENT_TIMESTAMP WHERE setting_key = ?""",
                        (setting_value, session['user_id'], setting_key))
            
            conn.commit()
            flash('Settings updated successfully', 'success')
            
        except Exception as e:
            flash(f'Error updating settings: {str(e)}', 'error')
    
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM settings ORDER BY setting_key")
    settings = c.fetchall()
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('settings.html', settings=settings, unread_count=unread_count)

@app.route('/admin/audit-log')
@admin_required
def audit_log():
    """View audit log"""
    page = request.args.get('page', 1, type=int)
    per_page = app.config['ITEMS_PER_PAGE']
    offset = (page - 1) * per_page
    
    conn = get_db()
    c = conn.cursor()
    
    c.execute(
        """SELECT a.*, u.username FROM audit_log a LEFT JOIN users u ON a.user_id = u.id ORDER BY a.created_at DESC LIMIT ? OFFSET ?""",
        (per_page, offset))
    logs = c.fetchall()
    
    c.execute('SELECT COUNT(*) as count FROM audit_log')
    total = c.fetchone()['count']
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    total_pages = (total + per_page - 1) // per_page
    
    return render_template('audit_log.html',
                         logs=logs,
                         page=page,
                         total_pages=total_pages,
                         unread_count=unread_count)

# ============================================================================
# NOTIFICATION ROUTES
# ============================================================================

@app.route('/notifications')
@login_required
def notifications():
    """View notifications"""
    conn = get_db()
    c = conn.cursor()
    
    c.execute(
        """SELECT * FROM notifications WHERE user_id = ? ORDER BY created_at DESC LIMIT 50""",
        (session['user_id'],))
    notifications = c.fetchall()
    
    conn.close()
    
    return render_template('notifications.html', notifications=notifications)

@app.route('/notifications/mark-read/<int:id>', methods=['POST'])
@login_required
def mark_notification_read(id):
    """Mark notification as read"""
    conn = get_db()
    c = conn.cursor()
    
    c.execute(
        """UPDATE notifications SET is_read = 1 WHERE id = ? AND user_id = ?""",
        (id, session['user_id']))
    
    conn.commit()
    conn.close()
    
    return jsonify({'success': True})

@app.route('/notifications/mark-all-read', methods=['POST'])
@login_required
def mark_all_notifications_read():
    """Mark all notifications as read"""
    conn = get_db()
    c = conn.cursor()
    
    c.execute(
        """UPDATE notifications SET is_read = 1 WHERE user_id = ? AND is_read = 0""",
        (session['user_id'],))
    
    conn.commit()
    conn.close()
    
    return jsonify({'success': True})

# ============================================================================
# PROFILE ROUTES
# ============================================================================

@app.route('/profile')
@login_required
def profile():
    """User profile page"""
    conn = get_db()
    c = conn.cursor()
    
    c.execute('SELECT * FROM users WHERE id = ?', (session['user_id'],))
    user = c.fetchone()
    
    if session['role'] == 'student':
        c.execute(
            """SELECT s.*, d.dept_name FROM students s LEFT JOIN departments d ON s.department_id = d.id WHERE s.user_id = ?""",
            (session['user_id'],))
        profile_data = c.fetchone()
    elif session['role'] == 'teacher':
        c.execute(
            """SELECT i.*, d.dept_name FROM instructors i LEFT JOIN departments d ON i.department_id = d.id WHERE i.user_id = ?""",
            (session['user_id'],))
        profile_data = c.fetchone()
    else:
        profile_data = None
    
    # Get unread notifications count
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0', 
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    
    conn.close()
    
    return render_template('profile.html', 
                         user=user,
                         profile_data=profile_data,
                         unread_count=unread_count)

@app.route('/profile/update', methods=['POST'])
@login_required
def update_profile():
    """Update user profile"""
    try:
        conn = get_db()
        c = conn.cursor()
        
        c.execute(
            """UPDATE users SET first_name = ?, last_name = ?, email = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?""",
            ( request.form.get('first_name'), request.form.get('last_name'), request.form.get('email'), session['user_id'] ))
        
        if session['role'] == 'student':
            c.execute(
                """UPDATE students SET phone = ?, address = ?, city = ?, state = ?, zip_code = ?, country = ?, updated_at = CURRENT_TIMESTAMP WHERE user_id = ?""",
                ( request.form.get('phone'), request.form.get('address'), request.form.get('city'), request.form.get('state'), request.form.get('zip_code'), request.form.get('country'), session['user_id'] ))
        elif session['role'] == 'teacher':
            c.execute(
                """UPDATE instructors SET phone = ?, office_hours = ?, updated_at = CURRENT_TIMESTAMP WHERE user_id = ?""",
                ( request.form.get('phone'), request.form.get('office_hours'), session['user_id'] ))
        
        conn.commit()
        conn.close()
        
        flash('Profile updated successfully', 'success')
        
    except Exception as e:
        flash(f'Error updating profile: {str(e)}', 'error')
    
    return redirect(url_for('profile'))

@app.route('/profile/change-password', methods=['POST'])
@login_required
def change_password():
    """Change user password"""
    current_password = request.form.get('current_password')
    new_password = request.form.get('new_password')
    confirm_password = request.form.get('confirm_password')
    
    if new_password != confirm_password:
        flash('New passwords do not match', 'error')
        return redirect(url_for('profile'))
    
    conn = get_db()
    c = conn.cursor()
    
    c.execute('SELECT password_hash FROM users WHERE id = ?', (session['user_id'],))
    user = c.fetchone()
    
    if not check_password_hash(user['password_hash'], current_password):
        flash('Current password is incorrect', 'error')
        conn.close()
        return redirect(url_for('profile'))
    
    new_password_hash = generate_password_hash(new_password)
    c.execute('UPDATE users SET password_hash = ? WHERE id = ?', 
              (new_password_hash, session['user_id']))
    
    conn.commit()
    conn.close()
    
    flash('Password changed successfully', 'success')
    return redirect(url_for('profile'))

# ============================================================================
# API ROUTES
# ============================================================================

@app.route('/api/stats/dashboard')
@login_required
def api_dashboard_stats():
    """API endpoint for dashboard statistics"""
    conn = get_db()
    c = conn.cursor()
    stats = {}
    
    if session['role'] == 'admin':
        c.execute("SELECT COUNT(*) as count FROM students")
        stats['students'] = c.fetchone()['count']
        
        c.execute("SELECT COUNT(*) as count FROM courses")
        stats['courses'] = c.fetchone()['count']
        
        c.execute("SELECT COUNT(*) as count FROM instructors")
        stats['instructors'] = c.fetchone()['count']
        
        c.execute("SELECT COUNT(*) as count FROM departments")
        stats['departments'] = c.fetchone()['count']
        
        c.execute("SELECT COUNT(*) as count FROM enrollments WHERE date(created_at) = date('now')")
        stats['today_enrollments'] = c.fetchone()['count']
    
    elif session['role'] == 'teacher':
        c.execute(
            """SELECT COUNT(*) as count FROM courses WHERE instructor_id = (SELECT id FROM instructors WHERE user_id = ?)""",
            (session['user_id'],))
        stats['my_courses'] = c.fetchone()['count']
        
        c.execute(
            """SELECT COUNT(DISTINCT student_id) as count FROM enrollments e JOIN courses c ON e.course_id = c.id WHERE c.instructor_id = (SELECT id FROM instructors WHERE user_id = ?)""",
            (session['user_id'],))
        stats['my_students'] = c.fetchone()['count']
    
    elif session['role'] == 'student':
        c.execute(
            "SELECT COUNT(*) as count FROM enrollments WHERE student_id = (SELECT id FROM students WHERE user_id = ?)",
            (session['user_id'],))
        stats['my_courses'] = c.fetchone()['count']
        
        c.execute(
            "SELECT COUNT(*) as count FROM attendance WHERE student_id = (SELECT id FROM students WHERE user_id = ?) AND status = 'present'",
            (session['user_id'],))
        stats['present_count'] = c.fetchone()['count']
    
    conn.close()
    return jsonify(stats)

@app.route('/api/students/search')
@admin_required
def api_students_search():
    """API endpoint for searching students"""
    query = request.args.get('q', '')
    
    conn = get_db()
    c = conn.cursor()
    
    c.execute(
        "SELECT id, student_id, first_name, last_name, email FROM students WHERE first_name LIKE ? OR last_name LIKE ? OR student_id LIKE ? OR email LIKE ? LIMIT 20",
        (f'%{query}%', f'%{query}%', f'%{query}%', f'%{query}%'))
    
    students = [dict(row) for row in c.fetchall()]
    conn.close()
    
    return jsonify(students)

@app.route('/api/courses/search')
@admin_required
def api_courses_search():
    """API endpoint for searching courses"""
    query = request.args.get('q', '')
    
    conn = get_db()
    c = conn.cursor()
    
    c.execute(
        "SELECT id, course_code, course_name, credits FROM courses WHERE (course_code LIKE ? OR course_name LIKE ?) AND status = 'active' LIMIT 20",
        (f'%{query}%', f'%{query}%'))
    
    courses = [dict(row) for row in c.fetchall()]
    conn.close()
    
    return jsonify(courses)

@app.route('/api/attendance/today')
@faculty_required
def api_today_attendance():
    """API endpoint for today's attendance"""
    course_id = request.args.get('course_id')
    date = request.args.get('date', datetime.now().strftime('%Y-%m-%d'))
    
    conn = get_db()
    c = conn.cursor()
    
    c.execute(
        """SELECT s.id, s.student_id, s.first_name, s.last_name, a.status
        FROM students s
        JOIN enrollments e ON s.id = e.student_id
        LEFT JOIN attendance a ON s.id = a.student_id AND a.course_id = ? AND a.date = ?
        WHERE e.course_id = ?
        ORDER BY s.last_name, s.first_name""",
        (course_id, date, course_id))
    
    attendance = [dict(row) for row in c.fetchall()]
    conn.close()
    
    return jsonify(attendance)

# ============================================================================
# ERROR HANDLERS
# ============================================================================

@app.errorhandler(404)
def not_found_error(error):
    """Handle 404 errors"""
    return render_template('404.html'), 404

@app.errorhandler(500)
def internal_error(error):
    """Handle 500 errors"""
    conn = get_db()
    c = conn.cursor()
    c.execute("ROLLBACK")
    conn.close()
    return render_template('500.html'), 500

@app.errorhandler(403)
def forbidden_error(error):
    """Handle 403 errors"""
    return render_template('403.html'), 403

# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

@app.route('/student/attendance-ai')
@student_required
def student_attendance_ai():
    """AI-powered attendance analysis page for students"""
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT id FROM students WHERE user_id = ?", (session['user_id'],))
    student = c.fetchone()

    if not student:
        conn.close()
        flash('Student profile not found', 'error')
        return redirect(url_for('dashboard'))

    c.execute("SELECT setting_value FROM settings WHERE setting_key = 'total_semester_classes'")
    row = c.fetchone()
    total_semester_classes = int(row['setting_value']) if row else 90

    c.execute(
        """SELECT c.id, c.course_code, c.course_name,
               i.first_name as instructor_first, i.last_name as instructor_last
        FROM courses c
        JOIN enrollments e ON c.id = e.course_id
        LEFT JOIN instructors i ON c.instructor_id = i.id
        WHERE e.student_id = ?
        ORDER BY c.course_code""",
        (student['id'],))
    courses = c.fetchall()

    course_analysis = []
    for course in courses:
        c.execute(
            """SELECT COUNT(CASE WHEN status IN ('present','late') THEN 1 END) as present_count,
                COUNT(*) as total_held
            FROM attendance
            WHERE student_id = ? AND course_id = ?""",
            (student['id'], course['id']))
        att = c.fetchone()
        present_count = att['present_count'] if att else 0
        total_held = att['total_held'] if att else 0

        analysis = analyse_student_course(present_count, total_held, total_semester_classes)
        course_analysis.append({
            'course_id':   course['id'],
            'course_code': course['course_code'],
            'course_name': course['course_name'],
            'instructor':  f"{course['instructor_first']} {course['instructor_last']}",
            **analysis
        })

    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']

    conn.close()

    return render_template('student_attendance_ai.html',
                           course_analysis=course_analysis,
                           unread_count=unread_count,
                           total_semester_classes=total_semester_classes)


@app.route('/student/ai-insights')
@student_required
def student_ai_insights():
    """AI Grade Predictor + Dropout Risk + Study Recommendations"""
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT id FROM students WHERE user_id = ?", (session['user_id'],))
    student = c.fetchone()

    if not student:
        conn.close()
        flash('Student profile not found', 'error')
        return redirect(url_for('dashboard'))

    c.execute("SELECT setting_value FROM settings WHERE setting_key = 'total_semester_classes'")
    row = c.fetchone()
    total_semester_classes = int(row['setting_value']) if row else 90

    c.execute("SELECT setting_value FROM settings WHERE setting_key = 'gemini_api_key'")
    api_key_row = c.fetchone()
    api_key = api_key_row['setting_value'] if api_key_row else None

    # Get all enrolled courses
    c.execute(
        """SELECT c.id, c.course_code, c.course_name, c.credits,
               i.first_name as instructor_first, i.last_name as instructor_last,
               e.grade, e.progress
        FROM courses c
        JOIN enrollments e ON c.id = e.course_id
        LEFT JOIN instructors i ON c.instructor_id = i.id
        WHERE e.student_id = ?
        ORDER BY c.course_code""",
        (student['id'],))
    courses = c.fetchall()

    ai_insights = []
    total_risk_score = 0
    courses_failing = 0

    for course in courses:
        # Get attendance
        c.execute(
            """SELECT COUNT(CASE WHEN status IN ('present','late') THEN 1 END) as present_count,
                COUNT(*) as total_held
            FROM attendance
            WHERE student_id = ? AND course_id = ?""",
            (student['id'], course['id']))
        att = c.fetchone()
        present_count = att['present_count'] if att else 0
        total_held = att['total_held'] if att else 0
        attendance_pct = (present_count / total_held * 100) if total_held > 0 else 0

        # Current score from progress or grade
        grade_map = {'A': 95, 'A-': 90, 'B+': 87, 'B': 83, 'B-': 80,
                     'C+': 77, 'C': 73, 'C-': 70, 'D+': 67, 'D': 63, 'F': 40}
        if course['grade']:
            current_score = grade_map.get(course['grade'], 50)
        else:
            current_score = course['progress'] if course['progress'] else 50

        if current_score < 60:
            courses_failing += 1

        # Grade prediction
        grade_pred = predict_final_grade(
            attendance_pct=attendance_pct,
            current_score=current_score,
            assignments_submitted=present_count,
            total_assignments=total_held
        )

        # Dropout risk per course
        dropout = calculate_dropout_risk(
            attendance_pct=attendance_pct,
            grade_score=current_score,
            assignments_rate=grade_pred['assignment_rate'],
            days_since_last_login=0,
            courses_failing=courses_failing
        )

        # Study recommendations using Gemini (if key exists)
        recommendations = generate_study_recommendations(
            attendance_pct=attendance_pct,
            grade_score=current_score,
            assignment_rate=grade_pred['assignment_rate'],
            course_name=course['course_name'],
            api_key=api_key
        )
        total_risk_score += dropout['risk_score']

        ai_insights.append({
            'course_id':   course['id'],
            'course_code': course['course_code'],
            'course_name': course['course_name'],
            'instructor':  f"{course['instructor_first']} {course['instructor_last']}",
            'attendance_pct': round(attendance_pct, 1),
            'current_score':  current_score,
            'grade_prediction': grade_pred,
            'dropout_risk':    dropout,
            'recommendations': recommendations,
        })

    # Overall risk
    avg_risk = round(total_risk_score / len(ai_insights), 1) if ai_insights else 0

    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']

    conn.close()

    return render_template('student_ai_insights.html',
                           ai_insights=ai_insights,
                           avg_risk=avg_risk,
                           unread_count=unread_count,
                           total_semester_classes=total_semester_classes)



# ============================================================================
# CLASSROOM ROUTES
# ============================================================================

@app.route('/classroom/teacher/create', methods=['GET', 'POST'])
@faculty_required
def classroom_create():
    """Teacher creates a new classroom session"""
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT id FROM instructors WHERE user_id = ?", (session['user_id'],))
    instructor = c.fetchone()

    if not instructor:
        conn.close()
        flash('Instructor profile not found', 'error')
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        course_id = request.form.get('course_id')
        session_title = request.form.get('session_title', 'Live Class')

        # Get course questions
        c.execute("SELECT setting_value FROM settings WHERE setting_key = 'gemini_api_key'")
        api_key_row = c.fetchone()
        api_key = api_key_row['setting_value'] if api_key_row else None
        
        c.execute("SELECT course_name, course_code FROM courses WHERE id = ?", (course_id,))
        course = c.fetchone()
        questions = get_questions_for_course(course['course_name'], course['course_code'], count=8, api_key=api_key)

        c.execute(
            "INSERT INTO classroom_sessions (course_id, instructor_id, session_title, status, questions_data) VALUES (?, ?, ?, ?, ?)",
            (course_id, instructor['id'], session_title, 'waiting', json.dumps(questions)))

        session_id = c.lastrowid

        # Notify all enrolled students
        c.execute(
            "SELECT s.id, u.id as user_id, s.first_name, s.last_name FROM students s JOIN users u ON s.user_id = u.id JOIN enrollments e ON s.id = e.student_id WHERE e.course_id = ?",
            (course_id,))
        students = c.fetchall()

        for stu in students:
            join_url = f"/classroom/respond/{session_id}/join"
            decline_url = f"/classroom/respond/{session_id}/decline"
            c.execute(
                "INSERT INTO notifications (user_id, title, message, type, action_url) VALUES (?, ?, ?, ?, ?)",
                (stu['user_id'],
                 f"🔴 Live Class Starting — {session_title}",
                 f"Your teacher has started a live class for {course['course_name']}. Click JOIN to enter or DECLINE if you cannot attend. [JOIN]({join_url}) [DECLINE]({decline_url})",
                 'warning',
                 join_url))

        conn.commit()
        conn.close()

        flash('Classroom created! Share the link with students.', 'success')
        return redirect(url_for('classroom_teacher', session_id=session_id))

    # Get teacher courses
    c.execute(
        "SELECT c.id, c.course_code, c.course_name FROM courses c WHERE c.instructor_id = ? ORDER BY c.course_code",
        (instructor['id'],))
    courses = c.fetchall()

    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']

    conn.close()
    return render_template('classroom_create.html', courses=courses, unread_count=unread_count)


@app.route('/classroom/teacher/<int:session_id>')
@faculty_required
def classroom_teacher(session_id):
    """Teacher classroom view"""
    conn = get_db()
    c = conn.cursor()

    c.execute(
        "SELECT cs.*, c.course_name, c.course_code FROM classroom_sessions cs JOIN courses c ON cs.course_id = c.id WHERE cs.id = ?",
        (session_id,))
    classroom = c.fetchone()

    if not classroom:
        conn.close()
        flash('Session not found', 'error')
        return redirect(url_for('teacher_dashboard'))

    questions = json.loads(classroom['questions_data']) if classroom['questions_data'] else []

    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']

    conn.close()
    return render_template('classroom_teacher.html',
                           classroom=classroom,
                           questions=questions,
                           unread_count=unread_count,
                           session_id=session_id)


@app.route('/classroom/student/<int:session_id>')
@student_required
def classroom_student(session_id):
    """Student classroom view"""
    conn = get_db()
    c = conn.cursor()

    c.execute(
        "SELECT cs.*, c.course_name, c.course_code FROM classroom_sessions cs JOIN courses c ON cs.course_id = c.id WHERE cs.id = ?",
        (session_id,))
    classroom = c.fetchone()

    if not classroom:
        conn.close()
        flash('Session not found', 'error')
        return redirect(url_for('student_dashboard'))

    c.execute("SELECT id FROM students WHERE user_id = ?", (session['user_id'],))
    student = c.fetchone()

    questions = json.loads(classroom['questions_data']) if classroom['questions_data'] else []

    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']

    conn.close()
    return render_template('classroom_student.html',
                           classroom=classroom,
                           questions=questions,
                           student_id=student['id'] if student else None,
                           unread_count=unread_count,
                           session_id=session_id)


@app.route('/classroom/sessions')
@faculty_required
def classroom_sessions():
    """List all classroom sessions for teacher"""
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT id FROM instructors WHERE user_id = ?", (session['user_id'],))
    instructor = c.fetchone()

    c.execute(
        "SELECT cs.*, c.course_name, c.course_code, COUNT(DISTINCT ca.student_id) as student_count FROM classroom_sessions cs JOIN courses c ON cs.course_id = c.id LEFT JOIN classroom_attendance ca ON cs.id = ca.session_id WHERE cs.instructor_id = ? GROUP BY cs.id ORDER BY cs.created_at DESC",
        (instructor['id'],))
    sessions_list = c.fetchall()

    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']

    conn.close()
    return render_template('classroom_sessions.html',
                           sessions=sessions_list,
                           unread_count=unread_count)


@app.route('/classroom/report/<int:session_id>')
@faculty_required
def classroom_report(session_id):
    """Detailed report for a classroom session"""
    conn = get_db()
    c = conn.cursor()

    c.execute(
        "SELECT cs.*, c.course_name, c.course_code FROM classroom_sessions cs JOIN courses c ON cs.course_id = c.id WHERE cs.id = ?",
        (session_id,))
    classroom = c.fetchone()

    # Get all student responses
    c.execute(
        """SELECT s.first_name, s.last_name, s.student_id,
               COUNT(pr.id) as questions_answered,
               SUM(CASE WHEN pr.is_correct THEN 1 ELSE 0 END) as correct_answers,
               AVG(pr.time_taken_seconds) as avg_time,
               ca.status, ca.was_kicked
        FROM classroom_attendance ca
        JOIN students s ON ca.student_id = s.id
        LEFT JOIN proctor_responses pr ON ca.student_id = pr.student_id
            AND pr.session_id = ?
        WHERE ca.session_id = ?
        GROUP BY s.id
        ORDER BY s.last_name
        """, (session_id, session_id))
    student_reports = c.fetchall()

    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']

    conn.close()
    return render_template('classroom_report.html',
                           classroom=classroom,
                           student_reports=student_reports,
                           unread_count=unread_count)


# ============================================================================
# SOCKET.IO EVENTS
# ============================================================================

@socketio.on('join_classroom')
def on_join_classroom(data):
    """Student or teacher joins classroom room"""
    session_id = data.get('session_id')
    user_type = data.get('user_type')
    student_id = data.get('student_id')
    room = f"classroom_{session_id}"
    join_room(room)

    if user_type == 'student' and student_id:
        # Auto mark attendance
        conn = get_db()
        c = conn.cursor()
        try:
            c.execute(
                "INSERT OR IGNORE INTO classroom_attendance (session_id, student_id, joined_at, status) VALUES (?, ?, CURRENT_TIMESTAMP, 'present')",
                (session_id, student_id))
            conn.commit()
        except:
            pass
        conn.close()

        # Notify teacher
        emit('student_joined', {
            'student_id': student_id,
            'name': data.get('student_name', 'Student'),
            'time': datetime.now().strftime('%H:%M:%S')
        }, room=room)


@socketio.on('start_class')
def on_start_class(data):
    """Teacher starts the class"""
    session_id = data.get('session_id')
    room = f"classroom_{session_id}"

    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE classroom_sessions SET status='active', started_at=CURRENT_TIMESTAMP WHERE id=?",
              (session_id,))
    conn.commit()
    conn.close()

    emit('class_started', {'session_id': session_id}, room=room)


@socketio.on('end_class')
def on_end_class(data):
    """Teacher ends the class"""
    session_id = data.get('session_id')
    notes = data.get('notes', '')
    room = f"classroom_{session_id}"

    conn = get_db()
    c = conn.cursor()
    c.execute(
        "UPDATE classroom_sessions SET status='ended', ended_at=CURRENT_TIMESTAMP, notes=? WHERE id=?",
        (notes, session_id))
    conn.commit()
    conn.close()

    emit('class_ended', {'notes': notes}, room=room)


@socketio.on('send_question')
def on_send_question(data):
    """Teacher sends proctor question to all students"""
    session_id = data.get('session_id')
    question = data.get('question')
    room = f"classroom_{session_id}"
    emit('receive_question', {'question': question}, room=room)


@socketio.on('submit_answer')
def on_submit_answer(data):
    """Student submits answer to proctor question"""
    session_id = data.get('session_id')
    student_id = data.get('student_id')
    student_name = data.get('student_name')
    question_id = data.get('question_id')
    question_text = data.get('question_text')
    answer_given = data.get('answer_given')
    correct_answer = data.get('correct_answer')
    time_taken = data.get('time_taken', 0)
    is_correct = (int(answer_given) == int(correct_answer))
    room = f"classroom_{session_id}"

    conn = get_db()
    c = conn.cursor()
    c.execute(
        "INSERT INTO proctor_responses (session_id, student_id, question_id, question_text, answer_given, correct_answer, is_correct, time_taken_seconds) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (session_id, student_id, question_id, question_text, answer_given, correct_answer, is_correct, time_taken))
    conn.commit()
    conn.close()

    # Notify teacher of answer
    emit('student_answered', {
        'student_id': student_id,
        'student_name': student_name,
        'question_id': question_id,
        'is_correct': is_correct,
        'time_taken': time_taken,
        'answer_given': answer_given
    }, room=room)


@socketio.on('missed_question')
def on_missed_question(data):
    """Student missed a question — notify teacher"""
    session_id = data.get('session_id')
    student_id = data.get('student_id')
    student_name = data.get('student_name')
    missed_count = data.get('missed_count', 0)
    room = f"classroom_{session_id}"

    # Save missed as no answer
    conn = get_db()
    c = conn.cursor()
    c.execute(
        "INSERT INTO proctor_responses (session_id, student_id, question_id, question_text, answer_given, correct_answer, is_correct, time_taken_seconds) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (session_id, student_id, data.get('question_id', 0), 'MISSED', -1, 0, False, 0))
    conn.commit()
    conn.close()

    if missed_count >= 3:
        # Auto kick
        emit('kick_student', {
            'student_id': student_id,
            'student_name': student_name,
            'reason': 'Missed 3 or more questions'
        }, room=room)

        # Update attendance
        conn = get_db()
        c = conn.cursor()
        c.execute(
            "UPDATE classroom_attendance SET status='kicked', was_kicked=1, left_at=CURRENT_TIMESTAMP WHERE session_id=? AND student_id=?",
            (session_id, student_id))
        conn.commit()
        conn.close()
    elif missed_count >= 2:
        # Alert teacher
        emit('alert_teacher', {
            'student_id': student_id,
            'student_name': student_name,
            'missed_count': missed_count,
            'message': f"{student_name} has missed {missed_count} questions!"
        }, room=room)


@socketio.on('teacher_modify_attendance')
def on_modify_attendance(data):
    """Teacher manually modifies student attendance"""
    session_id = data.get('session_id')
    student_id = data.get('student_id')
    new_status = data.get('status')

    conn = get_db()
    c = conn.cursor()
    c.execute(
        "UPDATE classroom_attendance SET status=? WHERE session_id=? AND student_id=?",
        (new_status, session_id, student_id))
    conn.commit()
    conn.close()

    emit('attendance_updated', {
        'student_id': student_id,
        'status': new_status
    }, room=f"classroom_{session_id}")


@socketio.on('offer')
def on_offer(data):
    """WebRTC offer from teacher"""
    room = f"classroom_{data.get('session_id')}"
    emit('offer', data, room=room, include_self=False)


@socketio.on('answer')
def on_answer(data):
    """WebRTC answer from student"""
    room = f"classroom_{data.get('session_id')}"
    emit('answer', data, room=room, include_self=False)


@socketio.on('ice_candidate')
def on_ice_candidate(data):
    """WebRTC ICE candidate"""
    room = f"classroom_{data.get('session_id')}"
    emit('ice_candidate', data, room=room, include_self=False)



# ============================================================================
# EXPORT ROUTES
# ============================================================================

def make_excel_response(wb, filename):
    """Helper to return Excel file as download"""
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    response = make_response(output.read())
    response.headers['Content-Disposition'] = f'attachment; filename={filename}'
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return response


def style_header_row(ws, row, cols):
    """Apply header styling to Excel row"""
    header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    for col in range(1, cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')


@app.route('/admin/export/students')
@admin_required
def export_students():
    """Export all students to Excel"""
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT s.student_id, s.first_name, s.last_name, s.email, s.phone,
               s.gender, s.date_of_birth, s.enrollment_date, s.status,
               d.dept_name, s.city, s.state, s.country,
               s.guardian_name, s.guardian_phone
        FROM students s
        LEFT JOIN departments d ON s.department_id = d.id
        ORDER BY s.last_name, s.first_name
    """)
    students = c.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Students"

    # Title
    ws.merge_cells('A1:O1')
    ws['A1'] = 'STUDENT REPORT'
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="0f3460", end_color="0f3460", fill_type="solid")
    ws['A1'].alignment = Alignment(horizontal='center')
    ws.row_dimensions[1].height = 30

    headers = ['Student ID', 'First Name', 'Last Name', 'Email', 'Phone',
               'Gender', 'Date of Birth', 'Enrollment Date', 'Status',
               'Department', 'City', 'State', 'Country',
               'Guardian Name', 'Guardian Phone']
    for col, header in enumerate(headers, 1):
        ws.cell(row=2, column=col, value=header)
    style_header_row(ws, 2, len(headers))

    for row_idx, stu in enumerate(students, 3):
        for col_idx, value in enumerate(stu, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if row_idx % 2 == 0:
                cell.fill = PatternFill(start_color="f0f4ff", end_color="f0f4ff", fill_type="solid")

    for col in ws.columns:
        try:
            col_letter = col[0].column_letter
            max_len = max((len(str(cell.value or '')) for cell in col if hasattr(cell, 'value')), default=10)
            ws.column_dimensions[col_letter].width = min(max_len + 4, 30)
        except AttributeError:
            pass

    return make_excel_response(wb, 'students_report.xlsx')


@app.route('/admin/export/attendance')
@admin_required
def export_attendance():
    """Export attendance report to Excel"""
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT s.student_id, s.first_name || ' ' || s.last_name as name,
               c.course_code, c.course_name,
               COUNT(*) as total_classes,
               SUM(CASE WHEN a.status IN ('present','late') THEN 1 ELSE 0 END) as attended,
               ROUND(SUM(CASE WHEN a.status IN ('present','late') THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) as percentage,
               SUM(CASE WHEN a.status = 'absent' THEN 1 ELSE 0 END) as absents
        FROM attendance a
        JOIN students s ON a.student_id = s.id
        JOIN courses c ON a.course_id = c.id
        GROUP BY s.id, c.id
        ORDER BY s.last_name, c.course_code
    """)
    records = c.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Attendance Report"

    ws.merge_cells('A1:H1')
    ws['A1'] = 'ATTENDANCE REPORT'
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="198754", end_color="198754", fill_type="solid")
    ws['A1'].alignment = Alignment(horizontal='center')
    ws.row_dimensions[1].height = 30

    headers = ['Student ID', 'Student Name', 'Course Code', 'Course Name',
               'Total Classes', 'Attended', 'Percentage %', 'Absents']
    for col, header in enumerate(headers, 1):
        ws.cell(row=2, column=col, value=header)
    style_header_row(ws, 2, len(headers))

    for row_idx, rec in enumerate(records, 3):
        for col_idx, value in enumerate(rec, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if row_idx % 2 == 0:
                cell.fill = PatternFill(start_color="f0fff4", end_color="f0fff4", fill_type="solid")
            # Color code attendance percentage
            if col_idx == 7 and value is not None:
                if float(value) < 60:
                    cell.font = Font(color="dc3545", bold=True)
                elif float(value) < 80:
                    cell.font = Font(color="fd7e14", bold=True)
                else:
                    cell.font = Font(color="198754", bold=True)

    for col in ws.columns:
        try:
            col_letter = col[0].column_letter
            max_len = max((len(str(cell.value or '')) for cell in col if hasattr(cell, 'value')), default=10)
            ws.column_dimensions[col_letter].width = min(max_len + 4, 30)
        except AttributeError:
            pass

    return make_excel_response(wb, 'attendance_report.xlsx')


@app.route('/admin/export/courses')
@admin_required
def export_courses():
    """Export courses to Excel"""
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT c.course_code, c.course_name, c.credits, c.semester,
               c.academic_year, d.dept_name,
               i.first_name || ' ' || i.last_name as instructor,
               c.max_students, c.enrolled_students, c.status,
               c.schedule, c.classroom
        FROM courses c
        LEFT JOIN departments d ON c.department_id = d.id
        LEFT JOIN instructors i ON c.instructor_id = i.id
        ORDER BY c.course_code
    """)
    courses = c.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Courses"

    ws.merge_cells('A1:L1')
    ws['A1'] = 'COURSES REPORT'
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="0f3460", end_color="0f3460", fill_type="solid")
    ws['A1'].alignment = Alignment(horizontal='center')
    ws.row_dimensions[1].height = 30

    headers = ['Course Code', 'Course Name', 'Credits', 'Semester',
               'Academic Year', 'Department', 'Instructor',
               'Max Students', 'Enrolled', 'Status', 'Schedule', 'Classroom']
    for col, header in enumerate(headers, 1):
        ws.cell(row=2, column=col, value=header)
    style_header_row(ws, 2, len(headers))

    for row_idx, course in enumerate(courses, 3):
        for col_idx, value in enumerate(course, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if row_idx % 2 == 0:
                cell.fill = PatternFill(start_color="f0f4ff", end_color="f0f4ff", fill_type="solid")

    for col in ws.columns:
        try:
            col_letter = col[0].column_letter
            max_len = max((len(str(cell.value or '')) for cell in col if hasattr(cell, 'value')), default=10)
            ws.column_dimensions[col_letter].width = min(max_len + 4, 30)
        except AttributeError:
            pass

    return make_excel_response(wb, 'courses_report.xlsx')


@app.route('/admin/export/full-report')
@admin_required
def export_full_report():
    """Export complete school report with multiple sheets"""
    conn = get_db()
    c = conn.cursor()
    wb = openpyxl.Workbook()

    # ── Sheet 1: Summary ──────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Summary"
    ws1.merge_cells('A1:D1')
    ws1['A1'] = 'SCHOOL MANAGEMENT SYSTEM — FULL REPORT'
    ws1['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws1['A1'].fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
    ws1['A1'].alignment = Alignment(horizontal='center')

    c.execute("SELECT COUNT(*) as count FROM students WHERE status='active'")
    total_students = c.fetchone()['count']
    c.execute("SELECT COUNT(*) as count FROM courses WHERE status='active'")
    total_courses = c.fetchone()['count']
    c.execute("SELECT COUNT(*) as count FROM instructors WHERE status='active'")
    total_instructors = c.fetchone()['count']
    c.execute("SELECT COUNT(*) as count FROM departments")
    total_depts = c.fetchone()['count']
    c.execute("SELECT COUNT(*) as count FROM enrollments")
    total_enrollments = c.fetchone()['count']

    summary_data = [
        ('Metric', 'Value'),
        ('Total Active Students', total_students),
        ('Total Active Courses', total_courses),
        ('Total Instructors', total_instructors),
        ('Total Departments', total_depts),
        ('Total Enrollments', total_enrollments),
        ('Report Generated', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
    ]
    for row_idx, (key, val) in enumerate(summary_data, 2):
        ws1.cell(row=row_idx, column=1, value=key).font = Font(bold=True)
        ws1.cell(row=row_idx, column=2, value=val)
    ws1.column_dimensions['A'].width = 30
    ws1.column_dimensions['B'].width = 25

    # ── Sheet 2: Students ─────────────────────────────────────────
    ws2 = wb.create_sheet("Students")
    c.execute("""
        SELECT s.student_id, s.first_name, s.last_name, s.email,
               s.enrollment_date, s.status, d.dept_name
        FROM students s LEFT JOIN departments d ON s.department_id = d.id
        ORDER BY s.last_name
    """)
    students = c.fetchall()
    headers2 = ['Student ID', 'First Name', 'Last Name', 'Email', 'Enrollment Date', 'Status', 'Department']
    for col, h in enumerate(headers2, 1):
        ws2.cell(row=1, column=col, value=h)
    style_header_row(ws2, 1, len(headers2))
    for ri, row in enumerate(students, 2):
        for ci, val in enumerate(row, 1):
            ws2.cell(row=ri, column=ci, value=val)

    # ── Sheet 3: Attendance Summary ───────────────────────────────
    ws3 = wb.create_sheet("Attendance")
    c.execute("""
        SELECT s.student_id, s.first_name || ' ' || s.last_name as name,
               c.course_code,
               ROUND(SUM(CASE WHEN a.status IN ('present','late') THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) as pct
        FROM attendance a
        JOIN students s ON a.student_id = s.id
        JOIN courses c ON a.course_id = c.id
        GROUP BY s.id, c.id ORDER BY s.last_name
    """)
    att_records = c.fetchall()
    headers3 = ['Student ID', 'Name', 'Course', 'Attendance %']
    for col, h in enumerate(headers3, 1):
        ws3.cell(row=1, column=col, value=h)
    style_header_row(ws3, 1, len(headers3))
    for ri, row in enumerate(att_records, 2):
        for ci, val in enumerate(row, 1):
            cell = ws3.cell(row=ri, column=ci, value=val)
            if ci == 4 and val is not None:
                if float(val) < 60: cell.font = Font(color="dc3545", bold=True)
                elif float(val) < 80: cell.font = Font(color="fd7e14", bold=True)
                else: cell.font = Font(color="198754", bold=True)

    # ── Sheet 4: Courses ──────────────────────────────────────────
    ws4 = wb.create_sheet("Courses")
    c.execute("""
        SELECT c.course_code, c.course_name, c.credits, d.dept_name,
               i.first_name || ' ' || i.last_name, c.status
        FROM courses c
        LEFT JOIN departments d ON c.department_id = d.id
        LEFT JOIN instructors i ON c.instructor_id = i.id
    """)
    courses = c.fetchall()
    headers4 = ['Code', 'Name', 'Credits', 'Department', 'Instructor', 'Status']
    for col, h in enumerate(headers4, 1):
        ws4.cell(row=1, column=col, value=h)
    style_header_row(ws4, 1, len(headers4))
    for ri, row in enumerate(courses, 2):
        for ci, val in enumerate(row, 1):
            ws4.cell(row=ri, column=ci, value=val)

    # Auto-fit all sheets
    for ws in [ws2, ws3, ws4]:
        for col in ws.columns:
            try:
                col_letter = col[0].column_letter
                max_len = max((len(str(cell.value or '')) for cell in col if hasattr(cell, 'value')), default=10)
                ws.column_dimensions[col_letter].width = min(max_len + 4, 35)
            except AttributeError:
                pass

    conn.close()
    return make_excel_response(wb, 'full_school_report.xlsx')


# ============================================================================
# ADVANCED SEARCH ROUTES
# ============================================================================

@app.route('/admin/search')
@admin_required
def advanced_search():
    """Advanced search across all data"""
    query = request.args.get('q', '').strip()
    search_type = request.args.get('type', 'all')
    department_id = request.args.get('department_id', '')
    status = request.args.get('status', '')
    results = {'students': [], 'courses': [], 'instructors': []}

    if query or department_id or status:
        conn = get_db()
        c = conn.cursor()

        # Build dynamic filters
        filters = []
        params = []

        if query:
            filters.append("(s.first_name LIKE ? OR s.last_name LIKE ? OR s.student_id LIKE ? OR s.email LIKE ?)")
            params.extend([f'%{query}%'] * 4)
        if department_id:
            filters.append("s.department_id = ?")
            params.append(department_id)
        if status:
            filters.append("s.status = ?")
            params.append(status)

        where = "WHERE " + " AND ".join(filters) if filters else ""

        if search_type in ['all', 'students']:
            c.execute(f"""
                SELECT s.*, d.dept_name
                FROM students s
                LEFT JOIN departments d ON s.department_id = d.id
                {where}
                ORDER BY s.last_name LIMIT 50
            """, params)
            results['students'] = c.fetchall()

        if search_type in ['all', 'courses']:
            course_filters = []
            course_params = []
            if query:
                course_filters.append("(c.course_code LIKE ? OR c.course_name LIKE ?)")
                course_params.extend([f'%{query}%'] * 2)
            if status:
                course_filters.append("c.status = ?")
                course_params.append(status)
            course_where = "WHERE " + " AND ".join(course_filters) if course_filters else ""
            c.execute(f"""
                SELECT c.*, d.dept_name, i.first_name || ' ' || i.last_name as instructor_name
                FROM courses c
                LEFT JOIN departments d ON c.department_id = d.id
                LEFT JOIN instructors i ON c.instructor_id = i.id
                {course_where}
                ORDER BY c.course_code LIMIT 50
            """, course_params)
            results['courses'] = c.fetchall()

        if search_type in ['all', 'instructors']:
            inst_filters = []
            inst_params = []
            if query:
                inst_filters.append("(i.first_name LIKE ? OR i.last_name LIKE ? OR i.email LIKE ?)")
                inst_params.extend([f'%{query}%'] * 3)
            if department_id:
                inst_filters.append("i.department_id = ?")
                inst_params.append(department_id)
            inst_where = "WHERE " + " AND ".join(inst_filters) if inst_filters else ""
            c.execute(f"""
                SELECT i.*, d.dept_name
                FROM instructors i
                LEFT JOIN departments d ON i.department_id = d.id
                {inst_where}
                ORDER BY i.last_name LIMIT 50
            """, inst_params)
            results['instructors'] = c.fetchall()

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM departments ORDER BY dept_name")
    departments = c.fetchall()
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    conn.close()

    return render_template('admin_search.html',
                           results=results,
                           query=query,
                           search_type=search_type,
                           departments=departments,
                           department_id=department_id,
                           status=status,
                           unread_count=unread_count)


# ============================================================================
# AI ANALYTICS DASHBOARD FOR ADMIN
# ============================================================================

@app.route('/admin/ai-analytics')
@admin_required
def admin_ai_analytics():
    """School-wide AI analytics dashboard"""
    conn = get_db()
    c = conn.cursor()

    # Get total semester classes setting
    c.execute("SELECT setting_value FROM settings WHERE setting_key = 'total_semester_classes'")
    row = c.fetchone()
    total_semester_classes = int(row['setting_value']) if row else 90

    # Overall stats
    c.execute("SELECT COUNT(*) as count FROM students WHERE status='active'")
    total_students = c.fetchone()['count']
    c.execute("SELECT COUNT(*) as count FROM courses WHERE status='active'")
    total_courses = c.fetchone()['count']
    c.execute("SELECT COUNT(*) as count FROM enrollments")
    total_enrollments = c.fetchone()['count']

    # Get all students with their attendance
    c.execute("""
        SELECT s.id, s.student_id, s.first_name, s.last_name,
               d.dept_name
        FROM students s
        LEFT JOIN departments d ON s.department_id = d.id
        WHERE s.status = 'active'
    """)
    all_students = c.fetchall()

    # Analyze each student
    at_risk_count = 0
    critical_count = 0
    safe_count = 0
    dept_risk = {}
    at_risk_students = []

    for stu in all_students:
        # Get all courses for this student
        c.execute("""
            SELECT course_id FROM enrollments WHERE student_id = ?
        """, (stu['id'],))
        courses = c.fetchall()

        worst_badge = 'SAFE'
        for course in courses:
            c.execute("""
                SELECT COUNT(CASE WHEN status IN ('present','late') THEN 1 END) as present_count,
                       COUNT(*) as total_held
                FROM attendance WHERE student_id = ? AND course_id = ?
            """, (stu['id'], course['course_id']))
            att = c.fetchone()
            present = att['present_count'] if att else 0
            total = att['total_held'] if att else 0
            analysis = analyse_student_course(present, total, total_semester_classes)
            if analysis['risk_badge'] == 'CRITICAL':
                worst_badge = 'CRITICAL'
            elif analysis['risk_badge'] == 'WARNING' and worst_badge != 'CRITICAL':
                worst_badge = 'WARNING'

        dept = stu['dept_name'] or 'Unassigned'
        if dept not in dept_risk:
            dept_risk[dept] = {'safe': 0, 'warning': 0, 'critical': 0, 'total': 0}
        dept_risk[dept]['total'] += 1

        if worst_badge == 'CRITICAL':
            critical_count += 1
            at_risk_count += 1
            dept_risk[dept]['critical'] += 1
            at_risk_students.append({
                'name': f"{stu['first_name']} {stu['last_name']}",
                'student_id': stu['student_id'],
                'dept': dept,
                'badge': 'CRITICAL',
                'color': '#dc3545'
            })
        elif worst_badge == 'WARNING':
            at_risk_count += 1
            dept_risk[dept]['warning'] += 1
            at_risk_students.append({
                'name': f"{stu['first_name']} {stu['last_name']}",
                'student_id': stu['student_id'],
                'dept': dept,
                'badge': 'WARNING',
                'color': '#fd7e14'
            })
        else:
            safe_count += 1
            dept_risk[dept]['safe'] += 1

    # Top performing departments by attendance
    c.execute("""
        SELECT d.dept_name,
               ROUND(AVG(CASE WHEN a.status IN ('present','late') THEN 1.0 ELSE 0.0 END) * 100, 1) as avg_attendance
        FROM attendance a
        JOIN students s ON a.student_id = s.id
        JOIN departments d ON s.department_id = d.id
        GROUP BY d.id
        ORDER BY avg_attendance DESC
    """)
    dept_attendance = c.fetchall()

    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']

    conn.close()

    return render_template('admin_ai_analytics.html',
                           total_students=total_students,
                           total_courses=total_courses,
                           total_enrollments=total_enrollments,
                           at_risk_count=at_risk_count,
                           critical_count=critical_count,
                           safe_count=safe_count,
                           dept_risk=dept_risk,
                           at_risk_students=at_risk_students,
                           dept_attendance=dept_attendance,
                           unread_count=unread_count)



# ============================================================================
# CLASSROOM ROUTES
# ============================================================================
# ============================================================================
# SOCKET.IO EVENTS

# ============================================================================
# AI CHATBOT ROUTE
# ============================================================================

@app.route('/chatbot', methods=['POST'])
@login_required
def chatbot():
    """AI Chatbot - answers questions based on user role and data using Gemini"""
    user_message = request.json.get('message', '').strip()
    role = session.get('role')
    user_id = session.get('user_id')
    first_name = session.get('first_name', 'there')

    conn = get_db()
    c = conn.cursor()
    
    # Check for Gemini API Key
    c.execute("SELECT setting_value FROM settings WHERE setting_key = 'gemini_api_key'")
    api_key_row = c.fetchone()
    api_key = api_key_row['setting_value'] if api_key_row else None
    
    # Gather exact backend context for the prompt
    context_data = {'role': role, 'name': first_name}

    try:
        if role == 'student':
            c.execute("SELECT id FROM students WHERE user_id = ?", (user_id,))
            student = c.fetchone()
            if student:
                sid = student['id']
                # Get Attendance
                c.execute("""SELECT c.course_code, COUNT(*) as total, SUM(CASE WHEN a.status IN ('present','late') THEN 1 ELSE 0 END) as attended
                           FROM attendance a JOIN courses c ON a.course_id = c.id WHERE a.student_id = ? GROUP BY c.id""", (sid,))
                context_data['attendance'] = [dict(r) for r in c.fetchall()]
                
                # Get GPA
                c.execute("""SELECT AVG(CASE grade WHEN 'A' THEN 4.0 WHEN 'A-' THEN 3.7 WHEN 'B+' THEN 3.3 WHEN 'B' THEN 3.0 WHEN 'B-' THEN 2.7 WHEN 'C+' THEN 2.3 WHEN 'C' THEN 2.0 WHEN 'C-' THEN 1.7 WHEN 'D+' THEN 1.3 WHEN 'D' THEN 1.0 ELSE 0.0 END) as gpa FROM enrollments WHERE student_id = ? AND grade IS NOT NULL""", (sid,))
                g_row = c.fetchone()
                context_data['gpa'] = round(g_row['gpa'], 2) if g_row and g_row['gpa'] else 0.0
                
                # Get Courses & Grades
                c.execute("""SELECT c.course_code, c.course_name, e.grade FROM enrollments e JOIN courses c ON e.course_id = c.id WHERE e.student_id = ?""", (sid,))
                context_data['enrolled_courses'] = [dict(r) for r in c.fetchall()]
    
        elif role == 'teacher':
            c.execute("SELECT id FROM instructors WHERE user_id = ?", (user_id,))
            inst = c.fetchone()
            if inst:
                iid = inst['id']
                # Get Courses Taught
                c.execute("""SELECT c.course_code, c.course_name, COUNT(DISTINCT e.student_id) as students FROM courses c LEFT JOIN enrollments e ON c.id = e.course_id WHERE c.instructor_id = ? GROUP BY c.id""", (iid,))
                context_data['courses_taught'] = [dict(r) for r in c.fetchall()]
                
        elif role == 'admin':
            c.execute("SELECT COUNT(*) as students FROM students WHERE status='active'")
            context_data['active_students'] = c.fetchone()['students']
            c.execute("SELECT COUNT(*) as courses FROM courses WHERE status='active'")
            context_data['active_courses'] = c.fetchone()['courses']
            c.execute("SELECT COUNT(*) as instructors FROM instructors")
            context_data['total_instructors'] = c.fetchone()['instructors']
    except Exception as e:
        print(f"[WARN] Chatbot context gathering error: {e}")
        
    conn.close()

    # Generate Response using Gemini
    response_text = ""
    if api_key:
        try:
            import google.generativeai as genai
            import json
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-2.5-flash')
            
            prompt = f"""You are EduBot, the friendly, helpful AI academic assistant for a school system.
You are talking directly to a {role} named {first_name}.
Here is their actual, real-time context data extracted from the application database: 
{json.dumps(context_data)}

The user asks: "{user_message}"

Instructions:
1. Respond directly, conversationally, concisely, and accurately based ONLY on their context data. 
2. Do not invent information. If they ask about something not in the context, politely state you don't have access to that specific info.
3. Format nicely using Markdown. Use emojis sparingly but appropriately. Do not use markdown code blocks for normal text."""
            
            response = model.generate_content(prompt)
            response_text = response.text.strip()
            
        except Exception as e:
            print(f"[WARN] Gemini Chatbot failed: {str(e)}")
            api_key = None # trigger fallback
    
    # Fallback to offline warning
    if not api_key:
        msg_lower = user_message.lower()
        if 'attendance' in msg_lower or 'present' in msg_lower:
            if role == 'student':
                fallback_msg = f"Hello {first_name}! I am in offline mode. You can view your detailed attendance by navigating to the **Attendance** section on your dashboard."
                if 'attendance' in context_data and context_data['attendance']:
                    total_p = sum(c.get('attended', 0) for c in context_data['attendance'])
                    total_t = sum(c.get('total', 0) for c in context_data['attendance'])
                    if total_t > 0:
                        pct = round((total_p / total_t) * 100, 1)
                        fallback_msg = f"Your overall attendance is **{pct}%** ({total_p}/{total_t} classes). "
                        if pct < 80:
                            fallback_msg += "<br>⚠️ You are currently **at risk** (below 80%). Please attend upcoming classes!"
                        else:
                            fallback_msg += "Great job keeping it above the required threshold!"
            else:
                fallback_msg = "As a teacher/admin, you can track course attendance from the **Daily Attendance** menu."
        elif 'grade' in msg_lower or 'score' in msg_lower:
            if role == 'student':
                gpa = context_data.get('gpa', 0)
                fallback_msg = f"In offline mode: Your current GPA is **{gpa}**. Check the **My Grades** section for details."
            else:
                fallback_msg = "Grades can be managed via the **Gradebook** module."
        elif 'risk' in msg_lower or 'fail' in msg_lower:
            if role == 'student':
                fallback_msg = "Please check the AI predictor on your dashboard to see your actual risk profile."
            else:
                fallback_msg = "Risk analytics are available in the dashboard."
        else:
            fallback_msg = f"Hello {first_name}! I am currently running in offline-fallback mode. "
            if role == 'admin':
                fallback_msg += "\n\n⚠️ **Action Required**: Please configure the `Google Gemini API Key` in the System Settings to unlock my true conversational AI capabilities!"
            else:
                fallback_msg += "\n\nI cannot answer complex queries right now. Please ask your administrator to configure the AI system."
            
        response_text = fallback_msg

    return jsonify({'response': response_text, 'role': role})



@socketio.on('teacher_video_on')
def on_teacher_video_on(data):
    """Teacher started video/screen"""
    room = f"classroom_{data.get('session_id')}"
    emit('teacher_video_on', data, room=room, include_self=False)


@socketio.on('student_video_stream')
def on_student_video_stream(data):
    """Student started their camera"""
    room = f"classroom_{data.get('session_id')}"
    emit('student_video_stream', data, room=room, include_self=False)

@socketio.on('raise_hand')
def on_raise_hand(data):
    room = f"classroom_{data.get('session_id')}"
    emit('student_raised_hand', data, room=room, include_self=False)

@socketio.on('draw_line')
def on_draw_line(data):
    room = f"classroom_{data.get('session_id')}"
    emit('draw_line', data, room=room, include_self=False)
    
@socketio.on('clear_whiteboard')
def on_clear_whiteboard(data):
    room = f"classroom_{data.get('session_id')}"
    emit('clear_whiteboard', data, room=room, include_self=False)

@socketio.on('grant_whiteboard_access')
def on_grant_whiteboard_access(data):
    room = f"classroom_{data.get('session_id')}"
    emit('grant_whiteboard_access', data, room=room, include_self=False)

@socketio.on('revoke_whiteboard_access')
def on_revoke_whiteboard_access(data):
    room = f"classroom_{data.get('session_id')}"
    emit('revoke_whiteboard_access', data, room=room, include_self=False)

@socketio.on('student_distracted')
def on_student_distracted(data):
    room = f"classroom_{data.get('session_id')}"
    emit('student_distracted', data, room=room, include_self=False)

@socketio.on('confusion_signal')
def on_confusion_signal(data):
    room = f"classroom_{data.get('session_id')}"
    emit('confusion_signal', data, room=room, include_self=False)

@socketio.on('start_poll')
def on_start_poll(data):
    room = f"classroom_{data.get('session_id')}"
    emit('start_poll', data, room=room, include_self=False)

@socketio.on('end_poll')
def on_end_poll(data):
    room = f"classroom_{data.get('session_id')}"
    emit('end_poll', data, room=room, include_self=False)

@socketio.on('poll_response')
def on_poll_response(data):
    room = f"classroom_{data.get('session_id')}"
    emit('poll_response', data, room=room, include_self=False)

@socketio.on('start_class')
def on_start_class(data):
    session_id = data.get('session_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE classroom_sessions SET status = 'live', started_at = CURRENT_TIMESTAMP WHERE id = ?", (session_id,))
    conn.commit()
    conn.close()
    room = f"classroom_{session_id}"
    emit('class_started', {}, room=room, include_self=False)

@socketio.on('end_class')
def on_end_class(data):
    session_id = data.get('session_id')
    notes = data.get('notes', '')
    transcript = data.get('transcript', '')

    conn = get_db()
    c = conn.cursor()

    final_notes = notes

    if transcript and len(transcript.split()) > 10:
        try:
            import google.generativeai as genai
            import os
            api_key = os.environ.get('GEMINI_API_KEY')
            if api_key:
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel('gemini-2.5-flash')
                prompt = "Summarize this live class transcript into crisp, formatted markdown notes for students with key takeaways. Output only the markdown summary:\\n\\n" + transcript
                response = model.generate_content(prompt)
                ai_summary = response.text.strip()
                final_notes = notes + "\\n\\n### 🤖 AI Class Summary\\n" + ai_summary
        except Exception as e:
            print("[WARN] Failed to summarize transcript:", e)

    # -----------------------------------------------------
    # GAMIFICATION: Award Badges based on Proctor Responses
    # -----------------------------------------------------
    c.execute("""
        SELECT student_id,
               COUNT(*) as total_questions,
               SUM(CASE WHEN is_correct = 1 THEN 1 ELSE 0 END) as correct,
               AVG(CASE WHEN answer_given != -1 THEN time_taken_seconds END) as avg_time
        FROM proctor_responses
        WHERE session_id = ?
        GROUP BY student_id
    """, (session_id,))
    stats = c.fetchall()
    
    for s in stats:
        stu_id = s['student_id']
        total = s['total_questions'] or 0
        correct = s['correct'] or 0
        avg_time = s['avg_time'] or 999
        
        if total >= 2 and correct == total:
            # Grant Perfect Score Badge
            c.execute("SELECT id FROM student_badges WHERE student_id = ? AND badge_name = 'Perfect Score'", (stu_id,))
            if not c.fetchone():
                c.execute("INSERT INTO student_badges (student_id, badge_name, icon, description) VALUES (?, ?, ?, ?)",
                          (stu_id, 'Perfect Score', '🏆', 'Answered all interactive classroom questions perfectly!'))
                          
        if total >= 2 and avg_time < 5.0:
            # Grant Lightning Fast Badge
            c.execute("SELECT id FROM student_badges WHERE student_id = ? AND badge_name = 'Lightning Responder'", (stu_id,))
            if not c.fetchone():
                c.execute("INSERT INTO student_badges (student_id, badge_name, icon, description) VALUES (?, ?, ?, ?)",
                          (stu_id, 'Lightning Responder', '⚡', 'Responded to all classroom questions extremely fast!'))
                          
    c.execute("UPDATE classroom_sessions SET status = 'completed', notes = ?, ended_at = CURRENT_TIMESTAMP WHERE id = ?", (final_notes, session_id))
    conn.commit()
    conn.close()

    room = f"classroom_{session_id}"
    emit('class_ended', {'notes': final_notes}, room=room, include_self=False)


# ============================================================================
# LIVE CLASS RESPONSE ROUTES
# ============================================================================

@app.route('/classroom/respond/<int:session_id>/<action>')
@student_required
def classroom_respond(session_id, action):
    """Student responds to live class notification — join or decline"""
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT id, first_name, last_name FROM students WHERE user_id = ?", (session['user_id'],))
    student = c.fetchone()

    c.execute("""
        SELECT cs.*, c.course_name, c.course_code
        FROM classroom_sessions cs
        JOIN courses c ON cs.course_id = c.id
        WHERE cs.id = ?
    """, (session_id,))
    classroom = c.fetchone()

    if not classroom or not student:
        conn.close()
        flash('Session not found', 'error')
        return redirect(url_for('student_dashboard'))

    if action == 'join':
        # Mark as joining
        c.execute("""
            INSERT OR IGNORE INTO classroom_attendance
            (session_id, student_id, joined_at, status)
            VALUES (?, ?, CURRENT_TIMESTAMP, 'present')
        """, (session_id, student['id']))
        conn.commit()
        conn.close()
        return redirect(url_for('classroom_student', session_id=session_id))

    elif action == 'decline':
        # Save decline record
        c.execute("""
            INSERT OR IGNORE INTO classroom_attendance
            (session_id, student_id, status)
            VALUES (?, ?, 'absent')
        """, (session_id, student['id']))

        # Get instructor user_id to notify
        c.execute("""
            SELECT u.id as user_id
            FROM instructors i
            JOIN users u ON i.user_id = u.id
            WHERE i.id = (SELECT instructor_id FROM classroom_sessions WHERE id = ?)
        """, (session_id,))
        instructor = c.fetchone()

        if instructor:
            c.execute("""
                INSERT INTO notifications (user_id, title, message, type)
                VALUES (?, ?, ?, ?)
            """, (
                instructor['user_id'],
                f"❌ Student Cannot Attend — {classroom['course_code']}",
                f"{student['first_name']} {student['last_name']} cannot attend the live class for {classroom['course_name']}.",
                'warning'
            ))

        conn.commit()
        conn.close()

        # Return page with decline confirmation
        return render_template('classroom_declined.html',
                               classroom=classroom,
                               student_name=f"{student['first_name']} {student['last_name']}")

    conn.close()
    return redirect(url_for('student_dashboard'))


@app.route('/classroom/session-report/<int:session_id>')
@faculty_required
def classroom_session_report(session_id):
    """Detailed AI-powered session report for faculty"""
    conn = get_db()
    c = conn.cursor()

    # Get session info
    c.execute("""
        SELECT cs.*, c.course_name, c.course_code, c.id as course_id
        FROM classroom_sessions cs
        JOIN courses c ON cs.course_id = c.id
        WHERE cs.id = ?
    """, (session_id,))
    classroom = c.fetchone()

    if not classroom:
        conn.close()
        flash('Session not found', 'error')
        return redirect(url_for('classroom_sessions'))

    # Get all enrolled students for this course
    c.execute("""
        SELECT s.id, s.student_id, s.first_name, s.last_name
        FROM students s
        JOIN enrollments e ON s.id = e.student_id
        WHERE e.course_id = ?
        ORDER BY s.last_name, s.first_name
    """, (classroom['course_id'],))
    enrolled_students = c.fetchall()

    # Get attendance records
    c.execute("""
        SELECT ca.*, s.first_name, s.last_name, s.student_id as stu_id
        FROM classroom_attendance ca
        JOIN students s ON ca.student_id = s.id
        WHERE ca.session_id = ?
    """, (session_id,))
    attendance_records = {row['student_id']: row for row in c.fetchall()}

    # Get proctor responses per student
    c.execute("""
        SELECT pr.student_id,
               COUNT(*) as total_questions,
               SUM(CASE WHEN pr.answer_given != -1 THEN 1 ELSE 0 END) as answered,
               SUM(CASE WHEN pr.answer_given = -1 THEN 1 ELSE 0 END) as missed,
               SUM(CASE WHEN pr.is_correct = 1 THEN 1 ELSE 0 END) as correct,
               SUM(CASE WHEN pr.is_correct = 0 AND pr.answer_given != -1 THEN 1 ELSE 0 END) as wrong,
               AVG(CASE WHEN pr.answer_given != -1 THEN pr.time_taken_seconds END) as avg_time
        FROM proctor_responses pr
        WHERE pr.session_id = ?
        GROUP BY pr.student_id
    """, (session_id,))
    proctor_stats = {row['student_id']: row for row in c.fetchall()}

    # Get screen time per student (joined_at to left_at)
    c.execute("""
        SELECT student_id,
               joined_at,
               left_at,
               CASE
                   WHEN left_at IS NOT NULL AND joined_at IS NOT NULL
                   THEN CAST((julianday(left_at) - julianday(joined_at)) * 24 * 60 AS INTEGER)
                   ELSE NULL
               END as screen_minutes
        FROM classroom_attendance
        WHERE session_id = ?
    """, (session_id,))
    screen_times = {row['student_id']: row for row in c.fetchall()}

    # Calculate total class duration in minutes
    class_duration = 0
    if classroom['started_at'] and classroom['ended_at']:
        from datetime import datetime as dt
        try:
            fmt = '%Y-%m-%d %H:%M:%S'
            start = dt.strptime(classroom['started_at'], fmt)
            end = dt.strptime(classroom['ended_at'], fmt)
            class_duration = int((end - start).total_seconds() / 60)
        except:
            class_duration = 0

    # Build comprehensive report per student
    student_reports = []
    for stu in enrolled_students:
        att = attendance_records.get(stu['id'], {})
        prc = proctor_stats.get(stu['id'], {})
        sct = screen_times.get(stu['id'], {})

        screen_mins = sct.get('screen_minutes') if sct else None
        screen_pct = round((screen_mins / class_duration * 100)) if screen_mins and class_duration > 0 else 0

        # AI attendance decision
        status = att.get('status', 'absent')
        was_kicked = att.get('was_kicked', 0)
        answered = prc.get('answered', 0) or 0
        missed = prc.get('missed', 0) or 0
        total_q = prc.get('total_questions', 0) or 0

        # AI auto-attendance logic
        ai_status = 'absent'
        ai_reason = 'Did not join the class'
        if status == 'present' or status == 'late':
            if was_kicked:
                ai_status = 'kicked'
                ai_reason = 'Removed for not responding to questions'
            elif screen_pct < 30 and class_duration > 0:
                ai_status = 'suspicious'
                ai_reason = f'Low screen time ({screen_pct}%) — possible cheating'
            elif missed >= 3:
                ai_status = 'warned'
                ai_reason = f'Missed {missed} questions'
            elif answered > 0:
                ai_status = 'present'
                ai_reason = 'Attended and responded to questions'
            else:
                ai_status = 'present'
                ai_reason = 'Joined but no questions answered'
        elif status == 'absent':
            ai_status = 'absent'
            ai_reason = att.get('left_at') and 'Declined to attend' or 'Did not join'

        student_reports.append({
            'id': stu['id'],
            'student_id': stu['stu_id'] if att else stu['student_id'],
            'name': f"{stu['first_name']} {stu['last_name']}",
            'status': status,
            'ai_status': ai_status,
            'ai_reason': ai_reason,
            'was_kicked': was_kicked,
            'total_questions': total_q,
            'answered': answered,
            'missed': missed,
            'correct': prc.get('correct', 0) or 0,
            'wrong': prc.get('wrong', 0) or 0,
            'avg_response_time': round(prc.get('avg_time', 0) or 0, 1),
            'screen_minutes': screen_mins,
            'screen_pct': screen_pct,
            'joined_at': att.get('joined_at', '—'),
            'left_at': att.get('left_at', '—'),
            'warned': 1 if missed >= 2 else 0,
        })

    # Summary stats
    total_enrolled = len(enrolled_students)
    total_present = sum(1 for r in student_reports if r['ai_status'] in ['present', 'warned'])
    total_absent = sum(1 for r in student_reports if r['ai_status'] == 'absent')
    total_kicked = sum(1 for r in student_reports if r['ai_status'] == 'kicked')
    total_suspicious = sum(1 for r in student_reports if r['ai_status'] == 'suspicious')

    # Auto-save AI attendance to regular attendance table
    today = datetime.now().strftime('%Y-%m-%d')
    for r in student_reports:
        final_status = 'present' if r['ai_status'] in ['present','warned','suspicious'] else 'absent'
        c.execute("""
            INSERT OR REPLACE INTO attendance
            (student_id, course_id, date, status, marked_by)
            VALUES (?, ?, ?, ?, ?)
        """, (r['id'], classroom['course_id'], today, final_status, session['user_id']))
    conn.commit()

    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    conn.close()

    return render_template('classroom_session_report.html',
                           classroom=classroom,
                           student_reports=student_reports,
                           class_duration=class_duration,
                           total_enrolled=total_enrolled,
                           total_present=total_present,
                           total_absent=total_absent,
                           total_kicked=total_kicked,
                           total_suspicious=total_suspicious,
                           unread_count=unread_count,
                           session_id=session_id)


@socketio.on('student_declined')
def on_student_declined(data):
    """Student declined to join — notify teacher"""
    room = f"classroom_{data.get('session_id')}"
    emit('student_declined', {
        'student_name': data.get('student_name'),
        'reason': data.get('reason', 'Cannot attend')
    }, room=room)


@socketio.on('student_left')
def on_student_left(data):
    """Student left the class"""
    session_id = data.get('session_id')
    student_id = data.get('student_id')
    room = f"classroom_{session_id}"

    conn = get_db()
    c = conn.cursor()
    c.execute("""
        UPDATE classroom_attendance
        SET left_at = CURRENT_TIMESTAMP
        WHERE session_id = ? AND student_id = ?
    """, (session_id, student_id))
    conn.commit()
    conn.close()

    emit('student_left', {
        'student_id': student_id,
        'student_name': data.get('student_name')
    }, room=room)



# ============================================================================
# EMAIL HELPER
# ============================================================================

def send_email(to_email, subject, html_body, attachment_path=None):
    """Send email using Gmail SMTP"""
    try:
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT * FROM email_settings WHERE enabled = 1 LIMIT 1")
        email_config = c.fetchone()
        # Do not close connection here, it is shared across the request (g.db)
        
        if not email_config:
            return False, "Email not configured"

        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['From'] = email_config['smtp_email']
        msg['To'] = to_email

        msg.attach(MIMEText(html_body, 'html'))

        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(attachment_path)}"')
                msg.attach(part)

        server = smtplib.SMTP(email_config['smtp_server'], email_config['smtp_port'])
        server.starttls()
        server.login(email_config['smtp_email'], email_config['smtp_password'])
        server.send_message(msg)
        server.quit()
        return True, "Email sent successfully"

    except Exception as e:
        return False, str(e)


def send_bulk_email(recipients, subject, html_body):
    """Send email to multiple recipients"""
    results = []
    for email in recipients:
        if email:
            success, msg = send_email(email, subject, html_body)
            results.append({'email': email, 'success': success, 'message': msg})
    return results


def get_email_template(title, body, footer="EduCore — Student Management System"):
    """Generate professional HTML email template"""
    return f"""
    <!DOCTYPE html>
    <html>
    <head><meta charset="UTF-8"></head>
    <body style="font-family:'Segoe UI',Arial,sans-serif;background:#f8fafc;margin:0;padding:20px;">
        <div style="max-width:600px;margin:0 auto;background:white;border-radius:12px;
                    overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.08);">
            <div style="background:linear-gradient(135deg,#1a56db,#0891b2);padding:28px 32px;">
                <h1 style="color:white;margin:0;font-size:1.4rem;">🎓 EduCore</h1>
                <p style="color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:0.9rem;">Smart Learning Management System</p>
            </div>
            <div style="padding:32px;">
                <h2 style="color:#0f172a;font-size:1.2rem;margin:0 0 16px;">{title}</h2>
                {body}
            </div>
            <div style="background:#f8fafc;padding:16px 32px;border-top:1px solid #e2e8f0;
                        text-align:center;font-size:0.8rem;color:#94a3b8;">
                {footer}
            </div>
        </div>
    </body>
    </html>
    """


# ============================================================================
# EMAIL SETTINGS ROUTES
# ============================================================================

@app.route('/admin/email-settings', methods=['GET', 'POST'])
@admin_required
def email_settings_page():
    """Configure email settings"""
    conn = get_db()
    c = conn.cursor()

    if request.method == 'POST':
        smtp_email = request.form.get('smtp_email')
        smtp_password = request.form.get('smtp_password')
        enabled = 1 if request.form.get('enabled') else 0

        c.execute("SELECT id FROM email_settings LIMIT 1")
        existing = c.fetchone()

        if existing:
            c.execute("""UPDATE email_settings
                SET smtp_email=?, smtp_password=?, enabled=?
                WHERE id=?""", (smtp_email, smtp_password, enabled, existing['id']))
        else:
            c.execute("""INSERT INTO email_settings (smtp_email, smtp_password, enabled)
                VALUES (?, ?, ?)""", (smtp_email, smtp_password, enabled))

        conn.commit()

        # Test email
        if enabled:
            success, msg = send_email(smtp_email, "EduCore Email Test",
                get_email_template("Email Setup Successful! ✅",
                "<p>Your EduCore email notifications are now active.</p>"))
            if success:
                flash('Email settings saved and test email sent successfully!', 'success')
            else:
                flash(f'Settings saved but test email failed: {msg}', 'warning')
        else:
            flash('Email settings saved!', 'success')

        conn.close()
        return redirect(url_for('email_settings_page'))

    c.execute("SELECT * FROM email_settings LIMIT 1")
    config = c.fetchone()
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    conn.close()

    return render_template('email_settings.html', config=config, unread_count=unread_count)


@app.route('/admin/email/send-bulk', methods=['POST'])
@admin_required
def send_bulk_email_route():
    """Send bulk email to all students or specific group"""
    target = request.form.get('target', 'all_students')
    subject = request.form.get('subject')
    message = request.form.get('message')

    conn = get_db()
    c = conn.cursor()

    if target == 'all_students':
        c.execute("SELECT u.email FROM users u JOIN students s ON u.id = s.user_id WHERE u.email IS NOT NULL")
    elif target == 'at_risk':
        c.execute("""SELECT DISTINCT u.email FROM users u
            JOIN students s ON u.id = s.user_id
            JOIN attendance a ON s.id = a.student_id
            WHERE u.email IS NOT NULL
            GROUP BY s.id
            HAVING AVG(CASE WHEN a.status IN ('present','late') THEN 1.0 ELSE 0.0 END) < 0.8""")
    elif target == 'all_faculty':
        c.execute("SELECT u.email FROM users u JOIN instructors i ON u.id = i.user_id WHERE u.email IS NOT NULL")

    emails = [row['email'] for row in c.fetchall()]
    conn.close()

    html_body = get_email_template(subject,
        f"<p style='color:#475569;line-height:1.6;'>{message}</p>")

    results = send_bulk_email(emails, subject, html_body)
    success_count = sum(1 for r in results if r['success'])

    flash(f'Email sent to {success_count}/{len(emails)} recipients!', 'success')
    return redirect(url_for('email_settings_page'))


# ============================================================================
# FILE UPLOAD HELPER
# ============================================================================

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


# ============================================================================
# ASSIGNMENT ROUTES
# ============================================================================

@app.route('/faculty/assignments/<int:course_id>')
@faculty_required
def faculty_assignments(course_id):
    """View and manage assignments for a course"""
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT * FROM courses WHERE id = ?", (course_id,))
    course = c.fetchone()

    c.execute("""SELECT a.*, COUNT(s.id) as submission_count
        FROM assignments a
        LEFT JOIN submissions s ON a.id = s.assignment_id
        WHERE a.course_id = ?
        GROUP BY a.id
        ORDER BY a.due_date DESC""", (course_id,))
    assignments = c.fetchall()

    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    conn.close()

    return render_template('faculty_assignments.html',
                           course=course,
                           assignments=assignments,
                           unread_count=unread_count)


@app.route('/faculty/assignments/create/<int:course_id>', methods=['GET', 'POST'])
@faculty_required
def create_assignment(course_id):
    """Create new assignment"""
    conn = get_db()
    c = conn.cursor()

    if request.method == 'POST':
        title = request.form.get('title')
        description = request.form.get('description')
        due_date = request.form.get('due_date')
        max_marks = request.form.get('max_marks', 100)
        file_path = None

        if 'file' in request.files:
            file = request.files['file']
            if file and file.filename and allowed_file(file.filename):
                filename = secure_filename(f"assign_{course_id}_{file.filename}")
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'assignments', filename)
                file.save(file_path)
                file_path = f"assignments/{filename}"

        c.execute("""INSERT INTO assignments (course_id, title, description, due_date, max_marks, file_path, created_by)
            VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (course_id, title, description, due_date, max_marks, file_path, session['user_id']))
        assignment_id = c.lastrowid

        # Notify enrolled students
        c.execute("""SELECT u.id as user_id, u.email, s.first_name, s.last_name
            FROM students s JOIN users u ON s.user_id = u.id
            JOIN enrollments e ON s.id = e.student_id
            WHERE e.course_id = ?""", (course_id,))
        students = c.fetchall()

        c.execute("SELECT course_name, course_code FROM courses WHERE id = ?", (course_id,))
        course = c.fetchone()

        for stu in students:
            # In-app notification
            c.execute("""INSERT INTO notifications (user_id, title, message, type)
                VALUES (?, ?, ?, ?)""",
                (stu['user_id'],
                 f"📝 New Assignment — {course['course_code']}",
                 f"New assignment '{title}' has been posted. Due: {due_date}",
                 'info'))

            # Email notification
            if stu['email']:
                html = get_email_template(
                    f"New Assignment Posted — {course['course_code']}",
                    f"""<p>Dear {stu['first_name']},</p>
                    <p>A new assignment has been posted for <strong>{course['course_name']}</strong>.</p>
                    <div style='background:#f8fafc;border-radius:10px;padding:16px;margin:16px 0;'>
                        <p><strong>Assignment:</strong> {title}</p>
                        <p><strong>Due Date:</strong> {due_date}</p>
                        <p><strong>Max Marks:</strong> {max_marks}</p>
                        <p><strong>Description:</strong> {description or 'No description'}</p>
                    </div>
                    <p>Login to EduCore to view and submit your assignment.</p>""")
                send_email(stu['email'], f"New Assignment — {title}", html)

        conn.commit()
        conn.close()
        flash('Assignment created and students notified!', 'success')
        return redirect(url_for('faculty_assignments', course_id=course_id))

    c.execute("SELECT * FROM courses WHERE id = ?", (course_id,))
    course = c.fetchone()
    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    conn.close()
    return render_template('create_assignment.html', course=course, unread_count=unread_count)


@app.route('/faculty/assignments/<int:assignment_id>/submissions')
@faculty_required
def view_submissions(assignment_id):
    """View all submissions for an assignment"""
    conn = get_db()
    c = conn.cursor()

    c.execute("""SELECT a.*, c.course_name, c.course_code
        FROM assignments a JOIN courses c ON a.course_id = c.id
        WHERE a.id = ?""", (assignment_id,))
    assignment = c.fetchone()

    c.execute("""SELECT s.*, st.first_name, st.last_name, st.student_id as stu_id
        FROM submissions s JOIN students st ON s.student_id = st.id
        WHERE s.assignment_id = ?
        ORDER BY s.submitted_at DESC""", (assignment_id,))
    submissions = c.fetchall()

    # Students who haven't submitted
    c.execute("""SELECT st.id, st.first_name, st.last_name, st.student_id
        FROM students st JOIN enrollments e ON st.id = e.student_id
        WHERE e.course_id = ?
        AND st.id NOT IN (SELECT student_id FROM submissions WHERE assignment_id = ?)""",
        (assignment['course_id'], assignment_id))
    not_submitted = c.fetchall()

    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    conn.close()

    return render_template('view_submissions.html',
                           assignment=assignment,
                           submissions=submissions,
                           not_submitted=not_submitted,
                           unread_count=unread_count)


@app.route('/faculty/submissions/<int:submission_id>/grade', methods=['POST'])
@faculty_required
def grade_submission(submission_id):
    """Grade a submission"""
    marks = request.form.get('marks')
    feedback = request.form.get('feedback')

    conn = get_db()
    c = conn.cursor()
    c.execute("""UPDATE submissions SET marks=?, feedback=?, status='graded'
        WHERE id=?""", (marks, feedback, submission_id))

    # Get student email for notification
    c.execute("""SELECT u.email, s.first_name, a.title, a.max_marks, e.course_id
        FROM submissions sub
        JOIN students s ON sub.student_id = s.id
        JOIN users u ON s.user_id = u.id
        JOIN assignments a ON sub.assignment_id = a.id
        JOIN enrollments e ON s.id = e.student_id AND e.course_id = a.course_id
        WHERE sub.id = ?""", (submission_id,))
    info = c.fetchone()

    if info and info['email']:
        html = get_email_template(
            f"Assignment Graded — {info['title']}",
            f"""<p>Dear {info['first_name']},</p>
            <p>Your assignment has been graded.</p>
            <div style='background:#f8fafc;border-radius:10px;padding:16px;margin:16px 0;'>
                <p><strong>Marks:</strong> {marks} / {info['max_marks']}</p>
                <p><strong>Feedback:</strong> {feedback or 'No feedback provided'}</p>
            </div>""")
        send_email(info['email'], f"Assignment Graded — {info['title']}", html)

    conn.commit()
    conn.close()
    flash('Submission graded!', 'success')
    return redirect(request.referrer or url_for('teacher_dashboard'))


# ============================================================================
# STUDENT SUBMISSION ROUTES
# ============================================================================

@app.route('/student/assignments')
@student_required
def student_assignments():
    """View all assignments for enrolled courses"""
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT id FROM students WHERE user_id = ?", (session['user_id'],))
    student = c.fetchone()

    c.execute("""SELECT a.*, c.course_name, c.course_code,
        s.id as submission_id, s.submitted_at, s.marks, s.status as sub_status,
        CASE WHEN a.due_date < date('now') THEN 1 ELSE 0 END as is_overdue
        FROM assignments a
        JOIN courses c ON a.course_id = c.id
        JOIN enrollments e ON c.id = e.course_id AND e.student_id = ?
        LEFT JOIN submissions s ON a.id = s.assignment_id AND s.student_id = ?
        ORDER BY a.due_date ASC""", (student['id'], student['id']))
    assignments = c.fetchall()

    c.execute('SELECT COUNT(*) as count FROM notifications WHERE user_id = ? AND is_read = 0',
              (session['user_id'],))
    unread_count = c.fetchone()['count']
    conn.close()

    return render_template('student_assignments.html',
                           assignments=assignments,
                           unread_count=unread_count)


@app.route('/student/assignments/<int:assignment_id>/submit', methods=['POST'])
@student_required
def submit_assignment(assignment_id):
    """Submit an assignment"""
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT id FROM students WHERE user_id = ?", (session['user_id'],))
    student = c.fetchone()

    # Check if already submitted
    c.execute("SELECT id FROM submissions WHERE assignment_id = ? AND student_id = ?",
              (assignment_id, student['id']))
    existing = c.fetchone()

    if existing:
        conn.close()
        flash('You have already submitted this assignment!', 'warning')
        return redirect(url_for('student_assignments'))

    file_path = None
    if 'file' in request.files:
        file = request.files['file']
        if file and file.filename and allowed_file(file.filename):
            filename = secure_filename(f"sub_{assignment_id}_{student['id']}_{file.filename}")
            full_path = os.path.join(app.config['UPLOAD_FOLDER'], 'submissions', filename)
            file.save(full_path)
            file_path = f"submissions/{filename}"
        else:
            conn.close()
            flash('Invalid file type! Allowed: PDF, DOC, DOCX, TXT, PNG, JPG, ZIP, PPTX, XLSX', 'error')
            return redirect(url_for('student_assignments'))

    c.execute("""INSERT INTO submissions (assignment_id, student_id, file_path, status)
        VALUES (?, ?, ?, 'submitted')""", (assignment_id, student['id'], file_path))

    # Notify faculty
    c.execute("""SELECT u.id as faculty_user_id, a.title, c.course_code,
        s.first_name, s.last_name
        FROM assignments a
        JOIN courses c ON a.course_id = c.id
        JOIN instructors i ON c.instructor_id = i.id
        JOIN users u ON i.user_id = u.id
        JOIN students s ON s.id = ?
        WHERE a.id = ?""", (student['id'], assignment_id))
    info = c.fetchone()

    if info:
        c.execute("""INSERT INTO notifications (user_id, title, message, type)
            VALUES (?, ?, ?, ?)""",
            (info['faculty_user_id'],
             f"📥 New Submission — {info['course_code']}",
             f"{info['first_name']} {info['last_name']} submitted '{info['title']}'",
             'info'))

    conn.commit()
    conn.close()
    flash('Assignment submitted successfully!', 'success')
    return redirect(url_for('student_assignments'))


@app.route('/uploads/<path:filename>')
@login_required
def download_file(filename):
    """Download uploaded file"""
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)


# ============================================================================
# AUTO EMAIL ALERTS FOR LOW ATTENDANCE
# ============================================================================

@app.route('/admin/email/send-attendance-alerts')
@admin_required
def send_attendance_alerts():
    """Send email alerts to at-risk students"""
    from ai_attendance import analyse_student_course

    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT setting_value FROM settings WHERE setting_key = 'total_semester_classes'")
    row = c.fetchone()
    total_sem = int(row['setting_value']) if row else 90

    c.execute("""SELECT s.id, s.first_name, s.last_name, u.email
        FROM students s JOIN users u ON s.user_id = u.id
        WHERE s.status = 'active' AND u.email IS NOT NULL""")
    students = c.fetchall()

    alert_count = 0
    for stu in students:
        c.execute("""SELECT c.id, c.course_code, c.course_name
            FROM courses c JOIN enrollments e ON c.id = e.course_id
            WHERE e.student_id = ?""", (stu['id'],))
        courses = c.fetchall()

        at_risk_courses = []
        for course in courses:
            c.execute("""SELECT COUNT(CASE WHEN status IN ('present','late') THEN 1 END) as p,
                COUNT(*) as t FROM attendance WHERE student_id = ? AND course_id = ?""",
                (stu['id'], course['id']))
            att = c.fetchone()
            analysis = analyse_student_course(att['p'] or 0, att['t'] or 0, total_sem)
            if analysis['send_alert']:
                at_risk_courses.append({
                    'code': course['course_code'],
                    'name': course['course_name'],
                    'pct': analysis['attendance_pct'],
                    'prob': analysis['probability_of_80'],
                    'days': analysis['days_needed_to_recover']
                })

        if at_risk_courses and stu['email']:
            courses_html = ""
            for cr in at_risk_courses:
                courses_html += f"""
                <div style='background:#fff5f5;border-left:4px solid #dc2626;
                            border-radius:8px;padding:12px;margin:8px 0;'>
                    <strong>{cr['code']} — {cr['name']}</strong><br>
                    Attendance: <strong style='color:#dc2626;'>{cr['pct']}%</strong> |
                    Probability of passing: <strong>{cr['prob']}%</strong><br>
                    Need to attend: <strong>{cr['days']} consecutive classes</strong>
                </div>"""

            html = get_email_template(
                "⚠️ Attendance Alert — Action Required",
                f"""<p>Dear {stu['first_name']},</p>
                <p>Your attendance is below the required 80% in the following courses:</p>
                {courses_html}
                <p style='margin-top:16px;color:#475569;'>
                Please attend all upcoming classes to avoid academic issues.
                Login to EduCore for detailed AI recovery plan.</p>""")

            success, _ = send_email(stu['email'],
                "⚠️ Attendance Alert — EduCore", html)
            if success:
                alert_count += 1

    conn.close()
    flash(f'Attendance alerts sent to {alert_count} students!', 'success')
    return redirect(url_for('email_settings_page'))



if __name__ == '__main__':
    print("=" * 60)
    print("[*] Student Management System - Backend")
    print("=" * 60)
    print("[>] Template folder: templates")
    print("[>] Static folder: static")
    print("[>] Database: database.db")
    print("[>] Logs: logs/app.log")
    print("=" * 60)
    print("[*] Default Login Credentials:")
    print("    Admin:   admin / admin123")
    print("    Student: student / student123")
    print("    Faculty: faculty / faculty123")
    print("=" * 60)
    print("[*] Server starting at: http://127.0.0.1:5001")
    print("[*] Student Login: http://127.0.0.1:5001/student/login")
    print("[*] Faculty Login: http://127.0.0.1:5001/faculty/login")
    print("[*] Admin Login: http://127.0.0.1:5001/admin/login")
    print("=" * 60)
    
    # Create required directories if they don't exist
    os.makedirs('templates', exist_ok=True)
    os.makedirs('static/css', exist_ok=True)
    os.makedirs('static/js', exist_ok=True)
    os.makedirs('static/img', exist_ok=True)
    os.makedirs('logs', exist_ok=True)
    
    socketio.run(app, debug=True, host='127.0.0.1', port=5001)