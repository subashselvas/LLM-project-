import sqlite3
from werkzeug.security import generate_password_hash
from datetime import datetime

conn = sqlite3.connect('database.db')
c = conn.cursor()

# Check if we already have users
c.execute("SELECT COUNT(*) FROM users")
if c.fetchone()[0] == 0:
    # Add admin
    admin_pw = generate_password_hash('admin123')
    c.execute("INSERT INTO users (username, email, password_hash, first_name, last_name, role) VALUES (?,?,?,?,?,?)",
              ('admin', 'admin@system.com', admin_pw, 'arjun', 'Admin', 'admin'))
    
    # Add faculty
    faculty_pw = generate_password_hash('faculty123')
    c.execute("INSERT INTO users (username, email, password_hash, first_name, last_name, role) VALUES (?,?,?,?,?,?)",
              ('selvam', 'selvam@school.edu', faculty_pw, 'Selvam', 'Faculty', 'teacher'))
    faculty_id = c.lastrowid
    c.execute("INSERT INTO instructors (user_id, instructor_id, first_name, last_name, email, joining_date, status) VALUES (?,?,?,?,?,?,?)",
              (faculty_id, 'FAC001', 'Selvam', 'Faculty', 'selvam@school.edu', datetime.now().strftime('%Y-%m-%d'), 'active'))
    
    # Add a student
    student_pw = generate_password_hash('student123')
    c.execute("INSERT INTO users (username, email, password_hash, first_name, last_name, role) VALUES (?,?,?,?,?,?)",
              ('murali', 'murali@school.edu', student_pw, 'Murali', 'Krishnan', 'student'))
    student_id = c.lastrowid
    c.execute("INSERT INTO students (user_id, student_id, first_name, last_name, email, enrollment_date, status) VALUES (?,?,?,?,?,?,?)",
              (student_id, 'STU001', 'Murali', 'Krishnan', 'murali@school.edu', datetime.now().strftime('%Y-%m-%d'), 'active'))
    
    # Add a department
    c.execute("INSERT OR IGNORE INTO departments (dept_code, dept_name) VALUES (?,?)", ('CS', 'Computer Science'))
    
    # Add a course
    c.execute("SELECT id FROM departments WHERE dept_code='CS'")
    dept_id = c.fetchone()[0]
    c.execute("SELECT id FROM instructors LIMIT 1")
    inst_id = c.fetchone()[0]
    c.execute("INSERT INTO courses (course_code, course_name, department_id, instructor_id) VALUES (?,?,?,?)",
              ('CS101', 'Programming Basics', dept_id, inst_id))
    
    conn.commit()
    print("Sample data added!")
else:
    print("Database already has users – skipping.")

conn.close()