import sqlite3
from datetime import datetime

conn = sqlite3.connect('database.db')
c = conn.cursor()

# Check if students table is empty
c.execute("SELECT COUNT(*) FROM students")
if c.fetchone()[0] == 0:
    # Find user IDs for students
    c.execute("SELECT id, username, first_name, last_name, email FROM users WHERE role='student'")
    students = c.fetchall()
    for user in students:
        user_id, username, first_name, last_name, email = user
        # Generate a student_id (e.g., STU001, STU002...)
        c.execute("SELECT COUNT(*) FROM students")
        count = c.fetchone()[0] + 1
        student_id = f"STU{count:03d}"
        c.execute('''
            INSERT INTO students (user_id, student_id, first_name, last_name, email, enrollment_date, status)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (user_id, student_id, first_name, last_name, email, datetime.now().strftime('%Y-%m-%d'), 'active'))
    print("Students added.")

# Check if courses table is empty
c.execute("SELECT COUNT(*) FROM courses")
if c.fetchone()[0] == 0:
    # Add a sample course
    c.execute("SELECT id FROM departments LIMIT 1")
    dept_row = c.fetchone()
    dept_id = dept_row[0] if dept_row else None
    if not dept_id:
        # Create a default department if none exists
        c.execute("INSERT INTO departments (dept_code, dept_name) VALUES (?, ?)", ('CS', 'Computer Science'))
        dept_id = c.lastrowid
    c.execute("SELECT id FROM instructors LIMIT 1")
    inst_row = c.fetchone()
    inst_id = inst_row[0] if inst_row else None
    c.execute('''
        INSERT INTO courses (course_code, course_name, department_id, instructor_id)
        VALUES (?, ?, ?, ?)
    ''', ('CS101', 'Introduction to Programming', dept_id, inst_id))
    print("Course added.")

conn.commit()
conn.close()
print("Done. Now run add_sample_attendance.py again.")