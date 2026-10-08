import sqlite3
from datetime import datetime, timedelta

conn = sqlite3.connect('database.db')
c = conn.cursor()

# Get student IDs and course IDs (assuming at least one student and one course exist)
students = c.execute("SELECT id FROM students LIMIT 5").fetchall()
courses = c.execute("SELECT id FROM courses LIMIT 2").fetchall()

if not students or not courses:
    print("No students or courses found. Add some first.")
else:
    # Create 10 random attendance records
    for i in range(10):
        student_id = students[i % len(students)][0]
        course_id = courses[i % len(courses)][0]
        date = (datetime.now() - timedelta(days=i)).strftime('%Y-%m-%d')
        status = ['present', 'absent', 'late', 'excused'][i % 4]
        c.execute('''
            INSERT OR IGNORE INTO attendance (student_id, course_id, date, status)
            VALUES (?, ?, ?, ?)
        ''', (student_id, course_id, date, status))
    
    conn.commit()
    print("Sample attendance added!")

conn.close()