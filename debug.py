import os
import sys

print("=" * 50)
print("DEBUGGING YOUR STUDENT MANAGEMENT SYSTEM")
print("=" * 50)

# Check if app.py exists
if os.path.exists('app.py'):
    print("✅ app.py found")
else:
    print("❌ app.py NOT found!")
    sys.exit(1)

# Check templates folder
if os.path.exists('templates'):
    print("✅ templates folder found")
    
    # List template files
    templates = os.listdir('templates')
    required_templates = [
        'layout.html', 'login.html', 'student_login.html', 
        'faculty_login.html', 'admin_login.html', 
        'student_dashboard.html', 'teacher_dashboard.html'
    ]
    
    print("\nChecking required templates:")
    for template in required_templates:
        if template in templates:
            print(f"  ✅ {template}")
        else:
            print(f"  ❌ {template} MISSING!")
else:
    print("❌ templates folder NOT found!")

# Check database
if os.path.exists('database.db'):
    print(f"\n✅ database.db found (size: {os.path.getsize('database.db')} bytes)")
else:
    print("\n❌ database.db NOT found! (will be created when you run app.py)")

print("\n" + "=" * 50)
print("Next steps:")
print("1. Run: python app.py")
print("2. Open browser to: http://localhost:5001")
print("3. Try these logins:")
print("   - Admin: admin / admin123")
print("   - Student: student / student123")
print("   - Faculty: faculty / faculty123")
print("=" * 50)
