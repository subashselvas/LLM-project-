import sys
import os

# Add parent directory to sys.path so APP and local modules are importable
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from APP import app

# Vercel looks for 'app' in api/index.py
if __name__ == '__main__':
    app.run()
