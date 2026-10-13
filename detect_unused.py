import ast
import glob

def check_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        src = f.read()
    try:
        tree = ast.parse(src)
    except Exception:
        return

    imported_names = {}
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for name in node.names:
                imported_names[name.asname or name.name.split('.')[0]] = node.lineno
        elif isinstance(node, ast.ImportFrom):
            for name in node.names:
                imported_names[name.asname or name.name.split('.')[0]] = node.lineno

    used_names = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Name):
            used_names.add(node.id)

    unused = []
    for imp, line in imported_names.items():
        if imp not in used_names:
            unused.append((imp, line))
    
    if unused:
        print(f"[{filepath}] Unused imports detected:")
        for u, l in unused:
            print(f"  Line {l}: {u}")

for f in ["APP.PY", "ai_features.py", "ai_questions.py", "ai_attendance.py"]:
    check_file(f)
