"""Helper script for install.bat — verifies all packages are importable."""
REQUIRED = ['pandas', 'numpy', 'openpyxl', 'reportlab', 'tkinter']

all_ok = True
for pkg in REQUIRED:
    try:
        __import__(pkg)
        print(f'  [OK] {pkg}')
    except ImportError:
        print(f'  [FAIL] {pkg}')
        all_ok = False

if not all_ok:
    exit(1)
