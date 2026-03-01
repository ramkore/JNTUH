"""Helper script for install.bat — checks which packages are missing."""
import sys
import os

REQUIRED = ['pandas', 'numpy', 'openpyxl', 'reportlab']

missing = []
for pkg in REQUIRED:
    try:
        __import__(pkg)
        print(f'  [SKIP] {pkg} - already installed')
    except ImportError:
        print(f'  [NEED] {pkg}')
        missing.append(pkg)

if not missing:
    print('  All required packages are already installed.')
else:
    # Write missing packages to temp file for install.bat to read
    tmp = os.path.join(os.environ.get('TEMP', '.'), 'marks_missing_pkgs.txt')
    with open(tmp, 'w') as f:
        f.write(' '.join(missing))
