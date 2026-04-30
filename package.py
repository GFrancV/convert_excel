#!/usr/bin/env python3
"""Build a distributable ZIP for excel-converter.

Output: dist/excel-converter-v<version>.zip
Contents: convert_excel.py, requirements.txt, README.md, README.txt
"""

import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).parent

INCLUDE = [
    "convert_excel.py",
    "requirements.txt",
    "README.md",
    "README.txt",
]


def get_version() -> str:
    src = (ROOT / "convert_excel.py").read_text(encoding="utf-8")
    m = re.search(r'^__version__\s*=\s*["\']([^"\']+)["\']', src, re.MULTILINE)
    if not m:
        sys.exit("Error: could not read __version__ from convert_excel.py")
    return m.group(1)


def main():
    version = get_version()
    dist_dir = ROOT / "dist"
    dist_dir.mkdir(exist_ok=True)

    zip_name = f"excel-converter-v{version}.zip"
    zip_path = dist_dir / zip_name

    missing = [f for f in INCLUDE if not (ROOT / f).exists()]
    if missing:
        sys.exit(f"Error: missing file(s): {', '.join(missing)}")

    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for filename in INCLUDE:
            zf.write(ROOT / filename, arcname=filename)
            print(f"  added  {filename}")

    print(f"\nPackage ready: {zip_path}")


if __name__ == "__main__":
    main()
