#!/usr/bin/env python3
"""Build the distributable .exe ZIP for excel-converter.

Runs PyInstaller with excel_converter.spec (--onefile mode), then
packages the resulting executable into a ZIP for distribution.

Output: dist/excel-converter-v<version>-win64.zip
Contents: excel-converter.exe, README.md, README.txt

Run from the project root:
  python scripts/package.py
"""

import re
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).parent.parent  # project root


def get_version() -> str:
    src = (ROOT / "src" / "excel_converter" / "__init__.py").read_text(encoding="utf-8")
    m = re.search(r'^__version__\s*=\s*["\']([^"\']+)["\']', src, re.MULTILINE)
    if not m:
        sys.exit("Error: could not read __version__ from src/excel_converter/__init__.py")
    return m.group(1)


def build_exe() -> Path:
    subprocess.run(
        [sys.executable, "-m", "PyInstaller", "excel_converter.spec", "--clean", "--noconfirm"],
        check=True,
        cwd=ROOT,
    )
    exe = ROOT / "dist" / "excel-converter.exe"
    if not exe.exists():
        sys.exit(f"Error: expected PyInstaller output at {exe}")
    return exe


def main():
    version = get_version()
    print(f"Building excel-converter v{version}...")

    exe = build_exe()

    zip_name = f"excel-converter-v{version}-win64.zip"
    zip_path = ROOT / "dist" / zip_name
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(exe, arcname="excel-converter.exe")
        print(f"  added  excel-converter.exe  ({exe.stat().st_size // 1024:,} KB)")
        for doc in ("README.md", "README.txt"):
            zf.write(ROOT / doc, arcname=doc)
            print(f"  added  {doc}")

    print(f"\nPackage ready: {zip_path}")


if __name__ == "__main__":
    main()
