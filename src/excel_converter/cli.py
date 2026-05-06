#!/usr/bin/env python3
"""Command-line entry point for Excel Legacy Converter."""

import argparse
import os
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

# Ensure UTF-8 output on Windows (avoids garbled chars in help/progress)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")

from excel_converter import __author__, __version__
from excel_converter.com_mode import _convert_with_excel, _start_excel, _stop_excel
from excel_converter.discovery import build_tasks, find_files
from excel_converter.fallback import _convert_data_only


def main():
    parser = argparse.ArgumentParser(
        description=f"Excel Legacy Converter v{__version__} by {__author__} — "
                    "convert .xls files (Excel 97-2003) to modern .xlsx format.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  excel-converter ./files\n"
            "  excel-converter ./files ./output\n"
            "  excel-converter ./files --recursive\n"
            "  excel-converter ./files --no-excel   # data-only fallback\n"
        ),
    )
    parser.add_argument("input_dir", help="Folder containing .xls files to convert")
    parser.add_argument(
        "output_dir",
        nargs="?",
        default=None,
        help="Destination folder (default: <input_dir>/converted)",
    )
    parser.add_argument(
        "--no-excel",
        action="store_true",
        help="Skip Excel COM automation; convert data values only (no formatting)",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=os.cpu_count() or 4,
        metavar="N",
        help="Parallel threads for fallback mode (default: CPU count)",
    )
    parser.add_argument(
        "--recursive",
        "-r",
        action="store_true",
        help="Search subdirectories recursively",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    args = parser.parse_args()

    input_dir = Path(args.input_dir).resolve()
    if not input_dir.is_dir():
        sys.exit(f"Error: '{input_dir}' is not a valid directory.")

    output_dir = (
        Path(args.output_dir).resolve() if args.output_dir else input_dir / "converted"
    )

    files = find_files(input_dir, args.recursive)
    if not files:
        print("No .xls files found in the specified path.")
        return

    tasks = build_tasks(files, input_dir, output_dir)
    ok_count = 0
    failed = []

    # ── COM mode (primary) ────────────────────────────────────────────────────
    excel = None if args.no_excel else _start_excel()

    if excel is not None:
        print(f"Found {len(files)} file(s) to convert...  [Excel COM — full formatting]")
        print(f"Output: {output_dir}\n")
        try:
            for src, dst in tasks:
                try:
                    _convert_with_excel(excel, src, dst)
                    ok_count += 1
                    print(f"  [OK]   {src.name}")
                except Exception as exc:
                    failed.append((src, str(exc)))
                    print(f"  [FAIL] {src.name}: {exc}")
        finally:
            _stop_excel(excel)

    # ── Fallback mode ─────────────────────────────────────────────────────────
    else:
        if not args.no_excel:
            print("Note: Excel COM unavailable — using data-only fallback converter.")
        print(f"Found {len(files)} file(s) to convert...  [data only — no formatting]")
        print(f"Output: {output_dir}\n")

        with ThreadPoolExecutor(max_workers=args.workers) as executor:
            futures = {
                executor.submit(_convert_data_only, src, dst): src
                for src, dst in tasks
            }
            for future in as_completed(futures):
                src, success, error, fmt = future.result()
                if success:
                    ok_count += 1
                    tag = f"  [{fmt.upper()}]" if fmt in ("html", "xml") else ""
                    print(f"  [OK]   {src.name}{tag}")
                else:
                    failed.append((src, error))
                    print(f"  [FAIL] {src.name}: {error}")

    print(f"\nDone: {ok_count} converted, {len(failed)} failed.")
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
