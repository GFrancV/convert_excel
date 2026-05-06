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


def run_conversion(tasks, no_excel=False, workers=None):
    """Generator that executes conversion tasks and yields a progress dict per file.

    First yield — mode info (src is None):
        {"src": None, "success": None, "error": "", "fmt": "",
         "done": 0, "total": int, "mode": "com"|"fallback", "com_unavailable": bool}

    Subsequent yields — per-file result:
        {"src": Path, "success": bool, "error": str, "fmt": str,
         "done": int, "total": int, "mode": str, "com_unavailable": bool}

    Generator return value (StopIteration.value):
        {"ok": int, "failed": int, "failed_files": [(Path, str)]}
    """
    if workers is None:
        workers = os.cpu_count() or 4

    total = len(tasks)
    ok_count = 0
    failed = []
    excel = None if no_excel else _start_excel()
    com_unavailable = not no_excel and excel is None
    mode = "com" if excel is not None else "fallback"

    yield {
        "src": None, "success": None, "error": "", "fmt": "",
        "done": 0, "total": total, "mode": mode, "com_unavailable": com_unavailable,
    }

    try:
        if excel is not None:
            for done, (src, dst) in enumerate(tasks, start=1):
                try:
                    _convert_with_excel(excel, src, dst)
                    ok_count += 1
                    yield {"src": src, "success": True, "error": "", "fmt": "COM",
                           "done": done, "total": total, "mode": mode, "com_unavailable": False}
                except Exception as exc:
                    failed.append((src, str(exc)))
                    yield {"src": src, "success": False, "error": str(exc), "fmt": "",
                           "done": done, "total": total, "mode": mode, "com_unavailable": False}
        else:
            with ThreadPoolExecutor(max_workers=workers) as executor:
                futures = {
                    executor.submit(_convert_data_only, src, dst): src
                    for src, dst in tasks
                }
                done = 0
                for future in as_completed(futures):
                    src, success, error, fmt = future.result()
                    done += 1
                    if success:
                        ok_count += 1
                        yield {"src": src, "success": True, "error": "", "fmt": fmt,
                               "done": done, "total": total, "mode": mode, "com_unavailable": False}
                    else:
                        failed.append((src, error))
                        yield {"src": src, "success": False, "error": error, "fmt": "",
                               "done": done, "total": total, "mode": mode, "com_unavailable": False}
    finally:
        if excel is not None:
            _stop_excel(excel)

    return {"ok": ok_count, "failed": len(failed), "failed_files": failed}


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

    gen = run_conversion(tasks, no_excel=args.no_excel, workers=args.workers)
    start = next(gen)  # mode info event
    if start["com_unavailable"]:
        print("Note: Excel COM unavailable — using data-only fallback converter.")
    mode_label = "[Excel COM — full formatting]" if start["mode"] == "com" else "[data only — no formatting]"
    print(f"Found {len(files)} file(s) to convert...  {mode_label}")
    print(f"Output: {output_dir}\n")

    for p in gen:
        if p["success"]:
            ok_count += 1
            tag = f"  [{p['fmt'].upper()}]" if p["fmt"] in ("html", "xml") else ""
            print(f"  [OK]   {p['src'].name}{tag}")
        else:
            failed.append((p["src"], p["error"]))
            print(f"  [FAIL] {p['src'].name}: {p['error']}")

    print(f"\nDone: {ok_count} converted, {len(failed)} failed.")
    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
