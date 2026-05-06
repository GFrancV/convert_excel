# excel-converter — Project Context

## What this is

A CLI tool that converts legacy Excel files (`.xls`, Excel 97-2003) to modern `.xlsx` format. The goal is to replicate **File → Save As → .xlsx** in Excel: full format preservation including colors, fonts, borders, merged cells, formulas, and charts.

## Project structure

```
excel-converter/
├── src/
│   └── excel_converter/
│       ├── __init__.py          # __version__, __author__
│       ├── __main__.py          # python -m excel_converter support
│       ├── cli.py               # argparse + main() entry point
│       ├── com_mode.py          # Excel COM conversion (primary mode)
│       ├── fallback.py          # Pure-Python fallback (data-only)
│       └── discovery.py         # find_files(), build_tasks()
├── scripts/
│   └── package.py               # builds standalone .exe ZIP via PyInstaller
├── tests/                       # test suite
├── test_files/                  # Sample .xls files for manual testing
│   ├── ventas_2003.xls
│   ├── empleados.xls
│   ├── inventario_multisheet.xls
│   └── por_region/
│       ├── norte.xls
│       ├── sur.xls
│       └── internacional/
│           └── global.xls
├── lolo/                        # Real-world test file (SpreadsheetML format)
│   ├── PFXX - Lista de conjuntos.xls     # Source (SpreadsheetML XML with .xls ext)
│   └── PFXX - Lista de conjuntos.xlsx    # Reference output (manual Save As)
├── excel_converter.spec         # PyInstaller build spec
└── pyproject.toml
```

## Architecture

Two completely independent conversion paths:

### Primary: Excel COM mode (default)
Requires Windows + Microsoft Excel installed + `pywin32`. Lives in `src/excel_converter/com_mode.py`.
- `_start_excel()` — launches a hidden `Excel.Application` via `DispatchEx`
- `_convert_with_excel(excel, src, dst)` — calls `Workbooks.Open` then `SaveAs(FileFormat=51)`
- `_stop_excel(excel)` — calls `excel.Quit()` + `CoUninitialize()`
- Single shared Excel instance processes all files sequentially
- Handles Protected View (Zone.Identifier / internet-origin files) via `ProtectedViewWindows.Edit()`

### Fallback: Pure-Python mode (`--no-excel`)
No Excel required. Data values only, no formatting. Lives in `src/excel_converter/fallback.py`.
- `_detect_format(path)` — reads first 1024 bytes, returns `'xls'` / `'xml'` / `'html'`
- `_sheets_from_xls(path)` — OLE2 binary via `xlrd`
- `_sheets_from_xml(path)` — SpreadsheetML XML via `xml.etree.ElementTree`
- `_sheets_from_html(path)` — HTML tables via `html.parser`
- `_convert_data_only(src, dst)` — detects format, reads sheets, writes via `openpyxl`
- Uses `ThreadPoolExecutor` for parallel processing (`--workers N`)

### Key constants (by module)
- `com_mode.py`: `_XLSX_FORMAT = 51` — Excel COM constant for `xlOpenXMLWorkbook` (.xlsx, no macros)
- `fallback.py`: `_OLE2_SIGNATURE = b"\xd0\xcf\x11\xe0"` — magic bytes of every genuine XLS binary
- `fallback.py`: `_SS` dict — pre-built namespace-qualified XML tag/attribute names for SpreadsheetML
- `discovery.py`: `OLD_FORMATS = {".xls"}` — extensions targeted for conversion

## Known .xls sub-formats

Many files have `.xls` extension but are not true binary Excel files:

| Variant | Detection | Handled by |
|---|---|---|
| OLE2 binary (true XLS) | `D0 CF 11 E0` magic bytes | `xlrd` |
| SpreadsheetML (XML) | `<?xml` or `<Workbook` in first 1024 bytes | `xml.etree.ElementTree` |
| HTML table | everything else | `html.parser` |

**Real-world quirk encountered:** `PFXX - Lista de conjuntos.xls` in `lolo/` is SpreadsheetML XML with a `.xls` extension, contains HTML comments before the `<?xml` declaration, and has `\x0c` (form feed) control characters in the body — illegal in XML 1.0. The parser strips these before parsing.

## Setup and dependencies

```bash
# End-user install
pip install .

# Developer install (includes PyInstaller)
pip install -e ".[dev]"
```

Core deps: `xlrd>=2.0.1`, `openpyxl>=3.1.0`, `pywin32>=306` (Windows only)

## Testing

```bash
# COM mode — full formatting (requires Excel)
excel-converter test_files --recursive

# Fallback mode — data only, parallel
excel-converter test_files --no-excel --recursive

# Real-world SpreadsheetML file
excel-converter lolo

# Custom output folder
excel-converter test_files ./output --recursive
```

Outputs go to `<input_dir>/converted/` by default, mirroring subdirectory structure.

## Build

```bash
python scripts/package.py
# Output: dist/excel-converter-v<version>-win64.zip
```

## Version history

| Version | Changes |
|---|---|
| 0.1.0 | Initial release — data-only conversion via xlrd + openpyxl |
| 0.1.1 | Added SpreadsheetML (XML) and HTML sub-format detection and parsing |
| 0.2.0 | Added Excel COM mode as primary path for full format preservation |
| 0.2.1 | Migrated to src layout — split into focused modules under src/excel_converter/ |
