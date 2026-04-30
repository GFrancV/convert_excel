# Excel Legacy Converter

**Version:** 0.2.0 | **Author:** GFrancV

Converts legacy Excel files (`.xls`, Excel 97-2003 format) to modern `.xlsx` format.

Two conversion modes are available:

| Mode | Requires | Preserves |
|---|---|---|
| **Excel COM** *(default)* | Windows + Microsoft Excel installed | Everything — formatting, colors, merges, borders, fonts, formulas, charts |
| **Fallback** *(--no-excel)* | Python only | Data values only (text, numbers, dates, booleans) |

The COM mode is equivalent to opening each file in Excel and using **File → Save As → xlsx**.

---

## Requirements

**Core (always required):**
- Python 3.8+
- `xlrd >= 2.0.1`
- `openpyxl >= 3.1.0`

**Recommended (for full format preservation):**
- `pywin32 >= 306` — Windows only
- Microsoft Excel installed on the machine

```bash
pip install -r requirements.txt
```

---

## Usage

```
python convert_excel.py <input_dir> [output_dir] [options]
```

### Arguments

| Argument | Type | Description |
|---|---|---|
| `input_dir` | required | Folder containing `.xls` files to convert |
| `output_dir` | optional | Destination folder. Default: `<input_dir>/converted/` |
| `--no-excel` | flag | Skip Excel COM; convert data values only |
| `--recursive`, `-r` | flag | Search subdirectories recursively |
| `--workers N` | optional | Parallel threads for fallback mode. Default: CPU count |

---

## Examples

**Basic — convert all `.xls` files in a folder:**
```bash
python convert_excel.py ./test_files
# Output goes to: ./test_files/converted/
```

**Custom output folder:**
```bash
python convert_excel.py ./test_files ./output_xlsx
```

**Include subdirectories (recursive):**
```bash
python convert_excel.py ./test_files ./output_xlsx --recursive
```

**Data-only mode (no Excel required):**
```bash
python convert_excel.py ./test_files --no-excel
```

---

## Sample output

**With Excel COM (default):**
```
Found 6 file(s) to convert...  [Excel COM — full formatting]
Output: C:\data\test_files\converted

  [OK]   ventas_2003.xls
  [OK]   empleados.xls
  [OK]   inventario_multisheet.xls
  [OK]   norte.xls
  [OK]   sur.xls
  [OK]   global.xls

Done: 6 converted, 0 failed.
```

**Without Excel (fallback):**
```
Found 6 file(s) to convert...  [data only — no formatting]
Output: C:\data\test_files\converted

  [OK]   ventas_2003.xls
  [OK]   empleados.xls
  [OK]   inventario_multisheet.xls  [XML]
  [OK]   norte.xls
  [OK]   sur.xls  [HTML]
  [OK]   global.xls

Done: 6 converted, 0 failed.
```

The `[XML]` / `[HTML]` tags indicate files that were disguised as `.xls` but were actually SpreadsheetML or HTML — the fallback parser handles all three sub-formats automatically.

If a file cannot be read (corrupt, locked, wrong format), it is reported individually and the rest of the batch continues. The process exits with code `1` if any file failed.

---

## Test files

The `test_files/` folder contains sample `.xls` files ready to use:

```
test_files/
├── ventas_2003.xls             # Sales table: text, numbers, dates, currency
├── empleados.xls               # Employee list: mixed types, booleans, dates
├── inventario_multisheet.xls   # Inventory with 3 sheets (Electronica, Muebles, Papeleria)
└── por_region/                 # Subdirectory — use --recursive to include
    ├── norte.xls               # Regional sales: North
    ├── sur.xls                 # Regional sales: South
    └── internacional/          # Nested subdirectory (2 levels deep)
        └── global.xls          # International sales with currency data
```

---

## What is preserved

### Excel COM mode (default)

| Item | Preserved |
|---|---|
| Cell values (text, numbers, dates, booleans) | Yes |
| Formulas (live, not just values) | Yes |
| Cell formatting (colors, fonts, borders) | Yes |
| Merged cells | Yes |
| Column widths and row heights | Yes |
| Multiple sheets (all, original names) | Yes |
| Charts and images | Yes |
| Named ranges, print areas | Yes |
| Macros (VBA) | No — saved as `.xlsx` (macro-free) |

### Fallback mode (--no-excel)

| Item | Preserved |
|---|---|
| Cell values (text, numbers, dates, booleans) | Yes |
| Formulas | Values only (result, not formula) |
| Cell formatting / styles | No |
| Merged cells | No |
| Charts, images, macros | No |

---

## Supported .xls sub-formats (fallback mode)

Many tools save files with a `.xls` extension that are not true binary Excel files.
The fallback parser detects and handles all three variants automatically:

| Variant | Detection | Parser |
|---|---|---|
| OLE2 binary (true Excel 97-2003) | `D0 CF 11 E0` magic bytes | `xlrd` |
| SpreadsheetML (XML-based, Excel 2002-2003) | `<?xml` / `<Workbook>` in header | `xml.etree.ElementTree` |
| HTML table with `.xls` extension | everything else | `html.parser` |

---

## Notes

- Files already in `.xlsx` format are **not** processed (only `.xls` is targeted).
- Subdirectory structure is preserved in the output when using `--recursive`.
- In COM mode, Protected View dialogs (files marked as internet-origin) are handled automatically.
- `xlrd 2.x` intentionally supports only `.xls`. Do not downgrade to `xlrd 1.x`.
