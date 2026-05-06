================================================================================
  EXCEL LEGACY CONVERTER                                 v0.2.0 by GFrancV
================================================================================

  Converts legacy Excel files (.xls, Excel 97-2003) to modern .xlsx format.

  Available as a graphical application (GUI) and a command-line tool (CLI).

  Two conversion modes:

  COM mode (default)   Requires Windows + Microsoft Excel installed.
                       Equivalent to File -> Save As -> xlsx in Excel.
                       Preserves EVERYTHING: formatting, colors, merged cells,
                       borders, fonts, formulas, charts, column widths.

  Fallback mode        Pure Python, no Excel required (--no-excel flag).
  (--no-excel)         Preserves data values only. No formatting.


--------------------------------------------------------------------------------
  GUI (GRAPHICAL INTERFACE)
--------------------------------------------------------------------------------

  The graphical interface lets you select files or folders through native
  Windows dialogs and convert them with one click -- no command line needed.

  Launch:

    excel-converter-gui               (if installed via pip)
    python -m excel_converter.gui     (from source)

  Features:

  - Select files    Opens a Windows file picker; supports multiple .xls files.
  - Select folder   Opens a folder browser with optional "Include subfolders".
  - Destination     Choose an output folder via Browse, or leave blank to use
                    the default (<origin>/converted/).
  - Data-only mode  Toggle to convert without requiring Excel installed.
  - Live progress   Progress bar and per-file log with color-coded OK/FAIL.
  - Cancel          Stops after the current file; Excel is always cleaned up.

  The standalone excel-converter-gui.exe requires no Python on the target machine.


--------------------------------------------------------------------------------
  REQUIREMENTS
--------------------------------------------------------------------------------

  Core (always required):
  - Python 3.8 or higher
  - xlrd >= 2.0.1        (reads .xls binary files)
  - openpyxl >= 3.1.0    (writes .xlsx files)

  Recommended (for full format preservation):
  - pywin32 >= 306       (Windows COM automation -- Windows only)
  - Microsoft Excel installed on the machine


--------------------------------------------------------------------------------
  SETUP
--------------------------------------------------------------------------------

  Install from source:

    git clone https://github.com/GFrancV/excel-converter.git
    cd excel-converter
    pip install .

  After installation two commands are available globally:
    excel-converter       CLI tool
    excel-converter-gui   Graphical interface

  Developer install (editable, includes PyInstaller):

    pip install -e ".[dev]"


--------------------------------------------------------------------------------
  USAGE
--------------------------------------------------------------------------------

  excel-converter <input_dir> [output_dir] [options]

  ARGUMENTS
  ---------
  input_dir          (required) Folder containing .xls files to convert.

  output_dir         (optional) Destination folder.
                     Default: <input_dir>/converted/

  --no-excel         (flag) Skip Excel COM automation.
                     Converts data values only -- no formatting preserved.

  --recursive, -r    (flag) Search subdirectories recursively.

  --workers N        (optional) Parallel threads for fallback mode.
                     Default: number of CPU cores. Ignored in COM mode.

  --help, -h         Show help message and exit.

  --version          Show version number and exit.


--------------------------------------------------------------------------------
  EXAMPLES
--------------------------------------------------------------------------------

  Basic usage -- convert all .xls files in a folder:

    excel-converter ./test_files
    (Output goes to: ./test_files/converted/)

  Custom output folder:

    excel-converter ./test_files ./output_xlsx

  Include subdirectories (recursive):

    excel-converter ./test_files ./output_xlsx --recursive

  Data-only mode (no Excel required):

    excel-converter ./test_files --no-excel


--------------------------------------------------------------------------------
  SAMPLE OUTPUT
--------------------------------------------------------------------------------

  With Excel COM (default):

    Found 6 file(s) to convert...  [Excel COM -- full formatting]
    Output: C:\data\test_files\converted

      [OK]   ventas_2003.xls
      [OK]   empleados.xls
      [OK]   inventario_multisheet.xls
      [OK]   norte.xls
      [OK]   sur.xls
      [OK]   global.xls

    Done: 6 converted, 0 failed.

  Without Excel (fallback):

    Found 6 file(s) to convert...  [data only -- no formatting]
    Output: C:\data\test_files\converted

      [OK]   ventas_2003.xls
      [OK]   empleados.xls
      [OK]   inventario_multisheet.xls  [XML]
      [OK]   norte.xls
      [OK]   sur.xls  [HTML]
      [OK]   global.xls

    Done: 6 converted, 0 failed.

  [XML] and [HTML] tags identify files disguised as .xls that were actually
  SpreadsheetML (XML) or HTML -- the fallback parser handles all three variants.

  If a file is corrupt or unreadable it is reported individually while the rest
  continues. Exit code is 1 if any file failed.


--------------------------------------------------------------------------------
  BUILD (STANDALONE EXECUTABLES)
--------------------------------------------------------------------------------

  Produces self-contained Windows executables -- no Python required on target.

  Prerequisites: pip install -e ".[dev]"

  Build both executables:

    python scripts/package.py

  Output: dist/excel-converter-v<version>-win64.zip
  Contents:
    excel-converter.exe       CLI  (with console window)
    excel-converter-gui.exe   GUI  (no console window)
    README.md
    README.txt

  Build individually:

    pyinstaller excel_converter.spec --clean       (CLI only)
    pyinstaller excel_converter_gui.spec --clean   (GUI only)


--------------------------------------------------------------------------------
  TEST FILES
--------------------------------------------------------------------------------

  The test_files/ folder contains sample .xls files ready to use:

  test_files/
  |
  +-- ventas_2003.xls              Sales table: text, numbers, dates, currency
  +-- empleados.xls                Employee list: mixed types, booleans, dates
  +-- inventario_multisheet.xls    Inventory with 3 sheets
  |
  +-- por_region/                  Subdirectory (use --recursive to include)
      |
      +-- norte.xls                Regional sales: North
      +-- sur.xls                  Regional sales: South
      |
      +-- internacional/           Nested subdirectory (2 levels deep)
          |
          +-- global.xls           International sales with currency data


--------------------------------------------------------------------------------
  WHAT IS PRESERVED
--------------------------------------------------------------------------------

  COM MODE (default -- requires Excel)
  -------------------------------------
  Cell values (text, numbers, dates)    Yes
  Formulas (live)                       Yes
  Cell formatting (colors, fonts)       Yes
  Merged cells                          Yes
  Column widths and row heights         Yes
  Multiple sheets (original names)      Yes
  Charts and images                     Yes
  Named ranges, print areas             Yes
  Macros (VBA)                          No  (saved as .xlsx, macro-free)

  FALLBACK MODE (--no-excel)
  --------------------------
  Cell values (text, numbers, dates)    Yes
  Formulas                              Values only (result, not formula)
  Cell formatting / styles              No
  Merged cells                          No
  Charts, images, macros                No


--------------------------------------------------------------------------------
  SUPPORTED .XLS SUB-FORMATS  (fallback mode)
--------------------------------------------------------------------------------

  Many tools save files as .xls that are not true binary Excel files.
  The fallback parser detects and handles all three variants automatically:

  OLE2 binary (true Excel 97-2003)      Magic bytes D0 CF 11 E0    -> xlrd
  SpreadsheetML (XML-based Excel 2002)  <?xml / <Workbook> header  -> xml.etree
  HTML table with .xls extension        Everything else            -> html.parser


--------------------------------------------------------------------------------
  NOTES
--------------------------------------------------------------------------------

  - Files already in .xlsx format are NOT processed (only .xls is targeted).

  - Subdirectory structure is preserved in the output when using --recursive.

  - In COM mode, Protected View dialogs (files from internet) are handled
    automatically -- no manual interaction required.

  - xlrd 2.x intentionally supports only .xls. Do NOT downgrade to xlrd 1.x.

================================================================================
