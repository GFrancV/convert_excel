from pathlib import Path

# Excel SaveAs format constant for .xlsx (no macros)
_XLSX_FORMAT = 51  # xlOpenXMLWorkbook


def _start_excel():
    """Launch a hidden Excel instance via COM. Returns the Application object, or None."""
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        return excel
    except Exception:
        return None


def _stop_excel(excel) -> None:
    """Quit an Excel COM instance and release the COM apartment."""
    try:
        excel.Quit()
    except Exception:
        pass
    try:
        import pythoncom
        pythoncom.CoUninitialize()
    except Exception:
        pass


def _convert_with_excel(excel, src: Path, dst: Path) -> None:
    """
    Open *src* in the provided Excel instance and save as .xlsx to *dst*.

    Handles automatically:
    - Format-mismatch dialogs (DisplayAlerts = False)
    - Protected View (files with Zone.Identifier / internet-origin mark)
    """
    wb = None
    try:
        dst.parent.mkdir(parents=True, exist_ok=True)
        pv_before = excel.ProtectedViewWindows.Count

        wb = excel.Workbooks.Open(
            Filename=str(src.resolve()),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
            Notify=False,
        )

        # If the file was intercepted by Protected View, exit it to get
        # a writable workbook reference before calling SaveAs.
        if excel.ProtectedViewWindows.Count > pv_before:
            pv = excel.ProtectedViewWindows.Item(excel.ProtectedViewWindows.Count)
            wb = pv.Edit()

        wb.SaveAs(
            Filename=str(dst.resolve()),
            FileFormat=_XLSX_FORMAT,
            CreateBackup=False,
        )
        wb.Close(SaveChanges=False)
        wb = None

    except Exception:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        for _ in range(excel.ProtectedViewWindows.Count):
            try:
                excel.ProtectedViewWindows.Item(1).Close()
            except Exception:
                break
        raise
