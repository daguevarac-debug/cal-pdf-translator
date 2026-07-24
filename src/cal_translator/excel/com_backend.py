from __future__ import annotations

import os
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator


@contextmanager
def open_excel_workbook(
    workbook_path: Path,
    *,
    read_only: bool,
) -> Iterator[tuple[Any, Any]]:
    """Open an Excel workbook in an isolated COM process and close it safely."""
    if os.name != "nt":
        raise RuntimeError("Excel COM automation requires Windows with Microsoft Excel installed.")
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook file not found: {workbook_path}")

    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required. Install the project with: python -m pip install -e ."
        ) from exc

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        workbook = excel.Workbooks.Open(
            str(workbook_path.resolve()),
            UpdateLinks=0,
            ReadOnly=read_only,
            AddToMru=False,
            IgnoreReadOnlyRecommended=True,
        )
        yield excel, workbook
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
