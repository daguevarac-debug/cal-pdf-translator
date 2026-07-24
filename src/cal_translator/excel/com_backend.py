from __future__ import annotations

import os
import signal
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Callable, Iterator, TypeVar

_T = TypeVar("_T")
_RPC_E_CALL_REJECTED = -2147418111
_RPC_E_SERVERCALL_RETRYLATER = -2147417846
_RETRYABLE_HRESULTS = {_RPC_E_CALL_REJECTED, _RPC_E_SERVERCALL_RETRYLATER}


def _com_hresult(exc: BaseException) -> int | None:
    value = getattr(exc, "hresult", None)
    if isinstance(value, int):
        return value
    args = getattr(exc, "args", ())
    return args[0] if args and isinstance(args[0], int) else None


def retry_com_call(
    operation: Callable[[], _T],
    *,
    attempts: int = 12,
    initial_delay: float = 0.15,
) -> _T:
    """Retry transient Excel COM rejections with bounded exponential backoff."""
    delay = initial_delay
    for attempt in range(1, attempts + 1):
        try:
            return operation()
        except Exception as exc:  # noqa: BLE001
            if _com_hresult(exc) not in _RETRYABLE_HRESULTS or attempt == attempts:
                raise
            time.sleep(delay)
            delay = min(delay * 1.5, 1.0)
    raise RuntimeError("Excel COM retry loop ended unexpectedly.")


def set_com_property(target: Any, name: str, value: Any) -> None:
    retry_com_call(lambda: setattr(target, name, value))


def _terminate_isolated_excel(process_id: int | None) -> None:
    if not process_id:
        return
    try:
        os.kill(process_id, signal.SIGTERM)
    except (OSError, ProcessLookupError):
        return
    time.sleep(0.8)


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
        import win32process
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 is required. Install the project with: python -m pip install -e ."
        ) from exc

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    process_id: int | None = None
    cleanup_failed = False
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        try:
            process_id = int(win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1])
        except Exception:  # noqa: BLE001
            process_id = None

        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.Interactive = False
        workbook = retry_com_call(
            lambda: excel.Workbooks.Open(
                str(workbook_path.resolve()),
                UpdateLinks=0,
                ReadOnly=read_only,
                AddToMru=False,
                IgnoreReadOnlyRecommended=True,
            )
        )
        yield excel, workbook
    finally:
        if workbook is not None:
            try:
                retry_com_call(lambda: workbook.Close(SaveChanges=False), attempts=4)
            except Exception:  # noqa: BLE001
                cleanup_failed = True
        if excel is not None:
            try:
                retry_com_call(excel.Quit, attempts=4)
            except Exception:  # noqa: BLE001
                cleanup_failed = True
        if cleanup_failed:
            _terminate_isolated_excel(process_id)
        pythoncom.CoUninitialize()
