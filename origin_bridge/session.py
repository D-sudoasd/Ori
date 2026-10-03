"""One Origin process owned by this program, reused across batch tasks.

The session refuses an Origin process it did not start. Ownership is the
process that holds a sentinel file opened by this COM instance, checked with
the Windows Restart Manager, then pinned by a process handle and its creation
time. A timeout or a forced stop terminates only that handle. An unproven
process is never terminated. This module never issues a global ``taskkill``.
"""

from __future__ import annotations

import ctypes
import hashlib
import os
import re
import shutil
import sys
import tempfile
import threading
import time
from pathlib import Path
from types import SimpleNamespace
from typing import Any

from .models import DataImportError, PlannedTable, PreparedImport

_PAGE_RE = re.compile(br"/Type\s*/Page(?!s)\b")
_ORIGIN_IMAGES = frozenset({"origin64.exe", "origin.exe"})


class OriginSessionLost(RuntimeError):
    """The Origin process started for this session exited, hung, or never appeared."""


def _is_live_origin(op: Any) -> bool:
    name = str(getattr(op, "__name__", "") or "")
    module = str(getattr(op, "__module__", "") or "")
    return name == "originpro" or name.startswith("originpro.") or module.startswith("originpro")


class _PROCESSENTRY32W(ctypes.Structure):
    _fields_ = [
        ("dwSize", ctypes.c_ulong),
        ("cntUsage", ctypes.c_ulong),
        ("th32ProcessID", ctypes.c_ulong),
        ("th32DefaultHeapID", ctypes.c_void_p),
        ("th32ModuleID", ctypes.c_ulong),
        ("cntThreads", ctypes.c_ulong),
        ("th32ParentProcessID", ctypes.c_ulong),
        ("pcPriClassBase", ctypes.c_long),
        ("dwFlags", ctypes.c_ulong),
        ("szExeFile", ctypes.c_wchar * 260),
    ]


_TOOLHELP_LOCK = threading.Lock()
_TOOLHELP = None


def _toolhelp():
    """One private kernel32 binding.

    A shared ``windll`` prototype plus a Structure class created on every call
    races when the batch thread and the memory sampler enumerate processes
    together. Python then rejects the pointer as the wrong PROCESSENTRY32W.
    """
    global _TOOLHELP
    if _TOOLHELP is None:
        library = ctypes.WinDLL("kernel32", use_last_error=True)
        library.CreateToolhelp32Snapshot.argtypes = [ctypes.c_ulong, ctypes.c_ulong]
        library.CreateToolhelp32Snapshot.restype = ctypes.c_void_p
        library.Process32FirstW.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        library.Process32FirstW.restype = ctypes.c_int
        library.Process32NextW.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        library.Process32NextW.restype = ctypes.c_int
        library.CloseHandle.argtypes = [ctypes.c_void_p]
        library.CloseHandle.restype = ctypes.c_int
        _TOOLHELP = library
    return _TOOLHELP


def origin_pids() -> set[int]:
    """Return process ids whose image is Origin64.exe or Origin.exe."""
    if sys.platform != "win32":
        return set()
    try:
        return _origin_pids_windows()
    except OSError:
        return set()


def _origin_pids_windows() -> set[int]:
    TH32CS_SNAPPROCESS = 0x00000002
    invalid = ctypes.c_void_p(-1).value
    with _TOOLHELP_LOCK:
        kernel32 = _toolhelp()
        snap = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        if not snap or snap == invalid:
            raise OSError("CreateToolhelp32Snapshot failed")
        found: set[int] = set()
        try:
            entry = _PROCESSENTRY32W()
            entry.dwSize = ctypes.sizeof(_PROCESSENTRY32W)
            if not kernel32.Process32FirstW(snap, ctypes.byref(entry)):
                return found
            while True:
                name = str(entry.szExeFile).casefold()
                if name in _ORIGIN_IMAGES:
                    found.add(int(entry.th32ProcessID))
                if not kernel32.Process32NextW(snap, ctypes.byref(entry)):
                    break
        finally:
            kernel32.CloseHandle(snap)
        return found


class _FILETIME(ctypes.Structure):
    _fields_ = [
        ("dwLowDateTime", ctypes.c_uint),
        ("dwHighDateTime", ctypes.c_uint),
    ]


def _filetime_int(value: _FILETIME) -> int:
    return (int(value.dwHighDateTime) << 32) | int(value.dwLowDateTime)


_KERNEL_LOCK = threading.Lock()
_KERNEL = None
_PROCESS_QUERY = 0x1000
_PROCESS_TERMINATE = 0x0001
_STILL_ACTIVE = 259
_OWNERSHIP_ERROR = (
    "Could not prove which Origin process belongs to this COM instance; "
    "no candidate process was terminated"
)


def _kernel32():
    """Private kernel32 binding for process handles.

    Prototypes stay on this object. Callers do not use the shared ``windll``
    binding, and they do not build a new Structure class per call.
    """
    global _KERNEL
    with _KERNEL_LOCK:
        if _KERNEL is None:
            library = ctypes.WinDLL("kernel32", use_last_error=True)
            library.OpenProcess.argtypes = [ctypes.c_uint, ctypes.c_int, ctypes.c_uint]
            library.OpenProcess.restype = ctypes.c_void_p
            library.TerminateProcess.argtypes = [ctypes.c_void_p, ctypes.c_uint]
            library.TerminateProcess.restype = ctypes.c_int
            library.CloseHandle.argtypes = [ctypes.c_void_p]
            library.CloseHandle.restype = ctypes.c_int
            library.GetProcessTimes.argtypes = [
                ctypes.c_void_p,
                ctypes.POINTER(_FILETIME),
                ctypes.POINTER(_FILETIME),
                ctypes.POINTER(_FILETIME),
                ctypes.POINTER(_FILETIME),
            ]
            library.GetProcessTimes.restype = ctypes.c_int
            library.GetExitCodeProcess.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong)]
            library.GetExitCodeProcess.restype = ctypes.c_int
            library.QueryFullProcessImageNameW.argtypes = [
                ctypes.c_void_p,
                ctypes.c_uint,
                ctypes.c_wchar_p,
                ctypes.POINTER(ctypes.c_uint),
            ]
            library.QueryFullProcessImageNameW.restype = ctypes.c_int
            _KERNEL = library
        return _KERNEL


def terminate_pid(pid: int) -> bool:
    """Stop one process id. This does not look up Origin by image name."""
    if sys.platform != "win32" or isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0:
        return False
    kernel32 = _kernel32()
    handle = kernel32.OpenProcess(_PROCESS_TERMINATE, 0, int(pid))
    if not handle:
        return False
    try:
        return bool(kernel32.TerminateProcess(handle, 1))
    finally:
        kernel32.CloseHandle(handle)


def handle_status(handle) -> dict[str, Any] | None:
    """Creation time and liveness of a process handle this program already holds."""
    if sys.platform != "win32" or not handle:
        return None
    kernel32 = _kernel32()
    created = _FILETIME()
    exited = _FILETIME()
    kernel_time = _FILETIME()
    user_time = _FILETIME()
    if not kernel32.GetProcessTimes(
        handle, ctypes.byref(created), ctypes.byref(exited), ctypes.byref(kernel_time), ctypes.byref(user_time),
    ):
        return None
    code = ctypes.c_ulong()
    alive = True
    if kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
        alive = int(code.value) == _STILL_ACTIVE
    return {"creation": _filetime_int(created), "alive": alive}


def pid_creation_filetime(pid: int) -> int | None:
    """Creation time of whatever process currently owns ``pid``.

    The temporary handle opened here is closed before this function returns.
    """
    if sys.platform != "win32" or isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0:
        return None
    kernel32 = _kernel32()
    handle = kernel32.OpenProcess(_PROCESS_QUERY, 0, int(pid))
    if not handle:
        return None
    try:
        status = handle_status(handle)
    finally:
        kernel32.CloseHandle(handle)
    if not status:
        return None
    return int(status["creation"])


def process_image_name(pid: int) -> str:
    if sys.platform != "win32" or isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0:
        return ""
    kernel32 = _kernel32()
    handle = kernel32.OpenProcess(_PROCESS_QUERY, 0, int(pid))
    if not handle:
        return ""
    try:
        size = ctypes.c_uint(32768)
        buffer = ctypes.create_unicode_buffer(size.value)
        if not kernel32.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
            return ""
        return str(buffer.value)
    finally:
        kernel32.CloseHandle(handle)


def terminate_process_handle(handle) -> bool:
    """Stop the process behind a handle. This does not open a process id."""
    if sys.platform != "win32" or not handle:
        return False
    return bool(_kernel32().TerminateProcess(handle, 1))


def close_process_handle(handle) -> None:
    if sys.platform != "win32" or not handle:
        return
    _kernel32().CloseHandle(handle)


def _origin_image(pid: int, image: str) -> bool:
    """True only when the live process image is Origin64.exe or Origin.exe.

    ``image`` is the Restart Manager friendly name. That name is not an
    executable path and is not accepted as proof.
    """
    del image
    return Path(process_image_name(pid)).name.casefold() in _ORIGIN_IMAGES


class _RM_UNIQUE_PROCESS(ctypes.Structure):
    _fields_ = [
        ("dwProcessId", ctypes.c_uint),
        ("ProcessStartTime", _FILETIME),
    ]


class _RM_PROCESS_INFO(ctypes.Structure):
    _fields_ = [
        ("Process", _RM_UNIQUE_PROCESS),
        ("strAppName", ctypes.c_wchar * 256),
        ("strServiceShortName", ctypes.c_wchar * 64),
        ("ApplicationType", ctypes.c_uint),
        ("AppStatus", ctypes.c_uint),
        ("TSSessionId", ctypes.c_uint),
        ("bRestartable", ctypes.c_int),
    ]


_RM_LOCK = threading.Lock()
_RSTRTMGR = None


def _rstrtmgr():
    """One private Restart Manager binding. Prototypes are set once."""
    global _RSTRTMGR
    with _RM_LOCK:
        if _RSTRTMGR is None:
            library = ctypes.WinDLL("rstrtmgr", use_last_error=True)
            library.RmStartSession.argtypes = [ctypes.POINTER(ctypes.c_uint), ctypes.c_uint, ctypes.c_wchar_p]
            library.RmStartSession.restype = ctypes.c_uint
            library.RmRegisterResources.argtypes = [
                ctypes.c_uint,
                ctypes.c_uint,
                ctypes.POINTER(ctypes.c_wchar_p),
                ctypes.c_uint,
                ctypes.c_void_p,
                ctypes.c_uint,
                ctypes.c_void_p,
            ]
            library.RmRegisterResources.restype = ctypes.c_uint
            library.RmGetList.argtypes = [
                ctypes.c_uint,
                ctypes.POINTER(ctypes.c_uint),
                ctypes.POINTER(ctypes.c_uint),
                ctypes.c_void_p,
                ctypes.POINTER(ctypes.c_uint),
            ]
            library.RmGetList.restype = ctypes.c_uint
            library.RmEndSession.argtypes = [ctypes.c_uint]
            library.RmEndSession.restype = ctypes.c_uint
            _RSTRTMGR = library
        return _RSTRTMGR


def sentinel_holders(path: Path) -> list[dict[str, Any]]:
    """Processes that have ``path`` open, according to the Restart Manager.

    Each item is ``pid``, ``creation_filetime``, and ``image`` (the Restart
    Manager application name, which is not an executable path and is not
    proof by itself).
    """
    if sys.platform != "win32":
        return []
    library = _rstrtmgr()
    session = ctypes.c_uint()
    key = ctypes.create_unicode_buffer(33)
    if library.RmStartSession(ctypes.byref(session), 0, key) != 0:
        raise OSError(f"RmStartSession failed for {path}")
    try:
        files = (ctypes.c_wchar_p * 1)(str(path))
        if library.RmRegisterResources(session, 1, files, 0, None, 0, None) != 0:
            raise OSError(f"RmRegisterResources failed for {path}")
        needed = ctypes.c_uint(0)
        count = ctypes.c_uint(0)
        reasons = ctypes.c_uint(0)
        status = library.RmGetList(session, ctypes.byref(needed), ctypes.byref(count), None, ctypes.byref(reasons))
        if status not in (0, 234):
            raise OSError(f"RmGetList failed for {path}: {status}")
        if needed.value == 0:
            return []
        array = (_RM_PROCESS_INFO * int(needed.value))()
        count = ctypes.c_uint(array._length_)
        status = library.RmGetList(
            session, ctypes.byref(needed), ctypes.byref(count), ctypes.cast(array, ctypes.c_void_p), ctypes.byref(reasons),
        )
        if status == 234 and needed.value > array._length_:
            array = (_RM_PROCESS_INFO * int(needed.value))()
            count = ctypes.c_uint(array._length_)
            status = library.RmGetList(
                session,
                ctypes.byref(needed),
                ctypes.byref(count),
                ctypes.cast(array, ctypes.c_void_p),
                ctypes.byref(reasons),
            )
        if status != 0:
            raise OSError(f"RmGetList failed for {path}: {status}")
        holders = []
        for index in range(int(count.value)):
            item = array[index]
            holders.append({
                "pid": int(item.Process.dwProcessId),
                "creation_filetime": _filetime_int(item.Process.ProcessStartTime),
                "image": str(item.strAppName),
            })
        return holders
    finally:
        library.RmEndSession(session)


def claim_sentinel_holder(holders: list[dict[str, Any]], *, is_origin_image) -> dict[str, Any] | None:
    """Return the only process holding the sentinel when it is an Origin image.

    Zero holders, more than one holder, or a holder that is not Origin yields
    None. The caller must not turn that into a guess about other processes.
    """
    if len(holders) != 1:
        return None
    only = holders[0]
    pid = only.get("pid")
    creation = only.get("creation_filetime")
    if isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0:
        return None
    if isinstance(creation, bool) or not isinstance(creation, int) or creation <= 0:
        return None
    image = str(only.get("image") or "")
    if not is_origin_image(pid, image):
        return None
    return {"pid": int(pid), "creation_filetime": int(creation), "image": image}


def _same_path(reported: str, expected: Path) -> bool:
    if not reported.strip():
        return False
    left = os.path.normcase(str(Path(reported).expanduser().resolve(strict=False)))
    right = os.path.normcase(str(expected.expanduser().resolve(strict=False)))
    return left == right


def _hold_origin_sentinel(op, path: Path) -> None:
    """Ask this COM instance to keep ``path`` open.

    ``file.filename$``, ``file.mode`` and ``file.open()`` are the documented
    LabTalk file object. The path is stored with ``set_lt_str`` so backslashes
    are not parsed as LabTalk escapes. A rejected script, or a reported name
    that is a different file, is not ownership. Restart Manager then has to
    show that this file is held by exactly one Origin image.
    """
    target = str(path)
    if not op.set_lt_str("file.filename$", target):
        if not op.set_lt_str("__ORI_SENTINEL", target):
            raise DataImportError(f"{_OWNERSHIP_ERROR} (could not store the sentinel path)")
        if op.lt_exec("file.filename$=__ORI_SENTINEL$;") is False:
            raise DataImportError(f"{_OWNERSHIP_ERROR} (could not copy the sentinel path)")
    if op.lt_exec("file.mode=2; file.open();") is False:
        raise DataImportError(f"{_OWNERSHIP_ERROR} (file.open was rejected)")
    reported = str(op.get_lt_str("file.filename$") or "")
    try:
        held = int(op.lt_int("file.hFile") or 0)
    except (TypeError, ValueError):
        held = 0
    if reported.strip() and not _same_path(reported, path):
        raise DataImportError(
            f"{_OWNERSHIP_ERROR} (file.filename$={reported!r}, file.hFile={held})"
        )
    if held <= 0 and not reported.strip():
        raise DataImportError(f"{_OWNERSHIP_ERROR} (file.open left no filename or handle)")


def _close_origin_sentinel(op) -> None:
    try:
        closer = getattr(op, "lt_exec", None)
        if callable(closer):
            closer("file.close();")
    except Exception:
        return


class OwnedProcess:
    """A process handle plus the creation time captured when it was proved."""

    def __init__(self, pid: int, creation_filetime: int, handle: int, image: str) -> None:
        self.pid = int(pid)
        self.creation_filetime = int(creation_filetime)
        self.handle = handle
        self.image = image

    def public(self) -> dict[str, Any]:
        return {
            "pid": self.pid,
            "creation_filetime": self.creation_filetime,
            "image": self.image,
        }


def validate_pdf_file(path: Path) -> dict[str, Any]:
    """Require a PDF header, trailer, cross-reference, and an uncompressed page.

    The page check is a byte search for an uncompressed ``/Type /Page``.
    Object streams (``/ObjStm``) and page dictionaries that exist only inside
    compressed streams are rejected. This checker does not decode them.
    """
    data = Path(path).read_bytes()
    if len(data) < 32 or not data.startswith(b"%PDF-"):
        raise DataImportError(f"PDF header missing: {path}")
    if b"%%EOF" not in data[-4096:] and b"%%EOF" not in data:
        raise DataImportError(f"PDF trailer missing: {path}")
    if b"startxref" not in data and b"/XRef" not in data:
        raise DataImportError(f"PDF cross-reference missing: {path}")
    pages = _PAGE_RE.findall(data)
    if not pages:
        if b"/ObjStm" in data:
            raise DataImportError(
                f"PDF stores page objects in an ObjStm; this checker only accepts an uncompressed /Type /Page: {path}"
            )
        raise DataImportError(
            f"PDF has no uncompressed /Type /Page object; compressed page objects are not accepted: {path}"
        )
    return {
        "size_bytes": len(data),
        "pages": len(pages),
        "header": data[:8].decode("ascii", "replace"),
    }


def _graph_identity(graph: Any) -> tuple[str, str]:
    short = str(getattr(graph, "name", "") or "")
    obj = getattr(graph, "obj", None)
    getter = getattr(obj, "GetName", None) if obj is not None else None
    if callable(getter):
        try:
            found = getter()
            if found:
                short = str(found)
        except Exception:
            pass
    return short, str(getattr(graph, "lname", "") or "")


def _save_graph_pdf(graph: Any, dest: Path, planned: PlannedTable) -> dict[str, Any]:
    """Export one graph to a new path and require a real PDF there.

    ``GPage.save_fig`` can return an empty string, can report success while an
    older file is left in place, and can name a path that is not the graph we
    asked it to export. The destination must not exist beforehand. An empty
    return is accepted only when that new path becomes a valid PDF and the
    graph's identity is unchanged.
    """
    if dest.exists():
        raise DataImportError(f"Refusing to export a PDF over an existing file: {dest}")
    identity = _graph_identity(graph)
    returned = graph.save_fig(str(dest), replace=False)
    if _graph_identity(graph) != identity:
        raise DataImportError(
            f"Graph identity changed while exporting {planned.name}: {identity!r}"
        )
    if not dest.is_file() or dest.stat().st_size <= 0:
        raise DataImportError(
            f"save_fig did not create a non-empty PDF at {dest} (returned {returned!r}). "
            "An empty return value is not treated as success."
        )
    if isinstance(returned, str) and returned.strip():
        reported = Path(returned).expanduser().resolve(strict=False)
        if reported != dest.resolve(strict=False):
            raise DataImportError(f"save_fig reported {returned}, not the requested path {dest}")
    info = validate_pdf_file(dest)
    _short, long_name = identity
    return {
        "table": planned.name,
        "graph": long_name or planned.plot.title or planned.name,
        "short_name": _short,
        "pages": info["pages"],
        "error": None,
    }


def iter_project_sheets(books) -> list[Any]:
    sheets = []
    for book in books:
        for sheet in book:
            sheets.append(sheet)
    return sheets


def provenance_text(sheets) -> str:
    chunks: list[str] = []
    for sheet in sheets:
        if str(getattr(sheet, "lname", "") or "") != "Import Provenance":
            continue
        cols = int(getattr(sheet, "cols", 0) or 0)
        for index in range(cols):
            chunks.extend(str(value) for value in sheet.to_list(index))
    return "\n".join(chunks)


def project_graph_for(sheets, graphs, expected: dict[str, Any]):
    """Match one sheet and one graph. Ambiguous or missing identities raise."""
    table = str(expected.get("table") or "")
    title = str(expected.get("graph") or "")
    short = str(expected.get("short_name") or "")
    if not table or not title:
        raise DataImportError("PDF recovery is missing the recorded table or graph name")
    sheet_hits = [sheet for sheet in sheets if str(getattr(sheet, "lname", "") or "") == table]
    graph_hits = []
    for graph in graphs:
        found_short, found_long = _graph_identity(graph)
        if short and found_short != short:
            continue
        if found_long != title:
            continue
        graph_hits.append(graph)
    if len(sheet_hits) != 1 or len(graph_hits) != 1:
        raise DataImportError(
            f"Refusing to export a PDF without a unique table/graph match for {table!r} / {title!r}"
            + (f" / {short!r}" if short else "")
        )
    return sheet_hits[0], graph_hits[0]


class OriginSession:
    """A single Origin instance created by this process."""

    launch_count = 0
    close_count = 0

    def __init__(self, timeout_s: float | None = None) -> None:
        self.timeout_s = float(timeout_s or 0)
        self.op: Any = None
        self.started = False
        self.starts = 0
        self.stops = 0
        self.owned_pids: set[int] = set()
        self._owned: list[OwnedProcess] = []
        self.pending_cleanup: list[Path] = []
        self._before: set[int] = set()
        self._live = False
        self._alive = False
        self._closed = False
        self._deadline: float | None = None
        self._timed_out = False
        self._stop_watchdog = threading.Event()
        self._watchdog: threading.Thread | None = None
        self._pid_creation = pid_creation_filetime
        self._handle_status = handle_status
        self._terminate_handle = terminate_process_handle
        self._close_handle = close_process_handle

    @property
    def owned_identities(self) -> list[dict[str, Any]]:
        """Proved processes for this session, including ones already closed.

        ``owned_pids`` is the same set of process ids. ``creation_filetime`` is
        the creation time of the held handle. A later process that reuses the
        id has a different creation time and is not this session.
        """
        return [item.public() for item in self._owned]

    def healthy(self) -> bool:
        if self._closed or not self.started or not self._alive or self.op is None:
            return False
        if not self._live:
            return True
        return any(self._identity_ok(item) for item in self._owned)

    def start(self) -> None:
        if self.started and self.healthy():
            return
        from . import exporter as ex

        sto = ex._origin_support()
        if sto.origin_process_running():
            raise sto.OriginExportError(
                "Origin is already running; close it before exporting a new project"
            )
        op = sto._load_originpro()
        if sto.origin_process_running():
            raise sto.OriginExportError(
                "Origin started during export setup; the current session was left untouched"
            )
        self._live = _is_live_origin(op)
        self._before = origin_pids() if self._live else set()
        if self._live and self._before:
            raise sto.OriginExportError(
                "Origin is already running; refusing to attach to or modify the existing session"
            )
        self.op = op
        self.pending_cleanup = []
        self._closed = False
        self._stop_watchdog = threading.Event()
        launched = False
        try:
            op.set_show(False)
            launched = True
            if self._live:
                self._bind_owned_process(op)
                if not self._owned:
                    raise DataImportError(_OWNERSHIP_ERROR)
            self.started = True
            self._alive = True
            self.starts += 1
            OriginSession.launch_count += 1
            if self._live and self.timeout_s > 0:
                self._watchdog = threading.Thread(
                    target=self._watch,
                    name="origin-session-watchdog",
                    daemon=True,
                )
                self._watchdog.start()
        except BaseException:
            self._stop_watchdog.set()
            if self._owned:
                self.started = True
                self._alive = True
                self.close()
            else:
                if launched and self.op is not None:
                    try:
                        sto.close_origin_app(self.op, started=True)
                    except Exception:
                        pass
                self._release_handles()
                self.op = None
                self.started = False
                self._alive = False
            raise

    def close(self) -> None:
        if self._closed:
            return
        self._stop_watchdog.set()
        if (
            self._watchdog is not None
            and self._watchdog.is_alive()
            and threading.current_thread() is not self._watchdog
        ):
            self._watchdog.join(timeout=1)
        try:
            if not self.started:
                return
            from . import exporter as ex

            sto = ex._origin_support()
            try:
                if self.op is not None:
                    sto.close_origin_app(self.op, started=True)
            except Exception:
                pass
            if self._live:
                self._wait_until_owned_exit()
            sto._cleanup_origin_temp_dirs(self.pending_cleanup)
            self.pending_cleanup = []
            self.stops += 1
            OriginSession.close_count += 1
        finally:
            self._release_handles()
            self.started = False
            self._alive = False
            self._closed = True
            self.op = None

    def terminate_owned(self) -> list[int]:
        """Stop proved processes whose handle and process id still match.

        An empty ownership list stops nothing, even if a new Origin process
        is visible. A reused process id is not terminated.
        """
        killed: list[int] = []
        for owned in self._owned:
            if not owned.handle or not self._identity_ok(owned):
                continue
            if self._terminate_handle(owned.handle):
                killed.append(owned.pid)
        return killed

    def _identity_ok(self, owned: OwnedProcess) -> bool:
        if not owned.handle:
            return False
        status = self._handle_status(owned.handle)
        if not status or not status.get("alive"):
            return False
        if int(status.get("creation") or -1) != int(owned.creation_filetime):
            return False
        current = self._pid_creation(owned.pid)
        if current is None or int(current) != int(owned.creation_filetime):
            return False
        return True

    def _release_handles(self) -> None:
        for owned in self._owned:
            handle = owned.handle
            owned.handle = None
            if not handle:
                continue
            try:
                self._close_handle(handle)
            except Exception:
                continue

    def _bind_owned_process(self, op) -> None:
        directory = Path(tempfile.mkdtemp(prefix="ori_owner_"))
        sentinel = directory / "sentinel.bin"
        sentinel.write_bytes(b"ori-owner")
        try:
            try:
                _hold_origin_sentinel(op, sentinel)
                holders = sentinel_holders(sentinel)
            finally:
                _close_origin_sentinel(op)
            claimed = claim_sentinel_holder(holders, is_origin_image=_origin_image)
            if claimed is None:
                detail = []
                for item in holders:
                    pid = item.get("pid")
                    image = process_image_name(pid) if isinstance(pid, int) and pid > 0 else ""
                    detail.append(f"{pid}:{Path(image).name}")
                raise DataImportError(f"{_OWNERSHIP_ERROR} (sentinel holders={detail})")
            handle = _kernel32().OpenProcess(_PROCESS_QUERY | _PROCESS_TERMINATE, 0, int(claimed["pid"]))
            if not handle:
                raise DataImportError(_OWNERSHIP_ERROR)
            status = handle_status(handle)
            if (
                not status
                or not status.get("alive")
                or int(status["creation"]) != int(claimed["creation_filetime"])
            ):
                close_process_handle(handle)
                raise DataImportError(_OWNERSHIP_ERROR)
            image = process_image_name(claimed["pid"]) or str(claimed.get("image") or "")
            if Path(image).name.casefold() not in _ORIGIN_IMAGES:
                close_process_handle(handle)
                raise DataImportError(_OWNERSHIP_ERROR)
            owned = OwnedProcess(int(claimed["pid"]), int(claimed["creation_filetime"]), handle, image)
            self._owned.append(owned)
            self.owned_pids.add(owned.pid)
        finally:
            shutil.rmtree(directory, ignore_errors=True)

    def write_project(
        self,
        stage_dir: Path,
        tables: tuple[PlannedTable, ...],
        prepared: PreparedImport,
        converted: list[list[list[Any]]],
        *,
        pdf: bool,
        conversion_notes: list[list[str]] | None = None,
    ) -> dict[str, Any]:
        """Reset the project, write tables, save ``.opju``, then optionally export PDFs.

        ``op.new`` is not required to return a value. PDF errors stay in the
        result when the project file was saved. A dead session after that save
        is reported with ``session_lost`` so the caller can keep the file and
        start a new process for later tasks.
        """
        if not self.healthy():
            raise OriginSessionLost("Origin session is not running")
        from . import exporter as ex

        sto = ex._origin_support()
        stage_dir.mkdir(parents=True, exist_ok=True)

        def populate() -> tuple[list[Any], list[Any]]:
            return ex._populate_origin_project(self.op, tables, converted, conversion_notes)

        _sheets, graphs = self._call(populate)
        staged = stage_dir / "candidate.opju"

        def save() -> Path:
            return sto._save_origin_project(self.op, staged, self.pending_cleanup)

        saved = self._call(save)
        if not Path(saved).is_file() or Path(saved).stat().st_size <= 0:
            raise sto.OriginExportError("Origin did not create a usable .opju file")
        pdfs: list[dict[str, Any]] = []
        session_lost = False
        if pdf:
            try:
                pdfs = self._export_pdfs(graphs, tables, stage_dir)
            except OriginSessionLost as exc:
                session_lost = True
                self._alive = False
                if not any(not item.get("ok") for item in pdfs):
                    pdfs.append({
                        "ok": False,
                        "table": "",
                        "graph": "",
                        "staged": None,
                        "pages": 0,
                        "error": str(exc),
                    })
        return {"opju": Path(saved), "pdfs": pdfs, "warnings": [], "session_lost": session_lost}

    def _export_pdfs(self, graphs: list[Any], tables: tuple[PlannedTable, ...], stage_dir: Path) -> list[dict[str, Any]]:
        results: list[dict[str, Any]] = []
        graph_index = 0
        for planned in tables:
            if planned.plot.kind == "none":
                continue
            if graph_index >= len(graphs):
                results.append({
                    "ok": False,
                    "table": planned.name,
                    "graph": planned.plot.title or planned.name,
                    "staged": None,
                    "pages": 0,
                    "error": "Origin graph object is missing for this table",
                })
                continue
            graph = graphs[graph_index]
            graph_index += 1
            dest = stage_dir / f"graph_{graph_index:02d}.pdf"
            short_name, long_name = _graph_identity(graph)
            identity = {
                "table": planned.name,
                "graph": long_name or planned.plot.title or planned.name,
                "short_name": short_name,
            }
            try:
                info = self._call(lambda graph=graph, dest=dest, planned=planned: _save_graph_pdf(graph, dest, planned))
                results.append({"ok": True, "staged": dest, **info})
            except OriginSessionLost:
                results.append({
                    "ok": False,
                    "staged": None,
                    "pages": 0,
                    "error": "Origin session was lost while exporting this PDF",
                    **identity,
                })
                raise
            except Exception as exc:
                results.append({
                    "ok": False,
                    "staged": None,
                    "pages": 0,
                    "error": str(exc) or exc.__class__.__name__,
                    **identity,
                })
        return results

    def export_saved_pdfs(self, project: Path, graphs: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Open a saved project and export only the requested graphs.

        The project is opened read-only. This method does not save it. Each
        graph must match one sheet long name and one graph identity. A missing
        or ambiguous match is an error for that graph and does not export a
        different graph.
        """
        if not self.healthy():
            raise OriginSessionLost("Origin session is not running")
        project = Path(project)
        before_bytes = project.read_bytes()
        before = hashlib.sha256(before_bytes).hexdigest()

        def work() -> list[dict[str, Any]]:
            opened = self.op.open(str(project), True, False)
            if opened is False:
                raise DataImportError(f"Origin did not open the saved project: {project}")
            books = list(self.op.pages("w"))
            sheets = iter_project_sheets(books)
            source_name = str((graphs[0] if graphs else {}).get("source_name") or "")
            if not source_name or source_name not in provenance_text(sheets):
                raise DataImportError(
                    f"Saved project provenance does not match {source_name or 'the recorded source'}: {project}"
                )
            pages = list(self.op.pages("g"))
            exported: list[dict[str, Any]] = []
            for item in graphs:
                try:
                    _sheet, graph = project_graph_for(sheets, pages, item)
                    dest = Path(item["dest"])
                    info = _save_graph_pdf(
                        graph,
                        dest,
                        SimpleNamespace(
                            name=str(item.get("table") or ""),
                            plot=SimpleNamespace(title=str(item.get("graph") or ""), kind="line"),
                        ),
                    )
                    exported.append({"ok": True, "staged": dest, **info})
                except OriginSessionLost:
                    raise
                except Exception as exc:
                    exported.append({
                        "ok": False,
                        "staged": None,
                        "pages": 0,
                        "table": str(item.get("table") or ""),
                        "graph": str(item.get("graph") or ""),
                        "short_name": str(item.get("short_name") or ""),
                        "error": str(exc) or exc.__class__.__name__,
                    })
            return exported

        try:
            exported = self._call(work)
        except Exception:
            _guard_saved_project(project, before_bytes, before)
            raise
        _guard_saved_project(project, before_bytes, before)
        return exported

    def _call(self, fn):
        if not self._live or self.timeout_s <= 0:
            return fn()
        self._timed_out = False
        self._deadline = time.monotonic() + self.timeout_s
        try:
            return fn()
        finally:
            fired = self._timed_out
            self._deadline = None
            if fired:
                self._alive = False
                raise OriginSessionLost(
                    "Origin timed out; only the process started for this session was stopped"
                )

    def _watch(self) -> None:
        while not self._stop_watchdog.wait(0.2):
            self._enforce_deadline()

    def _enforce_deadline(self) -> None:
        deadline = self._deadline
        if deadline is None or time.monotonic() < deadline:
            return
        self.terminate_owned()
        self._timed_out = True
        self._alive = False
        self._deadline = None

    def _wait_until_owned_exit(self) -> None:
        deadline = time.monotonic() + 8.0
        while time.monotonic() < deadline:
            if not any(self._identity_ok(item) for item in self._owned):
                return
            time.sleep(0.2)
        self.terminate_owned()


def _sha256_bytes(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _restore_exact(path: Path, payload: bytes) -> bool:
    temporary = path.with_name(path.name + ".restore-tmp")
    try:
        temporary.write_bytes(payload)
        os.replace(temporary, path)
        return _sha256_bytes(path) == hashlib.sha256(payload).hexdigest()
    except OSError:
        return False
    finally:
        if temporary.exists():
            temporary.unlink(missing_ok=True)


def _guard_saved_project(project: Path, before_bytes: bytes, before_hash: str) -> None:
    """Restore ``project`` when a PDF-only open changed its bytes."""
    try:
        after = _sha256_bytes(project)
    except OSError as exc:
        raise DataImportError(f"PDF recovery could not re-read the saved project: {project}") from exc
    if after == before_hash:
        return
    if _restore_exact(project, before_bytes):
        raise DataImportError(
            f"PDF recovery changed the saved project; the original bytes were restored: {project}"
        )
    raise DataImportError(
        f"PDF recovery changed the saved project and the original bytes could not be restored: {project}"
    )
