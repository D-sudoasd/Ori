"""One Origin process owned by this program, reused across batch tasks.

The session refuses an Origin process it did not start. A timeout or a forced
stop terminates only those recorded process ids. It never issues a global
``taskkill`` of Origin.
"""

from __future__ import annotations

import ctypes
import re
import sys
import threading
import time
from pathlib import Path
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


def terminate_pid(pid: int) -> bool:
    """Stop one process id. This does not look up Origin by image name."""
    if sys.platform != "win32" or isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0:
        return False
    kernel32 = ctypes.windll.kernel32
    kernel32.OpenProcess.restype = ctypes.c_void_p
    kernel32.OpenProcess.argtypes = [ctypes.c_uint, ctypes.c_int, ctypes.c_uint]
    kernel32.TerminateProcess.argtypes = [ctypes.c_void_p, ctypes.c_uint]
    kernel32.TerminateProcess.restype = ctypes.c_int
    kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
    handle = kernel32.OpenProcess(0x0001, 0, int(pid))
    if not handle:
        return False
    try:
        return bool(kernel32.TerminateProcess(handle, 1))
    finally:
        kernel32.CloseHandle(handle)


def validate_pdf_file(path: Path) -> dict[str, Any]:
    """Require a PDF header, trailer, cross-reference, and at least one page object."""
    data = Path(path).read_bytes()
    if len(data) < 32 or not data.startswith(b"%PDF-"):
        raise DataImportError(f"PDF header missing: {path}")
    if b"%%EOF" not in data[-4096:] and b"%%EOF" not in data:
        raise DataImportError(f"PDF trailer missing: {path}")
    if b"startxref" not in data and b"/XRef" not in data:
        raise DataImportError(f"PDF cross-reference missing: {path}")
    pages = _PAGE_RE.findall(data)
    if not pages:
        raise DataImportError(f"PDF has no page object: {path}")
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
        self.pending_cleanup: list[Path] = []
        self._before: set[int] = set()
        self._live = False
        self._alive = False
        self._closed = False
        self._deadline: float | None = None
        self._timed_out = False
        self._stop_watchdog = threading.Event()
        self._watchdog: threading.Thread | None = None

    def healthy(self) -> bool:
        if self._closed or not self.started or not self._alive or self.op is None:
            return False
        if not self._live:
            return True
        if not self.owned_pids:
            return False
        return not self.owned_pids.isdisjoint(origin_pids())

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
        if self._live and self.timeout_s > 0:
            self._watchdog = threading.Thread(
                target=self._watch,
                name="origin-session-watchdog",
                daemon=True,
            )
            self._watchdog.start()
        launched = False
        try:
            self._call(lambda: op.set_show(False))
            launched = True
            if self._live:
                self.owned_pids = self._await_owned_pids(sto)
                if not self.owned_pids:
                    raise sto.OriginExportError(
                        "Origin did not start a process owned by this session"
                    )
            else:
                self.owned_pids = set()
            self.started = True
            self._alive = True
            self.starts += 1
            OriginSession.launch_count += 1
        except BaseException:
            if launched or self.owned_pids:
                self.started = True
                self._alive = True
                self.close()
            else:
                self._stop_watchdog.set()
                self.op = None
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
        if not self.started:
            self._closed = True
            self.op = None
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
        self.started = False
        self._alive = False
        self._closed = True
        self.op = None

    def terminate_owned(self) -> list[int]:
        """Stop Origin processes that appeared after this session's baseline."""
        current = origin_pids()
        if self.owned_pids:
            victims = sorted(pid for pid in self.owned_pids if pid in current and pid not in self._before)
        else:
            victims = sorted(current - set(self._before))
        for pid in victims:
            terminate_pid(pid)
        return victims

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
            try:
                info = self._call(lambda graph=graph, dest=dest, planned=planned: _save_graph_pdf(graph, dest, planned))
                results.append({"ok": True, "staged": dest, **info})
            except OriginSessionLost:
                results.append({
                    "ok": False,
                    "table": planned.name,
                    "graph": planned.plot.title or planned.name,
                    "staged": None,
                    "pages": 0,
                    "error": "Origin session was lost while exporting this PDF",
                })
                raise
            except Exception as exc:
                results.append({
                    "ok": False,
                    "table": planned.name,
                    "graph": planned.plot.title or planned.name,
                    "staged": None,
                    "pages": 0,
                    "error": str(exc) or exc.__class__.__name__,
                })
        return results

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
            deadline = self._deadline
            if deadline is not None and time.monotonic() >= deadline:
                self.terminate_owned()
                self._timed_out = True
                self._alive = False
                self._deadline = None

    def _await_owned_pids(self, sto) -> set[int]:
        limit = self.timeout_s if self.timeout_s > 0 else 60.0
        deadline = time.monotonic() + min(limit, 60.0)
        while True:
            current = origin_pids()
            owned = current - set(self._before)
            if owned:
                return owned
            if current & set(self._before):
                raise sto.OriginExportError(
                    "Origin attached to a process that was already running; the existing session was left untouched"
                )
            if time.monotonic() >= deadline:
                return set()
            time.sleep(0.2)

    def _wait_until_owned_exit(self) -> None:
        deadline = time.monotonic() + 8.0
        while time.monotonic() < deadline:
            if self.owned_pids.isdisjoint(origin_pids()):
                return
            time.sleep(0.2)
        self.terminate_owned()
