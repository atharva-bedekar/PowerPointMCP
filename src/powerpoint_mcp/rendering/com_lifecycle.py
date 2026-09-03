"""PowerPoint COM lifecycle management, process isolation, and defensive file operations."""

from contextlib import contextmanager
import gc
import logging
import os
from pathlib import Path
import signal
import sys
import threading
import time
from typing import Any, Callable, Optional, Set, Tuple, TypeVar

logger = logging.getLogger("powerpoint_mcp.com_lifecycle")

T = TypeVar("T")

# Thread-safe registry of PowerPoint.exe PIDs spawned and owned by this MCP server instance
_ACTIVE_MCP_PIDS: Set[int] = set()
_MCP_PID_LOCK = threading.Lock()


def get_powerpoint_pids() -> Set[int]:
    """Retrieve all running POWERPNT.EXE process IDs on the system using fast native Windows APIs."""
    if sys.platform != "win32":
        return set()

    import ctypes
    from ctypes import wintypes

    TH32CS_SNAPPROCESS = 0x00000002

    class PROCESSENTRY32(ctypes.Structure):
        _fields_ = [
            ("dwSize", wintypes.DWORD),
            ("cntUsage", wintypes.DWORD),
            ("th32ProcessID", wintypes.DWORD),
            ("th32DefaultHeapID", ctypes.c_void_p),
            ("th32ModuleID", wintypes.DWORD),
            ("cntThreads", wintypes.DWORD),
            ("th32ParentProcessID", wintypes.DWORD),
            ("pcPriClassBase", ctypes.c_long),
            ("dwFlags", wintypes.DWORD),
            ("szExeFile", ctypes.c_char * 260),
        ]

    pids: Set[int] = set()
    kernel32 = ctypes.windll.kernel32
    h_snapshot = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    if h_snapshot == -1 or h_snapshot == wintypes.HANDLE(-1).value:
        return pids

    try:
        entry = PROCESSENTRY32()
        entry.dwSize = ctypes.sizeof(PROCESSENTRY32)
        if kernel32.Process32First(h_snapshot, ctypes.byref(entry)):
            while True:
                exe_name = entry.szExeFile.decode("latin1", errors="ignore").lower()
                if exe_name == "powerpnt.exe":
                    pids.add(entry.th32ProcessID)
                if not kernel32.Process32Next(h_snapshot, ctypes.byref(entry)):
                    break
    except Exception as exc:
        logger.debug(f"Error enumerating PowerPoint processes via Toolhelp32: {exc}")
    finally:
        kernel32.CloseHandle(h_snapshot)

    return pids


def register_mcp_pid(pid: int) -> None:
    """Register a PowerPoint PID as spawned and owned by this MCP server."""
    with _MCP_PID_LOCK:
        _ACTIVE_MCP_PIDS.add(pid)
    logger.debug(f"Registered MCP-owned PowerPoint PID: {pid}")


def unregister_mcp_pid(pid: int) -> None:
    """Unregister a PowerPoint PID from MCP server ownership."""
    with _MCP_PID_LOCK:
        _ACTIVE_MCP_PIDS.discard(pid)
    logger.debug(f"Unregistered MCP-owned PowerPoint PID: {pid}")


def get_active_mcp_pids() -> Set[int]:
    """Return a copy of the set of currently active MCP-owned PowerPoint PIDs."""
    with _MCP_PID_LOCK:
        return set(_ACTIVE_MCP_PIDS)


def terminate_mcp_pid(pid: int) -> bool:
    """Terminate a specific PowerPoint process ONLY if it is registered as owned by MCP."""
    if pid not in get_active_mcp_pids():
        logger.warning(f"Refusing to terminate unmanaged PowerPoint PID: {pid}")
        return False

    try:
        os.kill(pid, signal.SIGTERM)
        logger.info(f"Terminated MCP-owned PowerPoint process (PID: {pid})")
        return True
    except Exception as exc:
        logger.debug(f"Could not terminate PID {pid}: {exc}")
        return False


def ensure_mcp_pid_closed(pid: int, timeout: float = 0.5) -> bool:
    """Wait briefly for an MCP-spawned PowerPoint PID to exit; force terminate if it hangs."""
    t_end = time.time() + timeout
    while time.time() < t_end:
        if pid not in get_powerpoint_pids():
            unregister_mcp_pid(pid)
            return True
        time.sleep(0.05)

    # Process is still running and belongs to MCP: terminate it safely
    if pid in get_active_mcp_pids() and pid in get_powerpoint_pids():
        logger.warning(f"MCP PowerPoint process (PID {pid}) did not exit within {timeout}s; terminating.")
        terminate_mcp_pid(pid)

    unregister_mcp_pid(pid)
    return True


def cleanup_mcp_com_processes(timeout: float = 0.5) -> int:
    """Terminate only lingering PowerPoint COM processes owned by this MCP server instance.

    Does NOT touch any unrelated PowerPoint processes opened by the user.
    """
    if sys.platform != "win32":
        return 0

    active_pids = get_active_mcp_pids()
    if not active_pids:
        return 0

    current_pids = get_powerpoint_pids()
    cleaned = 0
    for pid in active_pids:
        if pid in current_pids:
            if terminate_mcp_pid(pid):
                cleaned += 1
        unregister_mcp_pid(pid)

    gc.collect()
    try:
        import pythoncom
        pythoncom.CoUninitialize()
    except Exception:
        pass

    if cleaned > 0:
        logger.info(f"Cleaned up {cleaned} lingering MCP PowerPoint process(es).")
    return cleaned


@contextmanager
def com_powerpoint_session():
    """Context manager for safe, isolated PowerPoint COM automation.

    Lifecycle:
    1. pythoncom.CoInitialize()
    2. Snapshot running PowerPoint PIDs before DispatchEx
    3. Launch PowerPoint.Application and record spawned PID
    4. Set Visible = 0
    5. Yield (ppt_app, spawned_pid)
    6. In finally:
       - Run gc.collect() to release child COM wrappers (slide, slides, presentation) while app is alive
       - ppt_app.Quit() (only if MCP spawned this instance)
       - Clear ppt_app = None
       - gc.collect()
       - pythoncom.CoUninitialize()
       - Verify spawned process terminated; clean up if lingering.
    """
    if sys.platform != "win32":
        raise RuntimeError("PowerPoint COM automation is only supported on Windows.")

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    before_pids = get_powerpoint_pids()
    ppt_app = None
    spawned_pid: Optional[int] = None
    is_owned = False

    try:
        ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
        after_pids = get_powerpoint_pids()
        spawned = after_pids - before_pids
        if spawned:
            spawned_pid = next(iter(spawned))
            register_mcp_pid(spawned_pid)
            is_owned = True

        try:
            ppt_app.Visible = 0
        except Exception:
            pass

        yield ppt_app, spawned_pid
    finally:
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass

        # Step 1: Force garbage collection to deallocate child COM proxies while server is alive
        gc.collect()

        # Step 2: Quit PowerPoint application if owned by MCP
        if ppt_app is not None and is_owned:
            try:
                ppt_app.Quit()
            except Exception as exc:
                logger.debug(f"Error during ppt_app.Quit(): {exc}")
        ppt_app = None

        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass

        # Step 3: Final garbage collection while apartment is active
        gc.collect()

        # Step 4: Ensure process is completely terminated (allow graceful exit before hard kill)
        if spawned_pid is not None and is_owned:
            ensure_mcp_pid_closed(spawned_pid, timeout=3.0)


def is_file_locked_error(exc: BaseException) -> bool:
    """Check if an exception is a file locking / sharing violation (e.g. WinError 32)."""
    if isinstance(exc, PermissionError):
        return True
    if isinstance(exc, OSError):
        winerror = getattr(exc, "winerror", None)
        if winerror == 32:  # ERROR_SHARING_VIOLATION
            return True
        if "being used by another process" in str(exc).lower():
            return True
        if sys.platform == "win32" and getattr(exc, "errno", None) == 13:  # EACCES
            return True
    return False


def defensive_file_operation(
    op: Callable[[], T],
    target_path: Any,
    action_name: str = "save",
    max_retries: int = 4,
    initial_delay: float = 0.1,
) -> T:
    """Execute a file operation with defensive retry behavior on Windows COM file locks (WinError 32).

    If a file lock is encountered:
    1. Cleans up any lingering MCP-owned PowerPoint COM processes.
    2. Waits with exponential backoff and retries the operation.
    3. If still locked after retries, raises an actionable PermissionError detailing the cause.
    """
    resolved_path = Path(target_path).resolve()
    last_error: Optional[BaseException] = None
    delay = initial_delay

    for attempt in range(1, max_retries + 1):
        try:
            return op()
        except Exception as exc:
            if not is_file_locked_error(exc):
                raise

            last_error = exc
            logger.warning(
                f"File '{resolved_path.name}' is locked during {action_name} (attempt {attempt}/{max_retries}). "
                f"Cleaning up MCP COM processes and retrying in {delay:.2f}s..."
            )

            # Cleanup MCP-owned COM processes that might be holding a file handle
            cleanup_mcp_com_processes()

            time.sleep(delay)
            delay = min(delay * 2, 1.0)

    # All retries failed - construct clear actionable error
    all_pids = get_powerpoint_pids()
    if all_pids:
        msg = (
            f"Cannot {action_name} presentation at '{resolved_path}': The file is locked by another process ([WinError 32]). "
            f"PowerPoint process(es) (PID(s): {sorted(all_pids)}) may currently have this presentation open. "
            f"Please save and close the file in PowerPoint, then retry."
        )
    else:
        msg = (
            f"Cannot {action_name} presentation at '{resolved_path}': The file is currently locked by another process ([WinError 32]). "
            f"Please ensure no other application is accessing the file and retry."
        )

    logger.error(msg)
    raise PermissionError(msg) from last_error
