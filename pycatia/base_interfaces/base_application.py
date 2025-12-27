from win32com.client import Dispatch
import pythoncom
import psutil

from pycatia.in_interfaces.application import Application


def _is_catia_running() -> bool:
    """
    Check whether a CATIA process (CNEXT.exe) is currently running.

    Returns:
        True if a CNEXT.exe process is found in the Windows process list,
        False otherwise.
    """
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and proc.info['name'].lower() == 'cnext.exe':
            return True
    return False


def catia_application(co_initialise: bool = False, require_catia_running: bool = False) -> Application:

    """
    Initialize a connection to CATIA and return a pycatia Application object.

    NOTE:
        By default, if CATIA is not already running, invoking
        pythoncom.Dispatch() (which this function uses) will create
        a new CATIA (CNEXT.exe) process. In many enterprise environments,
        this process may start in the background without displaying a UI.

         If this happens, even launching CATIA manually (e.g. from Windows
         Explorer) may attach pycatia to the original background process.
         In such cases, the user may be unable to interact with CATIA
         until the hidden process is terminated via Task Manager.

          To avoid spawning a hidden CATIA instance, set ``require_catia_running=True``
          to ensure that an existing CATIA session is used.

    Args:
        co_initialise:
            If True, initialise the COM library for the current thread
            using pythoncom.CoInitialize().
        require_catia_running:
            If True, ensure that an existing CATIA (CNEXT.exe) process
            is already running before attempting to connect.

    Returns:
        An initialized pycatia ``Application`` object.

    Raises:
        ProcessLookupError:
            If ``require_catia_running`` is True and no CATIA process is running.
    """
    if require_catia_running and not _is_catia_running():
        raise ProcessLookupError("CATIA is not running")

    if co_initialise:
        return Application(Dispatch('CATIA.Application', pythoncom.CoInitialize()))
    return Application(Dispatch('CATIA.Application'))
