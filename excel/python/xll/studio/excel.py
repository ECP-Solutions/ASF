"""Excel, reached from the studio's process through COM (pywin32): the add-in's commands run in
Excel's own Python. One thread owns COM; calls wait for Excel when it is busy."""
import queue
import threading
import time

_calls = queue.Queue()
_thread = None
BUSY = (-2147418111, -2146777998, -2147417846)   # RPC_E_CALL_REJECTED, VBA_E_IGNORE, RPC_E_SERVERCALL_RETRYLATER


class NotReachable(Exception):
    pass


def _worker():
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
    except ImportError:
        win32com = None
    app = None
    while True:
        fn, done = _calls.get()
        if win32com is None:
            done(None, NotReachable('Excel is reachable on Windows only, with pywin32'))
            continue
        for attempt in range(40):
            try:
                if app is None:
                    app = win32com.client.GetActiveObject('Excel.Application')
                done(fn(app), None)
                break
            except Exception as exc:           # pywintypes.com_error and the like
                code = getattr(exc, 'hresult', None) or (exc.args[0] if exc.args and isinstance(exc.args[0], int) else None)
                if code in BUSY and attempt < 39:
                    time.sleep(0.25)          # Excel is editing a cell or showing a dialog
                    continue
                app = None
                done(None, NotReachable('Excel isn\'t reachable: %s' % (exc,)))
                break


def call(fn, timeout=600):
    """fn(app) on the COM thread: its result, or NotReachable."""
    global _thread
    if _thread is None:
        _thread = threading.Thread(target=_worker, name='excel-com', daemon=True)
        _thread.start()
    box = {}
    ready = threading.Event()

    def done(result, error):
        box['result'], box['error'] = result, error
        ready.set()

    _calls.put((fn, done))
    if not ready.wait(timeout):
        raise NotReachable('Excel didn\'t answer')
    if box['error']:
        raise box['error']
    return box['result']


def run(name, *args):
    """Application.Run of a command of the add-in or of the Python folder."""
    return call(lambda app: app.Run(name, *args))


def reload():
    """The Python folder imported again (AsfPyReload), then the workbooks calculated in full."""
    def go(app):
        app.Run('AsfPyReload')
        app.CalculateFull()
        return True
    return call(go)


def connected():
    try:
        return call(lambda app: str(app.Version), timeout=5)
    except NotReachable:
        return None
