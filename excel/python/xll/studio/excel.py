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


class ExcelError(ValueError):
    """Excel answered, with an error: a macro it doesn't know, or one that failed in VBA."""


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
                if code == -2147352567:        # DISP_E_EXCEPTION: Excel's own answer
                    info = exc.args[2] if len(exc.args) > 2 and isinstance(exc.args[2], tuple) else ()
                    done(None, ExcelError((info[2] if len(info) > 2 and info[2] else str(exc)).strip()))
                    break
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


# Excel's error values as Application.Run returns them through COM
ERRORS = {-2146826288: '#NULL!', -2146826281: '#DIV/0!', -2146826273: '#VALUE!', -2146826265: '#REF!',
          -2146826259: '#NAME?', -2146826252: '#NUM!', -2146826246: '#N/A', -2146826245: '#GETTING_DATA'}


def run(name, *args):
    """Application.Run of a command of the add-in or of the Python folder."""
    return call(lambda app: app.Run(name, *args))


def front(app):
    """Excel's window in front, for a macro that shows a dialog (xl_app().InputBox, a message box)
    or works on the selection; Windows lets a background process do so through SwitchToThisWindow
    where SetForegroundWindow only flashes the taskbar button."""
    try:
        import ctypes
        user32 = ctypes.windll.user32
        hwnd = int(app.Hwnd)
        if user32.IsIconic(hwnd):
            user32.ShowWindow(hwnd, 9)                   # SW_RESTORE
        if not user32.SetForegroundWindow(hwnd) or user32.GetForegroundWindow() != hwnd:
            user32.SwitchToThisWindow(hwnd, True)
    except Exception:
        pass


def macro(name, args=(), to_front=False):
    """A macro of the Python folder run as VBA runs it (Application.Run), with its arguments:
    {'result': text, 'error': bool}. A failed macro gives Application.Run #VALUE! only, so its
    traceback is asked of the add-in (AsfPyLastError)."""
    def go(app):
        try:
            from pythoncom import Empty
        except ImportError:                                # a stand-in for Excel, in the tests
            Empty = None
        values = list(args)
        while values and values[-1] is None:
            values.pop()                                  # trailing blanks: the macro's defaults
        values = [Empty if v is None else v for v in values]
        if to_front:
            front(app)
        out = app.Run(name, *values)
        if isinstance(out, int) and out in ERRORS:
            detail = ''
            if out == -2146826273:
                try:
                    detail = app.Run('AsfPyLastError') or ''
                except Exception:
                    pass
            return {'result': ERRORS[out] + ('\n' + detail if detail else ''), 'error': True}
        return {'result': text(out), 'error': False}
    return call(go)


def text(value):
    """A result of Application.Run as text: a range of values as tab-separated rows."""
    if value is None:
        return ''
    if isinstance(value, tuple):
        rows = [r if isinstance(r, tuple) else (r,) for r in value]
        return '\n'.join('\t'.join(text(c) for c in r) for r in rows)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


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
