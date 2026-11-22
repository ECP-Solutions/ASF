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
    or works on the selection: True if it is. Windows lets only the program in front (here the
    studio's window, Edge's) bring another forward; a background process such as this one only
    makes the taskbar button flash. So, in turn: this thread shares the input of the window in
    front (AttachThreadInput), then sends an input event of its own (the program that sent the
    last input may bring a window forward), then switches as Alt+Tab does. If Windows still says
    no, Excel's taskbar button flashes."""
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.WinDLL('user32')
        kernel32 = ctypes.WinDLL('kernel32')
    except (ImportError, OSError, AttributeError):
        return False
    HWND, DWORD, BOOL = wintypes.HWND, wintypes.DWORD, wintypes.BOOL
    user32.GetForegroundWindow.restype = HWND
    user32.GetWindowThreadProcessId.argtypes = [HWND, ctypes.c_void_p]
    user32.GetWindowThreadProcessId.restype = DWORD
    user32.AttachThreadInput.argtypes = [DWORD, DWORD, BOOL]
    for f in ('IsIconic', 'SetForegroundWindow', 'BringWindowToTop', 'SetActiveWindow'):
        getattr(user32, f).argtypes = [HWND]
    user32.ShowWindow.argtypes = [HWND, ctypes.c_int]
    user32.SwitchToThisWindow.argtypes = [HWND, BOOL]
    try:
        hwnd = int(app.Hwnd)
    except Exception:
        return False

    def in_front():
        return (user32.GetForegroundWindow() or 0) == hwnd

    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, 9)                          # SW_RESTORE
    if in_front():
        return True
    me = kernel32.GetCurrentThreadId()
    fg = user32.GetForegroundWindow()
    them = user32.GetWindowThreadProcessId(fg, None) if fg else 0
    attached = bool(them and them != me and user32.AttachThreadInput(me, them, True))
    try:
        user32.BringWindowToTop(hwnd)
        user32.SetForegroundWindow(hwnd)
        user32.SetActiveWindow(hwnd)
    finally:
        if attached:
            user32.AttachThreadInput(me, them, False)
    if in_front():
        return True

    class MOUSEINPUT(ctypes.Structure):
        _fields_ = [('dx', ctypes.c_long), ('dy', ctypes.c_long), ('mouseData', DWORD), ('dwFlags', DWORD),
                    ('time', DWORD), ('dwExtraInfo', ctypes.c_size_t)]

    class INPUT(ctypes.Structure):
        _fields_ = [('type', DWORD), ('mi', MOUSEINPUT)]       # 40 bytes in 64-bit, 28 in 32-bit, as Windows' INPUT
    move = INPUT(0, MOUSEINPUT(0, 0, 0, 0x0001, 0, 0))       # INPUT_MOUSE, MOUSEEVENTF_MOVE by 0: nothing moves
    user32.SendInput(1, ctypes.byref(move), ctypes.sizeof(INPUT))
    user32.SetForegroundWindow(hwnd)
    if in_front():
        return True
    user32.SwitchToThisWindow(hwnd, True)
    if in_front():
        return True

    class FLASHWINFO(ctypes.Structure):
        _fields_ = [('cbSize', ctypes.c_uint), ('hwnd', HWND), ('dwFlags', DWORD), ('uCount', ctypes.c_uint),
                    ('dwTimeout', DWORD)]
    flash = FLASHWINFO(ctypes.sizeof(FLASHWINFO), hwnd, 0x3 | 0xC, 0, 0)   # caption and button, until in front
    user32.FlashWindowEx(ctypes.byref(flash))
    return False


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
        in_front = front(app) if to_front else None
        out = app.Run(name, *values)
        if isinstance(out, int) and out in ERRORS:
            detail = ''
            if out == -2146826273:
                try:
                    detail = app.Run('AsfPyLastError') or ''
                except Exception:
                    pass
            return {'result': ERRORS[out] + ('\n' + detail if detail else ''), 'error': True, 'front': in_front}
        return {'result': text(out), 'error': False, 'front': in_front}
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
