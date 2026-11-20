# The Python side of asfpy.xll, compiled into the add-in (build.py turns it into boot_py.h).
# asfpy.c runs it once, after Python starts, and then calls only the functions below whose names
# don't start with an underscore. Values cross to C as tagged pairs (TAG_*), so that the C side
# needs no type checks of its own. It imports no Windows module: tests/test_boot.py runs it on any OS.
import datetime
import importlib
import importlib.util
import math
import numbers
import os
import sys
import time
import traceback
import types
import typing

TAG_EMPTY, TAG_BOOL, TAG_FLOAT, TAG_INT, TAG_STR, TAG_LIST, TAG_OBJECT, TAG_FAIL, TAG_XLERR, TAG_GRID, TAG_STOP = range(11)
XLERR_NAME, XLERR_NUM, XLERR_VALUE = 29, 36, 15
MAX_ARGS = 32
MAX_TEXT = 32767
EPOCH = datetime.datetime(1899, 12, 30)   # Excel's day 0 in the 1900 date system

_functions = []          # (qualified name, function, worksheet name, type suffixes), by slot
_exports = {}            # 'module.function' -> function, for py.call and PY_CALL
_hints = {}              # function -> the converters of its arguments, from its type hints
_folder_modules = []     # the names of the modules imported from the Python folder
_python_folder = ''
_last_error = ''
_last_traceback = ''
_log = None
_start_info = []         # what the last load of the Python folder found, for py.startInfo and the log


class _Log:
    def __init__(self, path):
        self.path = path

    def write(self, text):
        try:
            with open(self.path, 'a', encoding='utf-8') as f:
                f.write(text)
        except OSError:
            pass

    def flush(self):
        pass


def _note(text):
    if _log:
        _log.write(time.strftime('%H:%M:%S ') + text + '\n')


def _info(text):
    _start_info.append(text)
    _note(text)


def _fail(exc):
    global _last_error, _last_traceback
    _last_error = ''.join(traceback.format_exception_only(type(exc), exc)).strip()
    _last_traceback = ''.join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    _note('error: ' + _last_traceback)
    return (TAG_FAIL, _last_error)


# --- the asfpy module the Python folder's modules import -------------------------------------

def _make_api():
    api = types.ModuleType('asfpy')
    api.__doc__ = 'Decorators of the ASF Python add-in.'

    def xl_func(fn=None, *, name=None, volatile=False, thread_safe=False, macro=False):
        """A worksheet function. volatile: calculated on every recalculation (!); thread_safe:
        may run on Excel's calculation threads ($); macro: a macro sheet equivalent (#)."""
        def mark(f):
            f.__asfpy_xl__ = name or f.__name__
            f.__asfpy_flags__ = ('!' if volatile else '') + ('#' if macro else '') + ('$' if thread_safe else '')
            return f
        return mark(fn) if fn is not None else mark

    def export(fn):
        """A function that PY_CALL and ASF scripts may call, by its module and name."""
        fn.__asfpy_export__ = True
        return fn

    api.xl_func, api.export = xl_func, export
    return api


def _arguments(fn):
    code = fn.__code__
    return list(code.co_varnames[:code.co_argcount])


# --- starting, and loading the Python folder ---------------------------------------------------

def start(python_folder, site_dir, executable, log_path):
    """Pins sys.path, then loads the Python folder (load_folder) and returns what it returns."""
    global _log, _python_folder
    _log = _Log(log_path)
    sys.stdout = sys.stderr = _log
    here = os.path.normcase(os.path.abspath(os.getcwd()))
    sys.path[:] = [p for p in sys.path if p and os.path.normcase(os.path.abspath(p)) != here]
    if site_dir and os.path.isdir(site_dir):
        import site
        site.addsitedir(site_dir)
    if python_folder:
        sys.path.append(python_folder)
    if executable and os.path.isfile(executable):
        sys.executable = sys._base_executable = executable
    sys.modules['asfpy'] = _make_api()
    _note('python %s, sys.path=%r' % (sys.version.split()[0], sys.path))
    _python_folder = python_folder
    return load_folder()


def reload():
    """The Python folder again: its modules are imported anew, while Python and the packages
    they imported stay. Returns the worksheet functions to register, as start() does."""
    _note('reload of the Python folder')
    return load_folder()


def _same_file(a, b):
    return bool(a) and os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def _import_source(stem, path):
    """A module of the Python folder, compiled from its source: a changed file is read again on
    reload whatever its time stamp, and no __pycache__ appears in the folder."""
    with open(path, 'rb') as f:
        source = f.read()
    mod = types.ModuleType(stem)
    mod.__file__ = path
    mod.__spec__ = importlib.util.spec_from_file_location(stem, path)
    sys.modules[stem] = mod
    try:
        exec(compile(source, path, 'exec'), mod.__dict__)
    except BaseException:
        sys.modules.pop(stem, None)
        raise
    return mod


def load_folder():
    """Imports the modules of the Python folder and returns the worksheet functions to register:
    a list of (name, argument names, description, type suffixes)."""
    t0 = time.perf_counter()
    del _start_info[:]
    del _functions[:]
    _exports.clear()
    _hints.clear()
    for stem in _folder_modules:
        sys.modules.pop(stem, None)
    del _folder_modules[:]
    importlib.invalidate_caches()
    folder = _python_folder
    names = []
    if not folder or not os.path.isdir(folder):
        _info('the Python folder %s does not exist' % folder)
    else:
        names = sorted(n for n in os.listdir(folder) if n.lower().endswith('.py'))
        others = sorted(n for n in os.listdir(folder) if not n.lower().endswith('.py') and n != '__pycache__')
        _info('Python folder %s: %s' % (folder, ', '.join(names) or 'no .py files'))
        if others:
            _info('not modules (only .py files are): %s' % ', '.join(others[:10]))
    for file in names:
        stem = file[:-3]
        path = os.path.join(folder, file)
        t = time.perf_counter()
        mod = sys.modules.get(stem)
        if not (mod and _same_file(getattr(mod, '__file__', None), path)):   # else imported by a module before it
            try:
                mod = _import_source(stem, path)
            except BaseException as exc:
                _fail(exc)
                _info('%s did not import: %s: %s' % (file, type(exc).__name__, exc))
                continue
        _folder_modules.append(stem)
        for attr, fn in sorted(vars(mod).items()):
            if not isinstance(fn, types.FunctionType) or fn.__module__ != stem:
                continue
            sheet = getattr(fn, '__asfpy_xl__', None)
            if sheet or getattr(fn, '__asfpy_export__', False):
                _exports[stem + '.' + attr] = fn
            if sheet:
                flags = getattr(fn, '__asfpy_flags__', '')
                if len(_arguments(fn)) > MAX_ARGS:
                    _info('%s.%s: more than %d arguments, not registered' % (stem, attr, MAX_ARGS))
                    continue
                if '$' in flags and '#' in flags:
                    _info("%s.%s: thread_safe and macro can't go together (Excel refuses $ with #), not registered"
                          % (stem, attr))
                    continue
                _functions.append((stem + '.' + attr, fn, sheet, flags))
        _info('imported %s in %.0f ms' % (file, (time.perf_counter() - t) * 1000))
    _info('%d worksheet functions, %d exported functions; load %.0f ms'
          % (len(_functions), len(_exports), (time.perf_counter() - t0) * 1000))
    return [(sheet, _arguments(fn), (fn.__doc__ or '').strip().split('\n')[0], flags)
            for _, fn, sheet, flags in _functions]


# --- arguments: type hints ------------------------------------------------------------------

def _to_int(v):
    if isinstance(v, float):
        if not v.is_integer():
            raise ValueError('expected a whole number, got %r' % v)
        return int(v)
    return v


def _to_float(v):
    return float(v) if isinstance(v, numbers.Real) and not isinstance(v, bool) else v


def _to_str(v):
    if v is None or isinstance(v, str):
        return v
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _to_bool(v):
    return v if v is None else bool(v)


def _to_datetime(v):
    if isinstance(v, numbers.Real) and not isinstance(v, bool):
        return EPOCH + datetime.timedelta(days=float(v))
    return v


def _to_date(v):
    v = _to_datetime(v)
    return v.date() if isinstance(v, datetime.datetime) else v


def _to_list(v):
    if v is None:
        return []
    return v if isinstance(v, list) else [v]


_CONVERTERS = {int: _to_int, float: _to_float, str: _to_str, bool: _to_bool,
               datetime.datetime: _to_datetime, datetime.date: _to_date, list: _to_list}
_HINT_NAMES = {'int': int, 'float': float, 'str': str, 'bool': bool, 'list': list,
               'datetime': datetime.datetime, 'datetime.datetime': datetime.datetime,
               'date': datetime.date, 'datetime.date': datetime.date}


def _converters(fn):
    """The converters of fn's arguments, from its type hints: int, float, str, bool, list, date
    and datetime; None where an argument has none of them."""
    conv = _hints.get(fn)
    if conv is None:
        try:
            annotations = dict(fn.__annotations__)
        except Exception:      # a hint that names something undefined
            annotations = {}
        conv = []
        for name in _arguments(fn):
            hint = annotations.get(name)
            if isinstance(hint, str):
                hint = _HINT_NAMES.get(hint.strip())
            hint = typing.get_origin(hint) or hint     # list[float] is a list
            conv.append(_CONVERTERS.get(hint) if isinstance(hint, type) else None)
        _hints[fn] = conv
    return conv


def _flat(rows):
    # one row or one column of a range is a flat list, as in ASFCells
    if len(rows) == 1:
        return list(rows[0])
    if rows and all(len(r) == 1 for r in rows):
        return [r[0] for r in rows]
    return rows


def _cell_arguments(fn, args):
    """Cell values for fn: a range as a list of rows, one row or column flat, then the type hints."""
    conv = _converters(fn)
    out = []
    for i, a in enumerate(args):
        if isinstance(a, list):
            a = _flat(a)
        c = conv[i] if i < len(conv) else None
        out.append(c(a) if c else a)
    return out


# --- results ---------------------------------------------------------------------------------

def _serial(v):
    """A date or datetime as Excel's serial number."""
    if isinstance(v, datetime.datetime):
        return (v.replace(tzinfo=None) - EPOCH) / datetime.timedelta(days=1)
    return float((datetime.datetime(v.year, v.month, v.day) - EPOCH).days)


def _cell(v):
    """A Python value as a tagged cell value."""
    if v is None:
        return (TAG_EMPTY, None)
    if isinstance(v, bool):
        return (TAG_BOOL, v)
    if isinstance(v, datetime.date):
        return (TAG_FLOAT, _serial(v))
    if isinstance(v, numbers.Real):
        v = float(v)
        return (TAG_FLOAT, v) if math.isfinite(v) else (TAG_XLERR, XLERR_NUM)
    if isinstance(v, str):
        return (TAG_STR, v[:MAX_TEXT])
    return (TAG_STR, str(v)[:MAX_TEXT])


def _result(result):
    if isinstance(result, (list, tuple)):
        rows = [list(r) if isinstance(r, (list, tuple)) else [r] for r in result]
        if not rows:
            return (TAG_STR, '')
        width = max(len(r) for r in rows) or 1
        return (TAG_GRID, [[_cell(r[i]) if i < len(r) else (TAG_STR, '') for i in range(width)] for r in rows])
    return _cell(result)


# --- worksheet functions and PY_CALL ---------------------------------------------------------

def xl_call(slot, args):
    """Runs the worksheet function of a slot. A range arrives as a list of rows."""
    try:
        fn = _functions[slot][1]
        return _result(fn(*_cell_arguments(fn, args)))
    except BaseException as exc:
        return _fail(exc)


def xl_call_export(name, args):
    """PY_CALL: an exported function, looked up by its name. The text of a cell never reaches
    importlib, so PY_CALL("os.system", ...) finds nothing: #NAME?, as for any unknown name."""
    fn = _exports.get(name) if isinstance(name, str) else None
    if fn is None:
        _fail(NameError("no exported function '%s' in the Python folder" % (name,)))
        return (TAG_XLERR, XLERR_NAME)
    try:
        return _result(fn(*_cell_arguments(fn, args)))
    except BaseException as exc:
        return _fail(exc)


# --- the objects VBA sees: roots and proxies -------------------------------------------------

class _Root:
    """py for level 1: calls of exported functions only."""
    def call(self, name, *args):
        fn = _exports.get(name)
        if fn is None:
            raise NameError("no exported function '%s' in the Python folder" % name)
        return fn(*args)

    @property
    def lastError(self):
        """The last Python exception, one line: from a worksheet function, PY_CALL or py."""
        return _last_error

    @property
    def lastTraceback(self):
        """The traceback of the last Python exception."""
        return _last_traceback

    def clearError(self):
        """Forgets the last exception: an error after it is a new one (Why this error?)."""
        global _last_error, _last_traceback
        _last_error = _last_traceback = ''

    @property
    def startInfo(self):
        return 'Python %s\n%s' % (sys.version.split()[0], '\n'.join(_start_info))

    def __repr__(self):
        return '<py>'


class _Root2(_Root):
    """py for level 2: modules too."""
    def module(self, name):
        if not isinstance(name, str) or name.startswith('_') or '._' in name:
            raise ImportError("module '%s' can't be used" % name)
        return importlib.import_module(name)


def root(level):
    return _Root2() if level >= 2 else _Root()


def _member(obj, name):
    if not name or name.startswith('_'):
        raise AttributeError("'%s' can't be used" % name)
    try:
        return getattr(obj, name)
    except AttributeError:
        pass
    # VBA and ASF may change the case of a name: a match without case, when there is one only
    low = name.lower()
    hits = [n for n in dir(obj) if n.lower() == low and not n.startswith('_')]
    if len(hits) == 1:
        return getattr(obj, hits[0])
    raise AttributeError("'%s' has no member '%s'" % (type(obj).__name__, name))


def _com(v, level):
    """A Python value as a tagged VARIANT."""
    if v is None:
        return (TAG_EMPTY, None)
    if isinstance(v, bool):
        return (TAG_BOOL, v)
    if isinstance(v, datetime.date):
        return (TAG_FLOAT, _serial(v))
    if isinstance(v, numbers.Integral) and abs(int(v)) < 2 ** 31:
        return (TAG_INT, int(v))
    if isinstance(v, numbers.Real):
        return (TAG_FLOAT, float(v))
    if isinstance(v, str):
        return (TAG_STR, v)
    if isinstance(v, (list, tuple)):
        return (TAG_LIST, [_com(x, level) for x in v])
    if isinstance(v, types.ModuleType) and level < 2:
        raise TypeError('modules are not handed out at this level')
    if (type(v).__module__ or '').split('.')[0] in ('win32com', 'pythoncom', 'comtypes', 'pywintypes'):
        raise TypeError('COM objects are not handed out')
    return (TAG_OBJECT, v)


def com_invoke(obj, name, level, want_call, args):
    """obj.name, called when it is callable and want_call; name '' is the object itself, whose
    value is its text (VBA's x = obj and Debug.Print obj)."""
    try:
        target = obj if name == '' else _member(obj, name)
        if callable(target) and (want_call or args):
            return _com(target(*args), level)
        if args:
            raise TypeError("'%s' is not callable" % (name or type(obj).__name__))
        if name == '':
            return (TAG_STR, com_text(target))
        return _com(target, level)
    except BaseException as exc:
        return _fail(exc)


def com_iter(obj):
    """iter(obj), for VBA's For Each."""
    try:
        return (TAG_OBJECT, iter(obj))
    except BaseException as exc:
        return _fail(exc)


def com_next(it, level):
    """The next item of a For Each, or TAG_STOP at the end."""
    try:
        return _com(next(it), level)
    except StopIteration:
        return (TAG_STOP, None)
    except BaseException as exc:
        return _fail(exc)


def com_text(obj):
    try:
        return str(obj)
    except BaseException:
        return '<object>'
