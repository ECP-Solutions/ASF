# The Python side of asfpy.xll, compiled into the add-in (build.py turns it into boot_py.h).
# asfpy.c runs it once, after Python starts, and then calls only the functions below whose names
# don't start with an underscore. Values cross to C as tagged pairs (TAG_*), so that the C side
# needs no type checks of its own.
import importlib.util
import math
import numbers
import os
import sys
import time
import traceback
import types

TAG_EMPTY, TAG_BOOL, TAG_FLOAT, TAG_INT, TAG_STR, TAG_LIST, TAG_OBJECT, TAG_FAIL, TAG_XLERR, TAG_GRID = range(10)
XLERR_NUM, XLERR_VALUE = 36, 15
MAX_ARGS = 32
MAX_TEXT = 32767

_functions = []          # (qualified name, function, number of arguments, worksheet name or None)
_exports = {}            # 'module.function' -> function, for py.call and PY_CALL
_last_error = ''
_log = None


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


def _fail(exc):
    global _last_error
    _last_error = ''.join(traceback.format_exception_only(type(exc), exc)).strip()
    _note('error: ' + ''.join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
    return (TAG_FAIL, _last_error)


# --- the asfpy module the Python folder's modules import -------------------------------------

def _make_api():
    api = types.ModuleType('asfpy')
    api.__doc__ = 'Decorators of the ASF Python add-in (phase 1a spike).'

    def xl_func(fn=None, *, name=None):
        def mark(f):
            f.__asfpy_xl__ = name or f.__name__
            return f
        return mark(fn) if fn is not None else mark

    def export(fn):
        fn.__asfpy_export__ = True
        return fn

    api.xl_func, api.export = xl_func, export
    return api


def _arguments(fn):
    code = fn.__code__
    return list(code.co_varnames[:code.co_argcount])


def start(python_folder, site_dir, executable, log_path):
    """Pins sys.path, imports the modules of the Python folder and returns the worksheet
    functions to register: a list of (name, argument names, description)."""
    global _log
    _log = _Log(log_path)
    sys.stdout = sys.stderr = _log
    t0 = time.perf_counter()
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
    names = sorted(n for n in os.listdir(python_folder) if n.endswith('.py')) if python_folder and os.path.isdir(python_folder) else []
    for file in names:
        stem = file[:-3]
        t = time.perf_counter()
        try:
            spec = importlib.util.spec_from_file_location(stem, os.path.join(python_folder, file))
            mod = importlib.util.module_from_spec(spec)
            sys.modules[stem] = mod
            spec.loader.exec_module(mod)
        except BaseException as exc:
            sys.modules.pop(stem, None)
            _fail(exc)
            continue
        for attr, fn in sorted(vars(mod).items()):
            if not isinstance(fn, types.FunctionType) or fn.__module__ != stem:
                continue
            sheet = getattr(fn, '__asfpy_xl__', None)
            if sheet or getattr(fn, '__asfpy_export__', False):
                _exports[stem + '.' + attr] = fn
            if sheet:
                if len(_arguments(fn)) > MAX_ARGS:
                    _note('%s.%s: more than %d arguments, not registered' % (stem, attr, MAX_ARGS))
                    continue
                _functions.append((stem + '.' + attr, fn, sheet))
        _note('imported %s in %.0f ms' % (file, (time.perf_counter() - t) * 1000))
    _note('start: %.0f ms' % ((time.perf_counter() - t0) * 1000))
    return [(sheet, _arguments(fn), (fn.__doc__ or '').strip().split('\n')[0]) for _, fn, sheet in _functions]


# --- worksheet functions ---------------------------------------------------------------------

def _cell(v):
    """A Python value as a tagged cell value."""
    if v is None:
        return (TAG_EMPTY, None)
    if isinstance(v, bool):
        return (TAG_BOOL, v)
    if isinstance(v, numbers.Real):
        v = float(v)
        return (TAG_FLOAT, v) if math.isfinite(v) else (TAG_XLERR, XLERR_NUM)
    if isinstance(v, str):
        return (TAG_STR, v[:MAX_TEXT])
    return (TAG_STR, str(v)[:MAX_TEXT])


def _flat(rows):
    # one row or one column of a range is a flat list, as in ASFCells
    if len(rows) == 1:
        return list(rows[0])
    if rows and all(len(r) == 1 for r in rows):
        return [r[0] for r in rows]
    return rows


def xl_call(slot, args):
    """Runs the worksheet function of a slot. A range arrives as a list of rows."""
    try:
        _, fn, _ = _functions[slot]
        result = fn(*[_flat(a) if isinstance(a, list) else a for a in args])
        if isinstance(result, (list, tuple)):
            rows = [list(r) if isinstance(r, (list, tuple)) else [r] for r in result]
            if not rows:
                return (TAG_STR, '')
            width = max(len(r) for r in rows) or 1
            return (TAG_GRID, [[_cell(r[i]) if i < len(r) else (TAG_STR, '') for i in range(width)] for r in rows])
        return _cell(result)
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
        return _last_error

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
    """obj.name, called when it is callable and want_call; name '' is the object itself."""
    try:
        target = obj if name == '' else _member(obj, name)
        if callable(target) and (want_call or args):
            return _com(target(*args), level)
        if args:
            raise TypeError("'%s' is not callable" % (name or type(obj).__name__))
        return _com(target, level)
    except BaseException as exc:
        return _fail(exc)


def com_text(obj):
    try:
        return str(obj)
    except BaseException:
        return '<object>'
