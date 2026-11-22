# The Python side of asfpy.xll, compiled into the add-in (build.py turns it into boot_py.h).
# asfpy.c runs it once, after Python starts, and then calls only the functions below whose names
# don't start with an underscore. Values cross to C as tagged pairs (TAG_*), so that the C side
# needs no type checks of its own. It imports no Windows module: tests/unit/test_boot.py runs it on any OS.
import datetime
import decimal
import importlib
import importlib.util
import inspect
import itertools
import math
import numbers
import os
import re
import sys
import threading
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
ADDIN_VERSION = ''        # the add-in's version, set by the add-in before start()
_start_info = []         # what the last load of the Python folder found, for py.startInfo and the log
_handles = {}            # handle number -> (workbook, object): results that aren't cell values
_owned = {}              # calling cell -> the numbers of the handles its last result made
_handle_numbers = itertools.count(1)
_HANDLE = re.compile(r'^<([A-Za-z_][\w.]*)#(\d+)>$')
_macro_depth = 0         # > 0 while a command (@xl_macro) runs: xl_app() works then only
_async_lock = threading.Lock()
_async_loop = None       # the event loop of async def functions, on a thread of its own
_async_pool = None       # the threads of functions marked asynchronous=True
_async_results = {}      # token -> tagged result, when no native async_return is set (tests)
CO_COROUTINE = 0x0080    # inspect.CO_COROUTINE, without importing inspect

# set by asfpy.c after this file runs; None outside Excel (tests/unit/test_boot.py sets fakes)
_native_caller = None        # () -> '[Book1.xlsx]Sheet1!R2C3', or '' when no cell called
_native_async_return = None  # (token, tagged result) -> bool: xlAsyncReturn
_native_app = None           # () -> the address of Excel's Application (IDispatch, one reference)
_native_release = None       # (address) -> None: IUnknown::Release


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

    def xl_func(fn=None, *, name=None, volatile=False, thread_safe=False, macro=False, asynchronous=False):
        """A worksheet function. volatile: calculated on every recalculation (!); thread_safe:
        may run on Excel's calculation threads ($); macro: a macro sheet equivalent (#);
        asynchronous: runs on a thread of the add-in while Excel calculates other cells (an
        async def function is asynchronous by itself)."""
        def mark(f):
            f.__asfpy_xl__ = name or f.__name__
            f.__asfpy_flags__ = ('!' if volatile else '') + ('#' if macro else '') + ('$' if thread_safe else '')
            f.__asfpy_async__ = bool(asynchronous)
            return f
        return mark(fn) if fn is not None else mark

    def xl_macro(fn=None, *, name=None):
        """A command: VBA runs it with Application.Run "name", and a button or shape can too.
        It runs on Excel's main thread and may use xl_app()."""
        def mark(f):
            f.__asfpy_macro__ = name or f.__name__
            return f
        return mark(fn) if fn is not None else mark

    def export(fn):
        """A function that PY_CALL and ASF scripts may call, by its module and name."""
        fn.__asfpy_export__ = True
        return fn

    api.xl_func, api.xl_macro, api.export, api.xl_app = xl_func, xl_macro, export, xl_app
    return api


def xl_app():
    """Excel's Application object, inside a command (@xl_macro) only: a worksheet function must not
    use Excel's object model. Needs the package pywin32 (or comtypes)."""
    if _native_app is None:
        raise RuntimeError('xl_app() works inside Excel only')
    if _macro_depth <= 0:
        raise RuntimeError('xl_app() works in a command (@xl_macro) only, not in a worksheet function')
    try:
        import pythoncom
        import win32com.client
        wrap = 'pywin32'
    except ImportError:
        try:
            import ctypes
            import comtypes.automation
            import comtypes.client.dynamic
            wrap = 'comtypes'
        except ImportError:
            raise ImportError('xl_app() needs pywin32 or comtypes (ASF tab: Python packages, install pywin32)') from None
    address = _native_app()
    if isinstance(address, str):
        raise RuntimeError(address)
    if wrap == 'pywin32':
        try:
            return win32com.client.Dispatch(pythoncom.ObjectFromAddress(address, pythoncom.IID_IDispatch))
        finally:
            _native_release(address)     # pywin32 took a reference of its own
    pointer = ctypes.cast(ctypes.c_void_p(address), ctypes.POINTER(comtypes.automation.IDispatch))
    return comtypes.client.dynamic.Dispatch(pointer)   # comtypes releases it when it goes


def _is_async_def(fn):
    return bool(getattr(getattr(fn, '__code__', None), 'co_flags', 0) & CO_COROUTINE)


def _arguments(fn):
    code = fn.__code__
    return list(code.co_varnames[:code.co_argcount])


# --- starting, and loading the Python folder ---------------------------------------------------

def start(python_folder, site_dir, executable, log_path, own_python=True):
    """Pins sys.path, then loads the Python folder (load_folder) and returns what it returns.
    own_python False: a Python this file shares, as LibreOffice's (asfpy_lo.py), whose output and
    sys.path stay as they are, but for the Python folder added at the end."""
    global _log, _python_folder
    _log = _Log(log_path)
    if own_python:
        sys.stdout = sys.stderr = _log
        here = os.path.normcase(os.path.abspath(os.getcwd()))
        sys.path[:] = [p for p in sys.path if p and os.path.normcase(os.path.abspath(p)) != here]
    if site_dir and os.path.isdir(site_dir):
        import site
        site.addsitedir(site_dir)
    if python_folder and python_folder not in sys.path:
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


def run_script(path):
    """The command AsfPyRunScript (the studio's Run in Excel): a .py file run as __main__ in
    Excel's Python, as a command, so xl_app() works in it. Its folder is first on sys.path while
    it runs. Returns what it printed, and the traceback if it failed."""
    global _macro_depth
    import io
    import runpy
    if not isinstance(path, str) or not path.lower().endswith('.py') or not os.path.isfile(path):
        return (TAG_STR, 'not a .py file: %s' % (path,))
    out = io.StringIO()
    folder = os.path.dirname(os.path.abspath(path))
    saved = sys.stdout, sys.stderr, sys.argv
    sys.path.insert(0, folder)
    sys.stdout = sys.stderr = out
    sys.argv = [path]
    _macro_depth += 1
    _note('run script %s' % path)
    try:
        runpy.run_path(path, run_name='__main__')
    except SystemExit as exc:
        if exc.code not in (None, 0):
            out.write('\nexit code %s' % (exc.code,))
    except BaseException as exc:
        _fail(exc)
        out.write(_last_traceback)
    finally:
        _macro_depth -= 1
        sys.stdout, sys.stderr, sys.argv = saved
        if folder in sys.path:
            sys.path.remove(folder)
    return (TAG_STR, out.getvalue())


def macro_list():
    """The command AsfPyMacros: one line per command (@xl_macro) of the
    Python folder, as its Excel name, the number of arguments it needs and its description,
    separated by tabs."""
    lines = []
    for _, fn, name, flags in _functions:
        if flags == '%':
            needed = fn.__code__.co_argcount - len(fn.__defaults__ or ())
            doc = (fn.__doc__ or '').strip().split('\n')[0].replace('\t', ' ')
            lines.append('%s\t%d\t%s' % (name, needed, doc))
    return (TAG_STR, '\n'.join(lines))


def _command(name):
    for _, fn, sheet, flags in _functions:
        if flags == '%' and isinstance(name, str) and sheet.upper() == name.upper():
            return sheet, fn
    return None, None


def macro_dialog(name):
    """For the command AsfPyRunMacro: (the dialogs' title, a prompt for each argument) of the
    command (@xl_macro) named name, or the text of why not."""
    sheet, fn = _command(name)
    if fn is None:
        return "ASF Python: no macro '%s' in the Python folder" % (name,)
    args = _arguments(fn)
    required = len(args) - len(fn.__defaults__ or ())
    summary, helps = _docs(fn)[0], _argument_help(fn)
    prompts = []
    for k, a in enumerate(args):
        text = a + (': ' + helps[k] if helps[k] else '')
        if k >= required:
            text += '\n(Optional: leave it blank for its default.)'
        if k == 0 and summary:
            text = summary + '\n\n' + text
        prompts.append(text[:255])                      # Excel's INPUT takes 255 characters
    return 'ASF Python: ' + sheet, prompts


def run_macro(name, values):
    """For the command AsfPyRunMacro: the command named name with the values typed in Excel's
    input boxes, blanks at the end leaving their defaults: (True, its result as text) or (False,
    the error), each as (a line for Excel's alert, the whole text)."""
    global _macro_depth
    sheet, fn = _command(name)
    if fn is None:
        text = "no macro '%s' in the Python folder" % (name,)
        return False, text, text
    values = list(values)
    while values and values[-1] in ('', None):
        values.pop()
    values = [None if v == '' else v for v in values]
    _macro_depth += 1
    try:
        result = fn(*_cell_arguments(fn, values, _Call()))
        if isinstance(result, types.CoroutineType):
            result = _await(result)
        text = _plain_text(result)
        return True, (text or 'Done.')[:250], text
    except BaseException as exc:
        _fail(exc)
        return False, ('%s failed: %s' % (sheet, _last_error))[:250], _last_traceback
    finally:
        _macro_depth -= 1


def _plain_text(value):
    """A result as text: a table as tab-separated rows."""
    kind = _kind(type(value))
    if kind == 'ndarray':
        value = _array_rows(value)
    elif kind:
        value = _frame_rows(value, kind)
    if isinstance(value, (list, tuple)):
        rows = [r if isinstance(r, (list, tuple)) else [r] for r in value]
        return '\n'.join('\t'.join(_plain_text(c) for c in r) for r in rows)
    if value is None:
        return ''
    if isinstance(value, bool):
        return 'TRUE' if value else 'FALSE'
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def last_traceback():
    """The command AsfPyLastError: the traceback of the last Python exception."""
    return (TAG_STR, _last_traceback)


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
            command = getattr(fn, '__asfpy_macro__', None)
            if sheet or command or getattr(fn, '__asfpy_export__', False):
                _exports[stem + '.' + attr] = fn
            if sheet:
                flags = getattr(fn, '__asfpy_flags__', '')
                if getattr(fn, '__asfpy_async__', False) or _is_async_def(fn):
                    flags += '&'    # for asfpy.c: an asynchronous function, not a suffix of Excel's
                limit = MAX_ARGS - 1 if '&' in flags else MAX_ARGS    # an async one also gets Excel's handle
                if len(_arguments(fn)) > limit:
                    _info('%s.%s: more than %d arguments, not registered' % (stem, attr, limit))
                elif '$' in flags and '#' in flags:
                    _info("%s.%s: thread_safe and macro can't go together (Excel refuses $ with #), not registered"
                          % (stem, attr))
                elif '&' in flags and '#' in flags:
                    _info("%s.%s: asynchronous and macro can't go together, not registered" % (stem, attr))
                else:
                    _functions.append((stem + '.' + attr, fn, sheet, flags))
            if command:
                if len(_arguments(fn)) > MAX_ARGS:
                    _info('%s.%s: more than %d arguments, not registered' % (stem, attr, MAX_ARGS))
                else:
                    _functions.append((stem + '.' + attr, fn, command, '%'))   # '%': a command
        _info('imported %s in %.0f ms' % (file, (time.perf_counter() - t) * 1000))
    commands = sum(1 for f in _functions if f[3] == '%')
    _info('%d worksheet functions, %d commands, %d exported functions; load %.0f ms'
          % (len(_functions) - commands, commands, len(_exports), (time.perf_counter() - t0) * 1000))
    _tips.clear()
    for _, fn, sheet, flags in _functions:
        if flags != '%':                        # a command isn't typed in a formula
            _tips[sheet.upper()] = (sheet, fn)
    return [(sheet, _arguments(fn), (fn.__doc__ or '').strip().split('\n')[0], flags, _argument_help(fn))
            for _, fn, sheet, flags in _functions]


# --- help while a formula is typed: the tooltip asfpy.c shows ------------------------------------

_tips = {}               # NAME in capitals: (its name in Excel, the function)
_HINT_TEXT = {int: 'a whole number', float: 'a number', str: 'a text', bool: 'TRUE or FALSE', list: 'a range',
              datetime.date: 'a date', datetime.datetime: 'a date and time', 'ndarray': 'a range of numbers',
              'DataFrame': 'a table, its first row the column names', 'Series': 'a row or column of values'}
_SECTION = re.compile(r'^\s*(Args|Arguments|Parameters|Params|Returns?|Yields?|Raises|Examples?|Notes?|See Also)\s*:?\s*$', re.I)


def _docs(fn):
    """(the summary of fn's docstring, {argument: its description}). The descriptions come from
    an Args: or Parameters: section (Google and NumPy styles) or from :param name: lines."""
    doc = inspect.cleandoc(fn.__doc__ or '')
    lines = doc.split('\n')
    summary = []
    for line in lines:
        if not line.strip() or _SECTION.match(line) or line.lstrip().startswith(':'):
            break
        summary.append(line.strip())
    names = set(_arguments(fn))
    found, current, section = {}, None, False
    for line in lines:
        m = re.match(r'^\s*:param\s+(?:[^:]*\s)?(\w+)\s*:\s*(.*)$', line)
        if m and m.group(1) in names:
            found[m.group(1)], current = m.group(2).strip(), m.group(1)
            continue
        head = _SECTION.match(line)
        if head:
            section, current = head.group(1).lower() in ('args', 'arguments', 'parameters', 'params'), None
            continue
        if not section or re.match(r'^\s*-{3,}\s*$', line):
            continue
        m = re.match(r'^\s*\*{0,2}(\w+)\s*(\([^)]*\))?\s*:\s*(.*)$', line)
        if m and m.group(1) in names:
            current = m.group(1)
            text = m.group(3).strip()
            # NumPy style: "name : type", the description on the lines below
            found[current] = '' if (text and not m.group(2) and line.count(':') == 1 and ' ' not in text.strip()) else text
        elif current and line.strip():
            found[current] = (found[current] + ' ' + line.strip()).strip()
        elif not line.strip():
            current = None
    return ' '.join(summary), found


def _argument_help(fn):
    """A text for each argument, for Insert Function and the tooltip: its description in the
    docstring, else what its type hint asks for, else nothing."""
    _, found = _docs(fn)
    hints = _annotations(fn)
    out = []
    for name in _arguments(fn):
        text = found.get(name) or _HINT_TEXT.get(_hint(hints.get(name)), '')
        out.append(text[:250])
    return out


def _signature(name, fn):
    """(name, [argument as shown], summary, [argument help]); arguments with a default in [ ]."""
    args = _arguments(fn)
    required = len(args) - len(fn.__defaults__ or ())
    shown = [a if k < required else '[%s]' % a for k, a in enumerate(args)]
    return name, shown, _docs(fn)[0], _argument_help(fn)


def _call_at(prefix, sep):
    """(the name the caret ends, or '', the innermost call the caret is in: (its name, the index of
    the argument the caret is in, [the arguments typed before it]) or None), from the text of a
    formula up to the caret."""
    if prefix[:1] not in ('=', '+', '-', '@') or not prefix:
        return '', None
    stack, i, n = [], 1, len(prefix)

    def call():
        while stack and not stack[-1][0]:       # a parenthesis of grouping, not a call
            stack.pop()
        return (stack[-1][0], len(stack[-1][1]), stack[-1][1]) if stack else None

    while i < n:
        c = prefix[i]
        if c in '"\'':                         # a text, or 'Sheet name'!A1
            j = i + 1
            while j < n and not (prefix[j] == c and prefix[j + 1:j + 2] != c):
                j += 2 if prefix[j] == c else 1
            if j >= n:
                return '', call() if c == '"' else None
            i = j + 1
            continue
        if c in '[{':
            close, depth, j = ']' if c == '[' else '}', 1, i + 1
            while j < n and depth:
                depth += (prefix[j] == c) - (prefix[j] == close)
                j += 1
            if depth:
                return '', None
            i = j
            continue
        if c == '(':
            m = re.search(r'([A-Za-z_\\][A-Za-z0-9_.\\]*)\s*$', prefix[1:i])
            stack.append([m.group(1) if m else '', [], i + 1])
        elif c == ')':
            if stack:
                stack.pop()
        elif c == sep and stack:
            stack[-1][1].append(prefix[stack[-1][2]:i].strip())
            stack[-1][2] = i + 1
        i += 1
    m = re.search(r'(?<![A-Za-z0-9_.\\])([A-Za-z_\\][A-Za-z0-9_.\\]*)$', prefix)
    return (m.group(1) if m else ''), call()


def _lookup(name):
    if name.upper().startswith('_XLL.'):
        name = name[5:]
    found = _tips.get(name.upper())
    return _signature(found[0], found[1]) if found else None


def tip(prefix, sep=','):
    """The tooltip for a formula typed up to its caret: (signature before the current argument,
    the current argument, the rest of the signature, the function's summary, the current argument's
    help), or None. While a name is typed: the function of the Python folder it starts, if only one."""
    try:
        word, at = _call_at(prefix, sep)
    except Exception:
        return None
    word = word.upper()
    if len(word) >= 2:
        matches = [k for k in _tips if k.startswith(word)]
        if word in _tips or len(matches) == 1:
            sig = _lookup(word if word in _tips else matches[0])
            return ('%s(%s)' % (sig[0], (sep + ' ').join(sig[1])), '', '', sig[2], '')
    if not at:
        return None
    name, index, before = at
    if name.upper() == 'PY_CALL':
        target = before[0] if before else ''
        fn = _exports.get(target[1:-1]) if index >= 1 and len(target) > 1 and target[0] == target[-1] == '"' else None
        if fn is None:
            sig = ('PY_CALL', ['name', '[arg1]', '...'], 'Calls a function of the Python folder exported with @export.',
                   ['"module.function": a function of the Python folder marked with @export'])
            return _shown(sig, min(index, 1), sep, '')
        return _shown(_signature('PY_CALL', fn), index - 1, sep, target + sep + ' ')
    sig = _lookup(name)
    return _shown(sig, index, sep, '') if sig else None


def _shown(sig, index, sep, lead):
    """The tooltip of a signature with the argument at index current; lead: what PY_CALL's first
    argument gives."""
    name, args, summary, helps = sig
    if not args:
        return '%s(%s)' % (name, lead.rstrip(sep + ' ')), '', '', summary, ''
    if index >= len(args):
        return None
    join = sep + ' '
    head = '%s(%s%s' % (name, lead, join.join(args[:index]) + (join if index else ''))
    tail = (join if index < len(args) - 1 else '') + join.join(args[index + 1:]) + ')'
    text = helps[index] if index < len(helps) else ''
    return head, args[index], tail, summary, ('%s: %s' % (args[index].strip('[]'), text) if text else '')


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


def _numpy_array(v):
    """A range as a numpy array: one row or column 1-D, more 2-D; numbers as float with empty cells
    as nan, else an array of objects."""
    import numpy
    if not isinstance(v, list):
        v = [v]
    nan = float('nan')
    try:
        return numpy.array([[nan if x is None else x for x in r] if isinstance(r, list) else (nan if r is None else r)
                            for r in v], dtype=float)
    except (TypeError, ValueError):
        return numpy.array(v, dtype=object)


def _data_frame(rows):
    """A range as a pandas DataFrame: its first row holds the column names."""
    import pandas
    if not isinstance(rows, list):
        rows = [[rows]]
    header = [_to_str(h) if h is not None else '' for h in rows[0]]
    return pandas.DataFrame(rows[1:], columns=header)


def _series(v):
    import pandas
    return pandas.Series(v if isinstance(v, list) else [v])


_CONVERTERS = {int: _to_int, float: _to_float, str: _to_str, bool: _to_bool,
               datetime.datetime: _to_datetime, datetime.date: _to_date, list: _to_list,
               'ndarray': _numpy_array, 'DataFrame': _data_frame, 'Series': _series}
_HINT_NAMES = {'int': int, 'float': float, 'str': str, 'bool': bool, 'list': list, 'object': object,
               'datetime': datetime.datetime, 'datetime.datetime': datetime.datetime,
               'date': datetime.date, 'datetime.date': datetime.date,
               'ndarray': 'ndarray', 'numpy.ndarray': 'ndarray', 'np.ndarray': 'ndarray',
               'DataFrame': 'DataFrame', 'pandas.DataFrame': 'DataFrame', 'pd.DataFrame': 'DataFrame',
               'Series': 'Series', 'pandas.Series': 'Series', 'pd.Series': 'Series'}


def _kind(t):
    """'ndarray', 'DataFrame' or 'Series' for those classes of numpy and pandas, without importing them."""
    root = (getattr(t, '__module__', None) or '').split('.')[0]
    name = getattr(t, '__name__', '')
    if root == 'numpy' and name == 'ndarray':
        return 'ndarray'
    if root == 'pandas' and name in ('DataFrame', 'Series'):
        return name
    return None


def _hint(h):
    """A type hint as a key of _CONVERTERS, object, or None."""
    if isinstance(h, str):
        h = h.strip()
        return _HINT_NAMES.get(h) or _HINT_NAMES.get(h.split('[')[0].strip())
    for _ in range(3):
        h = typing.get_origin(h) or h   # list[float] is a list
        if not hasattr(h, '__value__'):
            break
        h = h.__value__                 # an alias (type X = ...): numpy's NDArray[float64] is an ndarray
    if h is object or h in _CONVERTERS:
        return h
    return _kind(h)


def _annotations(fn):
    try:
        return dict(fn.__annotations__)
    except Exception:      # a hint that names something undefined
        return {}


def _converters(fn):
    """(the converters of fn's arguments, whether it returns object), from its type hints. An
    argument's converter is None without a hint of _CONVERTERS; its hint is kept too: a str hint
    keeps a handle's text as text, and a DataFrame takes the rows of a range as they are."""
    conv = _hints.get(fn)
    if conv is None:
        annotations = _annotations(fn)
        hints = [_hint(annotations.get(name)) for name in _arguments(fn)]
        conv = ([(_CONVERTERS.get(h), h) for h in hints], _hint(annotations.get('return')) is object)
        _hints[fn] = conv
    return conv


def _flat(rows):
    # one row or one column of a range is a flat list, as in ASFCells
    if len(rows) == 1:
        return list(rows[0])
    if rows and all(len(r) == 1 for r in rows):
        return [r[0] for r in rows]
    return rows


def _cell_arguments(fn, args, ctx):
    """Cell values for fn: a range as a list of rows, one row or column flat; a handle's text as
    its object; then the type hints."""
    conv = _converters(fn)[0]
    out = []
    for i, a in enumerate(args):
        c, hint = conv[i] if i < len(conv) else (None, None)
        if isinstance(a, str) and hint is not str and a.startswith('<'):
            a = ctx.object_of(a)
        elif isinstance(a, list) and hint != 'DataFrame':
            a = _flat(a)
        out.append(c(a) if c else a)
    return out


# --- object handles ---------------------------------------------------------------------------

def _workbook(caller):
    """'[Book1.xlsx]Sheet1!R2C3' -> '[Book1.xlsx]': a handle names objects of its workbook only."""
    end = caller.find(']')
    return caller[:end + 1] if caller.startswith('[') and end > 0 else ''


class _Call:
    """One call of a worksheet function: the cell that called it, read from Excel only when a
    handle needs it, and the handles its result makes."""
    __slots__ = ('_caller', 'new')

    def __init__(self, caller=None):
        self._caller = caller
        self.new = []

    @property
    def caller(self):
        if self._caller is None:
            self._caller = _native_caller() if _native_caller else ''
        return self._caller

    def handle(self, obj):
        number = next(_handle_numbers)
        self.new.append((number, obj))
        return '<%s#%d>' % (type(obj).__name__, number)

    def object_of(self, text):
        m = _HANDLE.match(text)
        if not m:
            return text
        entry = _handles.get(int(m.group(2)))
        if entry is None or entry[0] != _workbook(self.caller):
            raise LookupError('%s names no object of this workbook: calculate the cell that made it again' % text)
        return entry[1]

    def settle(self, ok):
        """The cell's handles from before go; those of this result stay, owned by the cell."""
        if not ok:
            del self.new[:]
        if not (self.new or _owned):
            return
        caller = self.caller
        for number in _owned.pop(caller, ()):
            _handles.pop(number, None)
        if self.new:
            workbook = _workbook(caller)
            for number, obj in self.new:
                _handles[number] = (workbook, obj)
            _owned[caller] = [number for number, _ in self.new]


# --- results ---------------------------------------------------------------------------------

def _serial(v):
    """A date or datetime as Excel's serial number."""
    if isinstance(v, datetime.datetime):
        return (v.replace(tzinfo=None) - EPOCH) / datetime.timedelta(days=1)
    return float((datetime.datetime(v.year, v.month, v.day) - EPOCH).days)


def _plain(v):
    """numpy's scalars as Python's: numpy.float64(2) as 2.0, a datetime64 as a datetime."""
    if (type(v).__module__ or '').split('.')[0] == 'numpy' and getattr(v, 'shape', None) == () and hasattr(v, 'item'):
        if getattr(getattr(v, 'dtype', None), 'kind', '') == 'M':
            v = v.astype('datetime64[us]')
        return v.item()
    if isinstance(v, decimal.Decimal):
        return float(v)
    return v


def _cell(v, call=None):
    """A Python value as a tagged cell value; another object as a handle, kept for the cell."""
    v = _plain(v)
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
    if call is None or isinstance(v, numbers.Number):
        return (TAG_STR, str(v)[:MAX_TEXT])
    return (TAG_STR, call.handle(v))


def _array_rows(a):
    if a.ndim == 0:
        return a.item()
    if a.ndim > 2:
        raise ValueError("a %d-dimensional array can't go to cells: 2 dimensions at most" % a.ndim)
    return a.tolist()


def _frame_rows(f, kind):
    """A DataFrame as its column names and rows, a Series as one column; missing values (NaN, NaT,
    None) empty. The index is a column too when it has a name or isn't made of whole numbers."""
    import pandas
    index = f.index
    keep = any(n is not None for n in index.names) or not pandas.api.types.is_integer_dtype(index.dtype)
    values = f.astype(object).where(f.notna(), None)
    if kind == 'Series':
        rows = [[x] for x in values.tolist()]
        header = None
    else:
        rows = values.values.tolist()
        header = [c if isinstance(c, str) else ' '.join(map(str, c)) if isinstance(c, tuple) else str(c)
                  for c in f.columns]
    if keep:
        rows = [[i if not isinstance(i, tuple) else ' '.join(map(str, i))] + r for i, r in zip(index.tolist(), rows)]
        if header is not None:
            header = [' '.join(str(n) for n in index.names if n is not None)] + header
    return ([header] if header is not None else []) + rows


def _result(result, call=None, as_object=False):
    if isinstance(result, types.CoroutineType):
        result = _await(result)
    if as_object and result is not None:
        return (TAG_STR, call.handle(result)) if call is not None else _cell(str(result))
    kind = _kind(type(result))
    if kind == 'ndarray':
        result = _array_rows(result)
    elif kind:
        result = _frame_rows(result, kind)
    if isinstance(result, (list, tuple)):
        rows = [list(r) if isinstance(r, (list, tuple)) else [r] for r in result]
        if not rows:
            return (TAG_STR, '')
        width = max(len(r) for r in rows) or 1
        return (TAG_GRID, [[_cell(r[i], call) if i < len(r) else (TAG_STR, '') for i in range(width)] for r in rows])
    return _cell(result, call)


# --- worksheet functions, commands and PY_CALL ----------------------------------------------------

def _run(fn, args, call):
    """fn with the cell values of args: its tagged result."""
    ok = False
    try:
        r = _result(fn(*_cell_arguments(fn, args, call)), call, _converters(fn)[1])
        ok = True
    except BaseException as exc:
        r = _fail(exc)
    call.settle(ok)
    return r


def xl_call(slot, args):
    """Runs the worksheet function or command of a slot. A range arrives as a list of rows."""
    global _macro_depth
    call = _Call()
    try:
        _, fn, _, flags = _functions[slot]
    except BaseException as exc:
        return _fail(exc)
    if flags != '%':
        return _run(fn, args, call)
    _macro_depth += 1
    try:
        return _run(fn, args, call)
    finally:
        _macro_depth -= 1


def xl_call_export(name, args):
    """PY_CALL: an exported function, looked up by its name. The text of a cell never reaches
    importlib, so PY_CALL("os.system", ...) finds nothing: #NAME?, as for any unknown name."""
    fn = _exports.get(name) if isinstance(name, str) else None
    if fn is None:
        _fail(NameError("no exported function '%s' in the Python folder" % (name,)))
        return (TAG_XLERR, XLERR_NAME)
    return _run(fn, args, _Call())


# --- asynchronous functions --------------------------------------------------------------------

def _loop():
    """The event loop of async def functions, started on a thread of its own on first use."""
    global _async_loop
    with _async_lock:
        if _async_loop is None:
            import asyncio
            loop = asyncio.new_event_loop()
            threading.Thread(target=loop.run_forever, name='asfpy-async', daemon=True).start()
            _async_loop = loop
    return _async_loop


def _pool():
    global _async_pool
    with _async_lock:
        if _async_pool is None:
            import concurrent.futures
            _async_pool = concurrent.futures.ThreadPoolExecutor(thread_name_prefix='asfpy-async')
    return _async_pool


def _await(coroutine):
    """A coroutine's result, waited for: an async def function called by PY_CALL or py.call."""
    import asyncio
    return asyncio.run_coroutine_threadsafe(coroutine, _loop()).result()


def _async_return(token, tagged):
    if _native_async_return is not None:
        _native_async_return(token, tagged)
    else:
        _async_results[token] = tagged


def xl_call_async(slot, args, token):
    """Starts the asynchronous function of a slot and returns at once; its result goes to Excel
    with xlAsyncReturn (token: Excel's handle, kept by asfpy.c) when it is done. An async def
    function runs on the add-in's event loop, any other on a thread of its pool. The calling cell is
    read now, while Excel's call lasts."""
    call = _Call()
    try:
        fn = _functions[slot][1]
        cell_args = _cell_arguments(fn, args, call)
        call.caller                       # handles need it later, on another thread
        if _is_async_def(fn):
            import asyncio
            future = asyncio.run_coroutine_threadsafe(fn(*cell_args), _loop())
        else:
            future = _pool().submit(fn, *cell_args)
    except BaseException as exc:
        call.settle(False)
        _async_return(token, _fail(exc))
        return

    def done(f):
        ok = False
        try:
            r = _result(f.result(), call, _converters(fn)[1])
            ok = True
        except BaseException as exc:
            r = _fail(exc)
        call.settle(ok)
        _async_return(token, r)
    future.add_done_callback(done)


# --- the objects VBA sees: roots and proxies -------------------------------------------------

class _Root:
    """py for level 1: calls of exported functions only."""
    def call(self, name, *args):
        fn = _exports.get(name)
        if fn is None:
            raise NameError("no exported function '%s' in the Python folder" % name)
        r = fn(*args)
        return _await(r) if isinstance(r, types.CoroutineType) else r

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
        return 'ASF Python %s\nPython %s, %d-bit: %s\n%s' % (ADDIN_VERSION or '?', sys.version.split()[0],
                                                             64 if sys.maxsize > 2 ** 32 else 32, sys.prefix,
                                                             '\n'.join(_start_info))

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
    v = _plain(v)
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
