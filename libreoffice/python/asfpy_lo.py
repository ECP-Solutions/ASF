# Python for ASF in LibreOffice (phase 5 of excel/python/DESIGN.md): what asfpy.xll gives Excel,
# through LibreOffice's own Python. libreoffice/tools/build.py puts this file into ASF.oxt with the
# text of excel/python/xll/boot.py in BOOT_SOURCE, so both hosts share the same rules: PY_CALL
# reaches functions exported with @export only, py.call the same, and py.module only at level 2.
#
# LOpython.bas (the ASF library) runs the functions below through the script provider, from the
# extension (location=user:uno_packages) and never from a document. Python objects go to Basic as
# Proxy, an XInvocation: Basic calls their members by name, as VBA does through IDispatch.
import os
import sys
import tempfile
import types
import weakref

import unohelper
from com.sun.star.lang import XServiceInfo
from com.sun.star.script import XInvocation
from com.sun.star.uno import RuntimeException

BOOT_SOURCE = None   # build.py: the text of boot.py
BOOT_MODULE = '_asfpy_boot_lo'
# A Python object back from Basic as an argument comes as this text and its Proxy's number:
# pyuno doesn't take its own XInvocation object back (it asks it for XInvocation2)
REFERENCE = 'asfpy:'
_proxies = weakref.WeakValueDictionary()


def _boot():
    """boot.py as a module, made once per LibreOffice: it keeps the Python folder's functions and
    the last error between calls, even when the script provider loads this file again."""
    mod = sys.modules.get(BOOT_MODULE)
    if mod is None:
        mod = types.ModuleType(BOOT_MODULE)
        exec(compile(BOOT_SOURCE, 'boot.py (ASF.oxt)', 'exec'), mod.__dict__)
        mod._lo_folder = None
        sys.modules[BOOT_MODULE] = mod
    return mod


def default_folder():
    """%APPDATA%\\ASF\\python on Windows, as for Excel; ~/.config/ASF/python elsewhere."""
    base = os.environ.get('APPDATA')
    if base:
        return os.path.join(base, 'ASF', 'python')
    return os.path.join(os.path.expanduser('~'), '.config', 'ASF', 'python')


def _started():
    b = _boot()
    if b._lo_folder is None:
        _load(b, default_folder())
    return b


def _load(b, folder):
    b.start(folder, '', '', os.path.join(tempfile.gettempdir(), 'asfpy-lo.log'), own_python=False)
    b._lo_folder = folder


# --- values between Basic and Python ------------------------------------------------------------

def _from_uno(v):
    """A Basic value for Python: arrays (tuples) as lists, a Proxy as its Python object."""
    if isinstance(v, tuple):
        return [_from_uno(x) for x in v]
    if isinstance(v, Proxy):
        return v._obj
    if isinstance(v, str) and v.startswith(REFERENCE):
        proxy = _proxies.get(int(v[len(REFERENCE):]))
        if proxy is None:
            raise RuntimeException('a Python object passed back is gone', None)
        return proxy._obj
    return v


def _to_uno(tagged, level):
    """A tagged value of boot.py for Basic; TAG_FAIL raises, with Python's message."""
    b = _boot()
    tag, v = tagged
    if tag == b.TAG_FAIL:
        raise RuntimeException(v, None)
    if tag == b.TAG_LIST:
        return tuple(_to_uno(x, level) for x in v)
    if tag == b.TAG_OBJECT:
        return Proxy(v, level)
    if tag == b.TAG_EMPTY:
        return None
    return v


class Proxy(unohelper.Base, XInvocation, XServiceInfo):
    """A Python object for Basic: obj.name(args) calls it, obj.name reads it. Names are found as
    in Excel: exactly first, then without case when one only matches; those that start with _
    are refused."""

    def __init__(self, obj, level):
        self._obj = obj
        self._level = level
        _proxies[id(self)] = self

    def _member(self, name, call, args):
        return _to_uno(_boot().com_invoke(self._obj, name, self._level, call, [_from_uno(a) for a in args]), self._level)

    # XInvocation
    def getIntrospection(self):
        return None

    def invoke(self, name, params, out_index, out_params):
        return self._member(name, True, params), (), ()

    def getValue(self, name):
        if name == '$GetTypeName':
            # Basic's TypeOf asks an XInvocation object its type this way (as it asks an OLE
            # object), and takes an object that doesn't answer as any class: ASF's TypeOf x Is
            # ASF_Map would be True
            return 'Python.' + type(self._obj).__name__
        if name == '$Reference':
            # what LOhelper passes back to Python for this object
            return REFERENCE + str(id(self))
        return self._member(name, False, ())

    def setValue(self, name, value):
        raise RuntimeException("a member of a Python object can't be set from Basic: %s" % name, None)

    def hasMethod(self, name):
        try:
            return callable(_boot()._member(self._obj, name))
        except Exception:
            return False

    def hasProperty(self, name):
        try:
            return not callable(_boot()._member(self._obj, name))
        except Exception:
            return False

    # XServiceInfo: what typeof shows in ASF
    def getImplementationName(self):
        return type(self._obj).__name__

    def supportsService(self, name):
        return False

    def getSupportedServiceNames(self):
        return ()


# --- the functions LOpython.bas runs ---------------------------------------------------------------

def root(level):
    """py: level 1 py.call of exported functions, level 2 py.module too."""
    level = 2 if int(level) >= 2 else 1
    return Proxy(_started().root(level), level)


def py_call(name, args):
    """PY_CALL(name, args...) of a cell: (0, value), (1, rows) or (2, error code)."""
    b = _started()
    t = b.xl_call_export(name, [_from_uno(a) for a in args])
    if t[0] == b.TAG_GRID:
        return 1, tuple(tuple(_cell(c) for c in row) for row in t[1])
    if t[0] in (b.TAG_XLERR, b.TAG_FAIL):
        return 2, t[1] if t[0] == b.TAG_XLERR else b.XLERR_VALUE
    return 0, _cell(t)


def _cell(t):
    b = _boot()
    tag, v = t
    if tag == b.TAG_EMPTY:
        return ''
    if tag in (b.TAG_XLERR, b.TAG_FAIL):
        return '#VALUE!' if tag == b.TAG_FAIL else {29: '#NAME?', 36: '#NUM!'}.get(v, '#VALUE!')
    return v


def set_folder(folder):
    """The Python folder: its modules are imported (again), and PY_CALL and py find their
    functions there. "" is the default folder."""
    _load(_boot(), folder or default_folder())
    return _boot()._lo_folder


def reload():
    """The Python folder's modules imported again, as AsfPyReload does in Excel."""
    b = _started()
    b.reload()
    return '\n'.join(b._start_info)


def start_info():
    return _started().root(1).startInfo


def last_error():
    return _started()._last_error


g_exportedScripts = (root, py_call, set_folder, reload, start_info, last_error)
