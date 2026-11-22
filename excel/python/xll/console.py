"""The Python console of the ASF tab (command AsfPyConsole): the Python of the add-in, interactive,
in a console window of its own, apart from Excel.

Run by asfpy.xll as  python.exe -I -i %TEMP%\\asfpy-console.py PYTHON_FOLDER : this file runs, then
Python's prompt. pip("...") manages the packages; the Python folder is on the import path, with a
stand-in asfpy module, so its modules import and their functions can be tried here.
"""
import os
import sys
import types

try:
    import ctypes
    ctypes.windll.kernel32.SetConsoleTitleW("ASF Python console")
except Exception:
    pass


def pip(command='list'):
    """Runs pip for this Python: pip("install numpy"), pip("list"), pip("show pandas")."""
    import shlex
    import subprocess
    args = shlex.split(command.replace("\\", "\\\\"))   # quotes group, backslashes stay
    if args and args[0].lower() == 'pip':
        args = args[1:]
    code = subprocess.call([sys.executable, '-I', '-m', 'pip', '--disable-pip-version-check'] + args)
    if code:
        print('pip exited with %d' % code)


def _stand_in_asfpy():
    """asfpy as the Python folder's modules import it; in Excel the add-in gives the real one."""
    api = types.ModuleType('asfpy')
    api.__doc__ = 'A stand-in for the console: the decorators mark nothing, xl_app() is only in Excel.'

    def _decorator(fn=None, **options):
        return fn if fn is not None else (lambda f: f)

    def xl_app():
        raise RuntimeError('xl_app() works inside Excel only, in a command (@xl_macro)')

    api.xl_func = api.xl_macro = _decorator
    api.export = lambda fn: fn
    api.xl_app = xl_app
    return api


folder = sys.argv[1].strip('"') if len(sys.argv) > 1 else ''
sys.argv = sys.argv[:1]
sys.modules.setdefault('asfpy', _stand_in_asfpy())
if folder and os.path.isdir(folder):
    sys.path.append(folder)
del os, types

print('ASF Python console: Python %s, %d-bit, %s' % (sys.version.split()[0], 64 if sys.maxsize > 2**32 else 32, sys.prefix))
print('Python folder (importable here): %s' % (folder or 'none'))
print()
print('  pip("install numpy")    pip("list")    pip("show pandas")    pip("uninstall -y openpyxl")')
print('  import numpy            to check a package, and the modules of the Python folder too')
print()
print('This Python runs apart from Excel: Excel uses new packages after a restart, and its own')
print('Python keeps the packages it loaded. exit() or Ctrl+Z, Enter closes the window.')
