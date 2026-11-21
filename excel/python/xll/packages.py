"""The console of ASF Python > Python packages: pip for the Python of the add-in.

Run by asfpy.xll as  python.exe -I %TEMP%\\asfpy-packages.py  in a console window of its own, so
pip's downloads never block Excel.
"""
import shlex
import subprocess
import sys

try:
    import ctypes
    ctypes.windll.kernel32.SetConsoleTitleW("ASF Python packages")
except Exception:
    pass


def pip(args):
    return subprocess.call([sys.executable, "-I", "-m", "pip", "--disable-pip-version-check"] + args)


print("Python %s, %d-bit: %s" % (sys.version.split()[0], 64 if sys.maxsize > 2**32 else 32, sys.prefix))
print()
pip(["list"])
print()
print("Type a pip command, such as:  install numpy  |  install --upgrade pandas  |  uninstall -y openpyxl  |  list")
print("An empty line closes this window. Excel uses new packages after a restart; close Excel")
print("before upgrading or removing a package that Excel has loaded.")
while True:
    try:
        line = input("pip> ").strip()
    except (EOFError, KeyboardInterrupt):
        break
    if not line or line.lower() in ("exit", "quit"):
        break
    try:
        args = shlex.split(line.replace("\\", "\\\\"))   # quotes group, backslashes stay
    except ValueError as e:
        print(e)
        continue
    if args[0].lower() == "pip":
        args = args[1:]
    if args:
        print("pip exited with %d" % pip(args))
