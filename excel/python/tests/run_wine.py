"""Runs asfpy64.xll or asfpy32.xll in tests/fakehost.c (a stand-in for Excel) under Wine, on Linux.

    python3 excel/python/tests/run_wine.py PYTHON [PYTHON...]

PYTHON is a folder with a Windows Python (python3XX.dll, Lib), such as the tools folder of the
"python" or "pythonx86" package on nuget.org, or a venv made from one. Each is tried in turn, with
the add-in of its bitness. Needs wine (wine32 too for 32-bit) and MinGW-w64; run
excel/python/xll/build.py first. The Python folder of the
run is excel/python/spike. WINEPREFIX is used when set, else a prefix in the system's temp folder.
"""
import glob
import struct
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
PYDIR = os.path.dirname(HERE)
BUILD = os.path.join(PYDIR, "xll", "build")
KEY = r"HKCU\Software\VB and VBA Program Settings\ASF\AddIn"


def win(path):
    return "Z:" + os.path.abspath(path).replace("/", "\\")


def bits(python):
    """32 or 64: the machine of the python3XX.dll of an install, or of the home of a venv."""
    home = python
    cfg = os.path.join(python, "pyvenv.cfg")
    if os.path.isfile(cfg):
        for line in open(cfg, encoding="utf-8"):
            key, _, value = line.partition("=")
            if key.strip().lower() == "home":
                home = value.strip()
                if home[:3].lower() == "z:\\":
                    home = home[2:].replace("\\", "/")
    for dll in glob.glob(os.path.join(home, "python3*.dll")):
        if os.path.basename(dll)[7:-4].isdigit():   # python3NN.dll, not python3.dll nor python3t.dll
            data = open(dll, "rb").read(4096)
            pe = struct.unpack_from("<I", data, 0x3C)[0]
            return 64 if struct.unpack_from("<H", data, pe + 4)[0] == 0x8664 else 32
    sys.exit("no python3XX.dll in %s" % home)


def logs(env):
    return glob.glob(os.path.join(env["WINEPREFIX"], "drive_c", "users", "*", "**", "asfpy.log"), recursive=True)


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    env = dict(os.environ, WINEDEBUG="-all")
    env.setdefault("WINEPREFIX", os.path.join(tempfile.gettempdir(), "asfpy-wine"))
    for b, gcc in ((64, "x86_64-w64-mingw32-gcc"), (32, "i686-w64-mingw32-gcc")):
        subprocess.run([gcc, "-O2", "-municode", os.path.join(HERE, "fakehost.c"), "-o",
                        os.path.join(BUILD, "fakehost%d.exe" % b), "-loleaut32", "-lole32", "-luuid",
                        "-Wl,--kill-at"], check=True)   # MdCallBack12 undecorated, as Excel exports it
    failed = 0
    for python in sys.argv[1:]:
        b = bits(python)
        print("=== %s (%d-bit)" % (python, b), flush=True)
        for name, value in (("PythonVenv", win(python)), ("PythonFolder", win(os.path.join(PYDIR, "spike")))):
            subprocess.run(["wine", "reg", "add", KEY, "/v", name, "/d", value, "/f"], env=env, check=True,
                           capture_output=True)
        for log in logs(env):
            os.remove(log)
        numpy = glob.glob(os.path.join(python, "Lib", "site-packages", "numpy"))
        r = subprocess.run(["wine", os.path.join(BUILD, "fakehost%d.exe" % b), win(os.path.join(BUILD, "asfpy%d.xll" % b))]
                           + (["numpy"] if numpy else []), env=env, capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout, end="")
        if r.returncode:
            failed += 1
            print(r.stderr[-2000:], end="")
            for log in logs(env):
                print("--- %s" % log)
                print(open(log, encoding="utf-8", errors="replace").read()[-4000:])
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
