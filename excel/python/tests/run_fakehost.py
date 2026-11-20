"""Runs asfpy64.xll and asfpy32.xll in tests/fakehost.c, a stand-in for Excel: natively on Windows,
under Wine on Linux.

    python excel/python/tests/run_fakehost.py PYTHON [PYTHON...]

PYTHON is a folder with a Windows Python (python3XX.dll, Lib): an install, a venv made from one,
or the tools folder of the "python" or "pythonx86" package on nuget.org. Each is tried with the
add-in of its bitness. Run excel/python/xll/build.py first: the fake host is built with the same
toolchain. The Python folder of the run is a temporary folder with spike/spike.py and
tests/zz_slots.py.

On Windows the add-in's settings (PythonVenv and PythonFolder under HKCU) are set for the run and
put back afterwards, and the add-in's log goes to a temporary folder, not to your %TEMP%. On Linux
it needs wine (wine32 too for 32-bit) and MinGW-w64; WINEPREFIX is used when set, else a prefix
in the system's temp folder.
"""
import glob
import os
import shutil
import struct
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
PYDIR = os.path.dirname(HERE)
BUILD = os.path.join(PYDIR, "xll", "build")
sys.path.insert(0, os.path.join(PYDIR, "xll"))
import build  # noqa: E402  (its toolchains)

KEY = r"Software\VB and VBA Program Settings\ASF\AddIn"
NAMES = ("PythonVenv", "PythonFolder")
WINDOWS = os.name == "nt"


def native(path):
    """A path as the add-in sees it: as is on Windows, on Wine's Z: drive on Linux."""
    path = os.path.abspath(path)
    return path if WINDOWS else "Z:" + path.replace("/", "\\")


def bits(python):
    """32 or 64: the machine of the python3XX.dll of an install, or of the home of a venv."""
    home = python
    cfg = os.path.join(python, "pyvenv.cfg")
    if os.path.isfile(cfg):
        for line in open(cfg, encoding="utf-8"):
            key, _, value = line.partition("=")
            if key.strip().lower() == "home":
                home = value.strip()
                if not WINDOWS and home[:3].lower() == "z:\\":
                    home = home[2:].replace("\\", "/")
    for dll in glob.glob(os.path.join(home, "python3*.dll")):
        if os.path.basename(dll)[7:-4].isdigit():   # python3NN.dll, not python3.dll nor python3t.dll
            data = open(dll, "rb").read(4096)
            pe = struct.unpack_from("<I", data, 0x3C)[0]
            return 64 if struct.unpack_from("<H", data, pe + 4)[0] == 0x8664 else 32
    sys.exit("no python3XX.dll in %s" % home)


def build_fakehost(tool, b):
    exe = os.path.join(BUILD, "fakehost%d.exe" % b)
    src = os.path.join(HERE, "fakehost.c")
    if tool == "mingw":
        gcc = "x86_64-w64-mingw32-gcc" if b == 64 else "i686-w64-mingw32-gcc"
        subprocess.run([gcc, "-O2", "-municode", src, "-o", exe, "-loleaut32", "-lole32", "-luuid", "-lpsapi",
                        "-Wl,--kill-at"], check=True)   # MdCallBack12 undecorated, as Excel exports it
    else:
        build.msvc_run(b, ["cl"] + build.MSVC_FLAGS + [src, "/Fo" + os.path.join(BUILD, "fakehost%d.obj" % b),
                                                      "/Fe" + exe, "/link", "oleaut32.lib", "ole32.lib", "uuid.lib",
                                                      "psapi.lib"], cwd=BUILD)
    return exe


def read_settings():
    import winreg
    old = {}
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, KEY) as k:
            for name in NAMES:
                try:
                    old[name] = winreg.QueryValueEx(k, name)
                except FileNotFoundError:
                    pass
    except FileNotFoundError:
        pass
    return old


def write_settings(values, old=None):
    """Sets the values; with old, the names it lacks are deleted, as they were before the run."""
    import winreg
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, KEY) as k:
        for name in NAMES:
            if name in values:
                value, kind = values[name] if isinstance(values[name], tuple) else (values[name], winreg.REG_SZ)
                winreg.SetValueEx(k, name, 0, kind, value)
            elif old is not None:
                try:
                    winreg.DeleteValue(k, name)
                except FileNotFoundError:
                    pass


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")   # the checks print Ω and 中文
    try:
        tool = build.toolchain(open(os.path.join(BUILD, "toolchain.txt")).read().strip())
    except FileNotFoundError:
        sys.exit("run excel/python/xll/build.py first")
    pythons = sys.argv[1:]
    exes = {b: build_fakehost(tool, b) for b in sorted({bits(p) for p in pythons})}
    tmp = tempfile.mkdtemp(prefix="asfpy-test-")
    folder = os.path.join(tmp, "python")
    os.makedirs(folder)
    shutil.copy(os.path.join(PYDIR, "spike", "spike.py"), folder)
    shutil.copy(os.path.join(HERE, "zz_slots.py"), folder)
    env = dict(os.environ)
    if WINDOWS:
        temp = os.path.join(tmp, "temp")
        os.makedirs(temp)
        env["TEMP"] = env["TMP"] = temp
        old = read_settings()
    else:
        env["WINEDEBUG"] = "-all"
        env.setdefault("WINEPREFIX", os.path.join(tempfile.gettempdir(), "asfpy-wine"))

    def logs():
        if WINDOWS:
            return glob.glob(os.path.join(temp, "asfpy.log"))
        return glob.glob(os.path.join(env["WINEPREFIX"], "drive_c", "users", "*", "**", "asfpy.log"), recursive=True)

    failed = 0
    try:
        for python in pythons:
            b = bits(python)
            print("=== %s (%d-bit, %s)" % (python, b, tool), flush=True)
            values = {"PythonVenv": native(python), "PythonFolder": native(folder)}
            if WINDOWS:
                write_settings(values)
            else:
                for name, value in values.items():
                    subprocess.run(["wine", "reg", "add", "HKCU\\" + KEY, "/v", name, "/d", value, "/f"], env=env,
                                   check=True, capture_output=True)
            for log in logs():
                os.remove(log)
            numpy = glob.glob(os.path.join(python, "Lib", "site-packages", "numpy"))
            cmd = ([] if WINDOWS else ["wine"]) + [exes[b], native(os.path.join(BUILD, "asfpy%d.xll" % b))]
            r = subprocess.run(cmd + (["numpy"] if numpy else []), env=env, capture_output=True, text=True,
                               encoding="utf-8", errors="replace", timeout=900)
            print(r.stdout, end="")
            if r.returncode:
                failed += 1
                print("exit code %d" % r.returncode)
                print(r.stderr[-2000:], end="")
                for log in logs():
                    print("--- %s" % log)
                    print(open(log, encoding="utf-8", errors="replace").read()[-4000:])
    finally:
        if WINDOWS:
            write_settings(old, old)
        shutil.rmtree(tmp, ignore_errors=True)
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
