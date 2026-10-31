"""Run RE2's test cases on the LibreOffice edition, headless, and compare with RE2's results.

    python3 test/re2-cases/tools/run_lo.py [SUITE ...] [--work=DIR] [--out=DIR] [--show=N]

SUITE is a suite number (000 .. 026, see README.md); default all. Run tools/prepare.py first.
Builds the engine from src/ (see libreoffice/tools/build.py), runs Re2Cases.bas in a headless
LibreOffice with a throw-away profile, writes OUT/NNN.txt (default test/re2-cases/out) and
compares it with WORK/expected/NNN.txt (see compare.py). Needs LibreOffice and python3-uno.
"""
import importlib.util, os, shutil, sys, time

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)
ROOT = os.path.dirname(os.path.dirname(BASE))
sys.path.insert(0, os.path.join(ROOT, "libreoffice", "tools"))
sys.path.insert(0, HERE)
import build    # noqa: E402
import compare  # noqa: E402

# soffice start-up shared with the regex corpora
_spec = importlib.util.spec_from_file_location("corpus_run_lo", os.path.join(ROOT, "test", "regex-corpus", "tools", "run_lo.py"))
corpus_run_lo = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(corpus_run_lo)


def main(argv):
    opts = {a.split("=", 1)[0]: a.split("=", 1)[1] for a in argv if a.startswith("--") and "=" in a}
    work = os.path.abspath(opts.get("--work", os.path.join(BASE, "work")))
    out_dir = os.path.abspath(opts.get("--out", os.path.join(BASE, "out")))
    show = int(opts.get("--show", "5"))
    suites = [a for a in argv if not a.startswith("--")] or compare.suites(work)
    os.makedirs(out_dir, exist_ok=True)

    mods = build.library_modules(ROOT)
    mods.append(("Re2Cases", build.normal_module(build.load(os.path.join(BASE, "Re2Cases.bas")), ROOT), "normal"))
    ctx, proc, profile = corpus_run_lo.start_office()
    try:
        libs = ctx.ServiceManager.createInstanceWithContext("com.sun.star.script.ApplicationScriptLibraryContainer", ctx)
        lib = libs.createLibrary(build.LIB_NAME)
        for name, code, _ in mods:
            lib.insertByName(name, code)
        libs.loadLibrary(build.LIB_NAME)
        provider = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.script.provider.MasterScriptProviderFactory", ctx).createScriptProvider("")
        script = provider.getScript(f"vnd.sun.star.script:{build.LIB_NAME}.Re2Cases.RunRe2Cases?language=Basic&location=application")
        for num in suites:
            t0 = time.time()
            script.invoke((os.path.join(work, "cases", num + ".txt"), os.path.join(out_dir, num + ".txt")), (), ())
            print(f"{num}: ran in {time.time() - t0:.1f} s", flush=True)
    finally:
        try:
            ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx).terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()
        shutil.rmtree(profile, ignore_errors=True)
    return compare.main([work, out_dir] + suites + [f"--show={show}"])


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
