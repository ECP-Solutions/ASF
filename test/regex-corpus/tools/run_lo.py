"""Run the regex corpora on the LibreOffice edition, headless, and compare with the expected output.

    python3 test/regex-corpus/tools/run_lo.py [js|full] [CORPUS ...] [--repo=PATH] [--out=DIR] [--show=N]

js (default) runs RunCorpusJs and compares with expected/js; full runs RunCorpusFull and
compares with expected/full.zip. CORPUS is c1, c2, c3 or extra (default: c1 c2 c3; extra has
no JavaScript output). --repo builds the engine from another checkout's src/ (to produce the
output of another version); --out is where outputs go (default test/regex-corpus/out).
Needs LibreOffice and python3-uno. Exit status 0 when every case matches.
"""
import os, socket, subprocess, sys, tempfile, time, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
CORPUS_DIR = os.path.dirname(HERE)
ROOT = os.path.dirname(os.path.dirname(CORPUS_DIR))
sys.path.insert(0, os.path.join(ROOT, "libreoffice", "tools"))
sys.path.insert(0, HERE)
import build    # noqa: E402
import compare  # noqa: E402


def start_office():
    import uno
    from com.sun.star.connection import NoConnectException
    work = tempfile.mkdtemp(prefix="asf-corpus-")
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        port = s.getsockname()[1]
    proc = subprocess.Popen(["soffice", "--headless", "--norestore", "--nologo", "--nodefault",
                             f"-env:UserInstallation=file://{work}/profile",
                             f"--accept=socket,host=127.0.0.1,port={port};urp;"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(120):
        try:
            return resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"), proc, work
        except NoConnectException:
            time.sleep(0.5)
    proc.kill()
    sys.exit("could not connect to soffice")


def main(argv):
    opts = {a.split("=", 1)[0]: a.split("=", 1)[1] for a in argv if a.startswith("--") and "=" in a}
    args = [a for a in argv if not a.startswith("--")]
    mode = args.pop(0) if args and args[0] in ("js", "full") else "js"
    corpora = args or ["c1", "c2", "c3"]
    repo = os.path.abspath(opts.get("--repo", ROOT))
    out_dir = os.path.abspath(opts.get("--out", os.path.join(CORPUS_DIR, "out")))
    show = int(opts.get("--show", "10"))
    os.makedirs(out_dir, exist_ok=True)

    mods = build.library_modules(repo)
    mods.append(("RegexCorpus", build.normal_module(build.load(os.path.join(CORPUS_DIR, "RegexCorpus.bas")), repo), "normal"))
    ctx, proc, work = start_office()
    bad = 0
    try:
        libs = ctx.ServiceManager.createInstanceWithContext("com.sun.star.script.ApplicationScriptLibraryContainer", ctx)
        lib = libs.createLibrary(build.LIB_NAME)
        for name, code, _ in mods:
            lib.insertByName(name, code)
        libs.loadLibrary(build.LIB_NAME)
        provider = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.script.provider.MasterScriptProviderFactory", ctx).createScriptProvider("")
        macro = "RunCorpusJs" if mode == "js" else "RunCorpusFull"
        script = provider.getScript(f"vnd.sun.star.script:{build.LIB_NAME}.RegexCorpus.{macro}?language=Basic&location=application")
        for c in corpora:
            cases = os.path.join(CORPUS_DIR, "cases", c + ".txt")
            out = os.path.join(out_dir, f"{c}.{mode}.txt")
            t0 = time.time()
            script.invoke((cases, out), (), ())
            print(f"{c}: ran in {time.time() - t0:.1f} s")
            expected = (os.path.join(CORPUS_DIR, "expected", "js", c + ".js.txt") if mode == "js"
                        else os.path.join(CORPUS_DIR, "expected", "full.zip", c + ".full.txt"))
            if mode == "js" and not os.path.exists(expected):
                print(f"{c}: no JavaScript output to compare with")
                continue
            bad += compare.compare(cases, expected, out, show)
    finally:
        try:
            ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx).terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()
        shutil.rmtree(work, ignore_errors=True)
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
