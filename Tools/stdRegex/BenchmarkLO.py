"""Run the regex benchmark on the LibreOffice edition, headless.

    python3 Tools/stdRegex/BenchmarkLO.py [--out=DIR] [--engines=ASF,ASF v3.2.0,ASF 290ad3e]

Builds the ASF library from src/ (see libreoffice/tools/build.py), adds the ASF regex engines
of v3.2.0 and 290ad3e from git history under their own class names, and Benchmark.bas, then
measures every scenario, operation and engine in the order Benchmark.ps1 uses in Excel.
Writes samples.csv and samples.json to --out (default Tools/stdRegex/benchmarks/latest-lo).
The stdRegex engines need Windows (VBScript, Scripting.Dictionary) and run in Excel only.
Needs LibreOffice, python3-uno and git.
"""
import hashlib, json, os, platform, socket, subprocess, sys, tempfile, time, shutil
from datetime import datetime, timezone

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, "libreoffice", "tools"))
import build  # noqa: E402

SCENARIOS = ["Literal scan", "Short validation", "Capture extraction", "Alternation and ignore case",
             "Multiline anchors", "No-match scan", "Word boundaries", "Backreference"]
OPERATIONS = ["Compile + first Test", "Test", "MatchAll", "Replace"]
OLD_ASF = {"ASF v3.2.0": ("v3.2.0", "ASF_RegexEngine_v320"), "ASF 290ad3e": ("290ad3e", "ASF_RegexEngine_290ad3e")}
HEADER = "engine,scenario,operation,input_chars,expected_matches,status,note,iterations,sample_1_us,sample_2_us,sample_3_us,sample_4_us,sample_5_us"
STUB = '''Public Sub protInit(ByVal pattern As String, ByVal flags As String, Optional ByVal stepsLimit As Long = 0)
    Err.Raise 5, "{0}", "{0} needs Windows: run the benchmark in Excel"
End Sub
'''


def git(*args):
    return subprocess.run(["git", "-C", ROOT, *args], check=True, capture_output=True, text=True, encoding="latin-1").stdout


def benchmark_module():
    """Benchmark.bas with LibreOffice's millisecond clock instead of QueryPerformanceCounter."""
    s = build.load(os.path.join(HERE, "Benchmark.bas"))
    start, end = s.index("#If VBA7 Then"), s.index("#End If") + len("#End If")
    s = s[:start] + s[end:]
    i = s.index("Private Function Seconds() As Double")
    j = s.index("End Function", i)
    s = s[:i] + "Private Function Seconds() As Double\n    Seconds = LOhelper.TimerValue()\n" + s[j:]
    return build.normal_module(s)


def modules():
    mods = build.library_modules(ROOT)
    pub = build.public_collections(ROOT)
    for rev, cls in OLD_ASF.values():
        text = git("show", f"{rev}:src/ASF_RegexEngine.cls").replace("\r\n", "\n")
        text = text.replace("ASF_RegexEngine", cls)
        # the passes for the regex engine, as for src/ASF_RegexEngine.cls
        body = build.member_collections(build.asf_passes(build.REGEX_ENGINE, build.generic(text)), pub)
        mods.append((cls, build.class_module(cls, body), "class"))
    for cls in ("stdRegex", "legacyStdRegex"):
        mods.append((cls, build.class_module(cls, STUB.format(cls)), "class"))
    mods.append(("RegexBenchmark", benchmark_module(), "normal"))
    return mods


def start_office():
    import uno
    from com.sun.star.connection import NoConnectException
    work = tempfile.mkdtemp(prefix="asf-bench-")
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


def sha256(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest().upper()


def cpu_name():
    try:
        with open("/proc/cpuinfo") as f:
            return next(l.split(":", 1)[1].strip() for l in f if l.startswith("model name"))
    except (OSError, StopIteration):
        return platform.processor()


def main(argv):
    opts = {a.split("=", 1)[0]: a.split("=", 1)[1] for a in argv if a.startswith("--") and "=" in a}
    out_dir = os.path.abspath(opts.get("--out", os.path.join(HERE, "benchmarks", "latest-lo")))
    engines = opts.get("--engines", "ASF,ASF v3.2.0,ASF 290ad3e").split(",")
    os.makedirs(out_dir, exist_ok=True)
    csv_path = os.path.join(out_dir, "samples.csv")
    ctx, proc, work = start_office()
    try:
        libs = ctx.ServiceManager.createInstanceWithContext("com.sun.star.script.ApplicationScriptLibraryContainer", ctx)
        lib = libs.createLibrary(build.LIB_NAME)
        for name, code, _ in modules():
            lib.insertByName(name, code)
        libs.loadLibrary(build.LIB_NAME)
        provider = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.script.provider.MasterScriptProviderFactory", ctx).createScriptProvider("")
        script = lambda m: provider.getScript(f"vnd.sun.star.script:{build.LIB_NAME}.RegexBenchmark.{m}?language=Basic&location=application")
        measure, allocations = script("Measure"), script("MeasureAllocations")
        office = ctx.ServiceManager.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)
        node = office.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess",
                                                  (uno_property("nodepath", "/org.openoffice.Setup/Product"),))
        version = f"{node.getByName('ooName')} {node.getByName('ooSetupVersionAboutBox')}"
        with open(csv_path, "w", encoding="utf-8", newline="\n") as f:
            f.write(HEADER + "\n")
            cell = 0
            for scenario in SCENARIOS:
                for operation in OPERATIONS:
                    # rotate the engine order, as Benchmark.ps1 does
                    for offset in range(len(engines)):
                        engine = engines[(cell + offset) % len(engines)]
                        row = measure.invoke((engine, scenario, operation), (), ())[0]
                        f.write(row + "\n")
                        print(row, flush=True)
                    cell += 1
            row = allocations.invoke(("Collection",), (), ())[0]
            f.write(row + "\n")
            print(row)
    finally:
        try:
            ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx).terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()
        shutil.rmtree(work, ignore_errors=True)
    metadata = dict(
        measured_at_utc=datetime.now(timezone.utc).isoformat(), office=version,
        operating_system=f"{platform.system()} {platform.release()}", cpu=cpu_name(),
        engines=engines,
        asf_commit=git("rev-parse", "HEAD").strip(), asf_sha256=sha256(os.path.join(ROOT, "src", "ASF_RegexEngine.cls")),
        asf_v320_commit=git("rev-parse", "v3.2.0^{commit}").strip(),
        asf_290ad3e_commit=git("rev-parse", "290ad3e^{commit}").strip(),
        benchmark_sha256=sha256(os.path.join(HERE, "Benchmark.bas")),
        samples=5, target_batch_ms=30, timer="GetSystemTicks (1 ms resolution)",
        engine_step_limits=2147483647,
        description="Public API timings of the ASF regex engines on the LibreOffice edition: fresh object + Init + first Test for compile; "
                    "reused preinitialized objects for other operations. Correctness checked before each timed cell.")
    with open(os.path.join(out_dir, "samples.json"), "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=2)
    print(f"Benchmark samples saved to {csv_path}")


def uno_property(name, value):
    from com.sun.star.beans import PropertyValue
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


if __name__ == "__main__":
    main(sys.argv[1:])
