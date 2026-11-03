"""Run the stdLambda/ASF benchmark on LibreOffice, headless.

    python3 Tools/stdLambda/BenchmarkLO.py [--out=DIR] [--engines=stdLambda,ASF] [--workloads=Loop,Recursion]

Builds the ASF library from src/ (see libreoffice/tools/build.py), adds stdLambda from stdVBA at
the pinned commit, patched to run on LibreOffice (see lo_stdlambda.py), and Benchmark.bas, then
measures every workload, operation and engine, rotating the engine order between cells. Writes
samples.csv and samples.json to --out (default Tools/stdLambda/benchmarks/latest-lo).
Needs LibreOffice, python3-uno, git and internet access.
"""
import json, os, platform, sys, time, shutil
from datetime import datetime, timezone

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, "libreoffice", "tools"))
sys.path.insert(0, os.path.join(ROOT, "Tools", "stdRegex"))
import build  # noqa: E402
import lo_stdlambda  # noqa: E402
from BenchmarkLO import start_office, sha256, cpu_name, git, uno_property  # noqa: E402

WORKLOADS = ["Arithmetic expression", "Predicate", "String functions", "Conditional", "Local variables",
             "Object members", "Loop", "While and branches", "Recursion", "Callback over 300 items"]
OPERATIONS = ["Compile", "Run"]
HEADER = "engine,workload,operation,status,note,iterations,sample_1_us,sample_2_us,sample_3_us,sample_4_us,sample_5_us"


def benchmark_module():
    """Benchmark.bas with LibreOffice's millisecond clock instead of QueryPerformanceCounter."""
    s = build.load(os.path.join(HERE, "Benchmark.bas"))
    start, end = s.index("#If VBA7 Then"), s.index("#End If") + len("#End If")
    s = s[:start] + s[end:]
    i = s.index("Private Function Seconds() As Double")
    j = s.index("End Function", i)
    s = s[:i] + "Private Function Seconds() As Double\n    Seconds = LOhelper.TimerValue()\n" + s[j:]
    return build.normal_module(s)


def main(argv):
    opts = {a.split("=", 1)[0]: a.split("=", 1)[1] for a in argv if a.startswith("--") and "=" in a}
    out_dir = os.path.abspath(opts.get("--out", os.path.join(HERE, "benchmarks", "latest-lo")))
    engines = opts.get("--engines", "stdLambda,ASF").split(",")
    workloads = opts["--workloads"].split(",") if "--workloads" in opts else WORKLOADS
    os.makedirs(out_dir, exist_ok=True)
    csv_path = os.path.join(out_dir, "samples.csv")
    mods = build.library_modules(ROOT) + lo_stdlambda.lo_modules(build) + [("LambdaBenchmark", benchmark_module(), "normal")]
    ctx, proc, work = start_office()
    try:
        libs = ctx.ServiceManager.createInstanceWithContext("com.sun.star.script.ApplicationScriptLibraryContainer", ctx)
        lib = libs.createLibrary(build.LIB_NAME)
        for name, code, _ in mods:
            lib.insertByName(name, code)
        libs.loadLibrary(build.LIB_NAME)
        provider = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.script.provider.MasterScriptProviderFactory", ctx).createScriptProvider("")
        measure = provider.getScript(f"vnd.sun.star.script:{build.LIB_NAME}.LambdaBenchmark.Measure?language=Basic&location=application")
        office = ctx.ServiceManager.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)
        node = office.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess",
                                                  (uno_property("nodepath", "/org.openoffice.Setup/Product"),))
        version = f"{node.getByName('ooName')} {node.getByName('ooSetupVersionAboutBox')}"
        with open(csv_path, "w", encoding="utf-8", newline="\n") as f:
            f.write(HEADER + "\n")
            cell = 0
            for workload in workloads:
                for operation in OPERATIONS:
                    for offset in range(len(engines)):
                        engine = engines[(cell + offset) % len(engines)]
                        row = measure.invoke((engine, workload, operation), (), ())[0]
                        f.write(row + "\n")
                        print(row, flush=True)
                    cell += 1
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
        operating_system=f"{platform.system()} {platform.release()}", cpu=cpu_name(), engines=engines,
        asf_commit=git("rev-parse", "HEAD").strip(), stdvba_commit=lo_stdlambda.STDVBA_COMMIT,
        benchmark_sha256=sha256(os.path.join(HERE, "Benchmark.bas")),
        samples=5, target_batch_ms=30, timer="GetSystemTicks (1 ms resolution)",
        stdlambda_patches="LibreOffice port (lo_stdlambda.py): LambdaDict for Scripting.Dictionary, ASF_RegexEngine for "
                          "VBScript.RegExp in the tokeniser, CallByName for rtcCallByName, Variant arrays for arrays of Types; "
                          "parser and evaluation loop unchanged",
        description="Compile: fresh object + compile (stdLambda protInit, ASF New + Compile). Run: one call of the compiled "
                    "object with the workload's arguments. Results checked before each timed cell.")
    with open(os.path.join(out_dir, "samples.json"), "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=2)
    print(f"Benchmark samples saved to {csv_path}")


if __name__ == "__main__":
    main(sys.argv[1:])
