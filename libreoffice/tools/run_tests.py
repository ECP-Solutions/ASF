"""Run the ASF test suites on the LibreOffice edition, headless.

    python3 libreoffice/tools/run_tests.py [regex|runner|all|bench|asfbench|fastbench] [--keep-log] [--vba-expressions=VBAExpressions.oxt] [--fastvm]

Builds the ASF library from src/ (see build.py), adds test/RegexTests.bas and an adapted
test/TestRunner.bas (or test/RegexBench.bas for bench, which prints timings), starts soffice with a throw-away profile and runs the suites through
UNO. Needs LibreOffice and its Python bindings (python3-uno). The VBA expression tests run
when the VBA Expressions extension (.oxt, or its unpacked VBAExpressionsLib folder) is
given. Exit status 0 when every test passes.
"""
import os, re, shutil, socket, subprocess, sys, tempfile, time, zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build  # noqa: E402

TEST_DIR = os.path.join(build.ROOT, "test")

# TestRunner tests that need VBAExpressionsLib
NEEDS_VBA_EXPRESSIONS = ("mutating_VBAexpressions_Arrays", "vbexpr_embedded", "reusing_results_from_VBAexpressions",
                         "write_into_array_from_VBAexpressions")

# TestRunner tests that cannot run on LibreOffice
EXCLUDED = {
    "calling_native_function": "Excel workbook, UDF of the VBA UDFunctions class",
    "host_sees_array_mutation": "UDF of the VBA UDFunctions class",
    "module_system_prototype_imports": "Excel Worksheet/Range prototypes",
    "module_system_prototype_single_import": "Excel Worksheet/Range prototypes",
    "module_system_prototype_namespace_import": "Excel Worksheet/Range prototypes",
    "module_system_prototype_side_effect_import": "Excel Worksheet/Range prototypes",
    "calling_native_function_app_level_2": "Excel workbook sheets",
    "monkey_patch_worksheets_name": "Excel workbook sheets",
}


def runner_module(skip):
    """TestRunner.bas for LibreOffice: test files come from test/, the tests in skip are left out."""
    s = build.load(os.path.join(TEST_DIR, "TestRunner.bas"))
    # the run log goes to the temporary folder, test data comes from test/
    s = s.replace('If ThisWorkbook.path = vbNullString Then Exit Sub', 'If LOLogDir() = vbNullString Then Exit Sub')
    s = s.replace('With New ASF: pLogPath = .JoinPath(ThisWorkbook.path, "TestRunner.log"): End With', 'pLogPath = LOLogDir() & "/TestRunner.log"')
    s = s.replace("ThisWorkbook.path", "LOTestDir()")
    s = re.sub(r'(wd & ")\\', r"\1/", s)
    for name in skip:
        s, n = re.subn(r'(?m)^(\s*)(If StartTest\("[^"]*",\s*"' + name + r'"\).*)$', r"\1' \2", s)
        assert n == 1, name
    return build.normal_module(s)


def fastvm_module(code):
    """The ASF class with the bytecode backend on in every engine, and each program the tree VM
    runs written to the TestRunner log, under the test that ran it"""
    code = code.replace("Private Sub INIT_()\n", "Private Sub INIT_()\n    FASTVM_ON_ = True\n", 1)
    code = code.replace("    If VM_.verbose Then\n        LogLine \"=== Runtime Log ===\"",
                        "    LOTraceBackend BACKEND_, FALLBACK_\n    If VM_.verbose Then\n        LogLine \"=== Runtime Log ===\"", 1)
    code = code.replace("        Compile = BC_.CompileSource(src, progName)\n",
                        "        Compile = BC_.CompileSource(src, progName)\n        LOTraceCompile Compile, BC_.Reason, src\n", 1)
    assert "LOTraceBackend" in code and "FASTVM_ON_ = True" in code and "LOTraceCompile" in code
    return code


def main_module(log_path, suites):
    calls = []
    if "regex" in suites:
        calls.append("    RunAllRegexTests")
    if "runner" in suites:
        calls.append("    RunAllTests")
    if "bench" in suites:
        calls.append("    RunRegexBenchmarks")
    if "asfbench" in suites:
        calls.append("    RunAsfBenchmarks")
    if "fastbench" in suites:
        calls.append("    RunFastVMBenchmarks")
    return f'''Option VBASupport 1
Public Function LOTestDir() As String
    LOTestDir = "{TEST_DIR}"
End Function
Public Function LOLogDir() As String
    LOLogDir = "{os.path.dirname(log_path)}"
End Function
Public Sub LOTraceBackend(ByVal backend As String, ByVal reason As String)
    Dim f As Integer: f = FreeFile
    Open LOLogDir() & "/TestRunner.log" For Append As #f
    If backend = "bytecode" Then Print #f, "BACKEND bytecode" Else Print #f, "BACKEND tree " & reason
    Close #f
End Sub
Public Sub LOTraceCompile(ByVal idx As Long, ByVal reason As String, ByVal src As String)
    Dim f As Integer: f = FreeFile
    Open LOLogDir() & "/TestRunner.log" For Append As #f
    If idx > 0 Then Print #f, "COMPILE direct" Else Print #f, "COMPILE ast " & reason
    Close #f
    f = FreeFile
    Open LOLogDir() & "/sources.txt" For Append As #f
    Print #f, "=====SRC====="
    Print #f, src
    Close #f
End Sub
Public Sub RunSuites()
    On Error GoTo EH
    SetOutputFile "{log_path}"
{chr(10).join(calls)}
    LogLine "=== SUITES DONE ==="
    SetOutputFile ""
    Exit Sub
EH:
    LogLine "=== SUITES ERROR #" & Err.Number & " " & Err.Description & " ==="
    SetOutputFile ""
End Sub
'''


def free_port():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def vba_expressions_dir(path, work):
    """Folder of the VBAExpressionsLib library, unpacking an .oxt into work."""
    if os.path.isdir(path):
        return os.path.abspath(path)
    with zipfile.ZipFile(path) as z:
        z.extractall(os.path.join(work, "vbaexpr"))
    return os.path.join(work, "vbaexpr", build.VBA_EXPRESSIONS_LIB)


def run(suites, keep_log=False, vba_expressions=None, fastvm=False):
    import uno  # noqa: F401  (python3-uno)
    from com.sun.star.connection import NoConnectException

    mods = build.library_modules()
    if fastvm:
        mods = [(n, fastvm_module(c) if n == "ASF" else c, k) for n, c, k in mods]
    if "regex" in suites:
        mods.append(("RegexTests", build.normal_module(build.load(os.path.join(TEST_DIR, "RegexTests.bas"))), "normal"))
    if "bench" in suites:
        mods.append(("RegexBench", build.normal_module(build.load(os.path.join(TEST_DIR, "RegexBench.bas"))), "normal"))
    if "asfbench" in suites:
        mods.append(("AsfBench", build.normal_module(build.load(os.path.join(TEST_DIR, "AsfBench.bas"))), "normal"))
    if "fastbench" in suites:
        mods.append(("FastVMBench", build.normal_module(build.load(os.path.join(TEST_DIR, "FastVMBench.bas"))), "normal"))
    skip = list(EXCLUDED) + ([] if vba_expressions else list(NEEDS_VBA_EXPRESSIONS))
    if "runner" in suites:
        mods.append(("TestRunner", runner_module(skip), "normal"))

    work = tempfile.mkdtemp(prefix="asf-lo-")
    log_path = os.path.join(work, "tests.log")
    mods.append(("LOTestMain", main_module(log_path, suites), "normal"))
    port = free_port()
    proc = subprocess.Popen(["soffice", "--headless", "--norestore", "--nologo", "--nodefault",
                             f"-env:UserInstallation=file://{work}/profile",
                             f"--accept=socket,host=127.0.0.1,port={port};urp;"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    ctx = None
    for _ in range(120):
        try:
            ctx = resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
            break
        except NoConnectException:
            time.sleep(0.5)
    if ctx is None:
        proc.kill()
        sys.exit("could not connect to soffice")
    try:
        libs = ctx.ServiceManager.createInstanceWithContext("com.sun.star.script.ApplicationScriptLibraryContainer", ctx)
        if vba_expressions:
            folder = vba_expressions_dir(vba_expressions, work)
            libs.createLibraryLink(build.VBA_EXPRESSIONS_LIB, uno.systemPathToFileUrl(os.path.join(folder, "script.xlb")), True)
        if libs.hasByName(build.LIB_NAME):
            libs.removeLibrary(build.LIB_NAME)
        lib = libs.createLibrary(build.LIB_NAME)
        for name, code, _ in mods:
            lib.insertByName(name, code)
        libs.loadLibrary(build.LIB_NAME)
        factory = ctx.ServiceManager.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", ctx)
        script = factory.createScriptProvider("").getScript(
            f"vnd.sun.star.script:{build.LIB_NAME}.LOTestMain.RunSuites?language=Basic&location=application")
        t0 = time.time()
        script.invoke((), (), ())
        elapsed = time.time() - t0
    finally:
        try:
            ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx).terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()
    log = open(log_path, encoding="utf-8", errors="replace").read() if os.path.exists(log_path) else ""
    runner_log = os.path.join(work, "TestRunner.log")
    if fastvm and os.path.exists(runner_log):
        log += "\n" + open(runner_log, encoding="utf-8", errors="replace").read()
    sources = os.path.join(work, "sources.txt")
    if os.environ.get("ASF_CAPTURE_SOURCES") and os.path.exists(sources):
        shutil.copy(sources, os.environ["ASF_CAPTURE_SOURCES"])
    if not keep_log:
        shutil.rmtree(work, ignore_errors=True)
    return log, elapsed, log_path, len(skip)


def backend_report(log):
    """Programs each backend ran, and the reasons of the fallbacks with the tests that had them"""
    counts, reasons, test = {"bytecode": 0, "tree": 0}, {}, "?"
    for line in log.split("\n"):
        if line.startswith("RUN  "):
            test = line[5:].strip()
        elif line.startswith("BACKEND bytecode"):
            counts["bytecode"] += 1
        elif line.startswith("BACKEND tree"):
            counts["tree"] += 1
            reason = line[len("BACKEND tree"):].strip() or "(none)"
            reasons.setdefault(reason, []).append(test)
    compiled = {"direct": 0, "ast": 0}
    why = {}
    for line in log.split("\n"):
        if line.startswith("COMPILE direct"):
            compiled["direct"] += 1
        elif line.startswith("COMPILE ast"):
            compiled["ast"] += 1
            why.setdefault(line[len("COMPILE ast"):].strip() or "(none)", 0)
            why[line[len("COMPILE ast"):].strip() or "(none)"] += 1
    if compiled["direct"] + compiled["ast"]:
        print(f"compiled from tokens: {compiled['direct']}, through the AST: {compiled['ast']}")
        for reason, n in sorted(why.items(), key=lambda kv: -kv[1]):
            print(f"  {n:4d}  {reason}")
    total = counts["bytecode"] + counts["tree"]
    print(f"programs: {total}, bytecode {counts['bytecode']} ({100 * counts['bytecode'] / max(total, 1):.0f}%), tree {counts['tree']}")
    for reason, tests in sorted(reasons.items(), key=lambda kv: -len(kv[1])):
        uniq = list(dict.fromkeys(tests))
        print(f"  {len(tests):4d}  {reason}: {', '.join(uniq[:4])}{' ...' if len(uniq) > 4 else ''}")


def summarize(log):
    ok = "=== SUITES DONE ===" in log
    failures = [l for l in log.split("\n") if l.startswith(("[FAIL]", "FAIL "))]
    totals = [l for l in log.split("\n") if l.startswith(("Total:", "=== ")) and ("passed" in l or "Passed" in l)]
    return ok, failures, totals


def main(argv):
    which = argv[1] if len(argv) > 1 and not argv[1].startswith("--") else "all"
    suites = {"regex": ["regex"], "runner": ["runner"], "all": ["regex", "runner"], "bench": ["bench"], "asfbench": ["asfbench"],
              "fastbench": ["fastbench"]}[which]
    vbe = next((a.split("=", 1)[1] for a in argv if a.startswith("--vba-expressions=")), os.environ.get("ASF_VBA_EXPRESSIONS"))
    fastvm = "--fastvm" in argv
    log, elapsed, log_path, skipped = run(suites, keep_log="--keep-log" in argv, vba_expressions=vbe, fastvm=fastvm)
    ok, failures, totals = summarize(log)
    if fastvm:
        backend_report(log)
    if "bench" in suites or "asfbench" in suites or "fastbench" in suites:
        print("\n".join(l for l in log.split("\n") if " | " in l))
    for l in totals:
        print(l)
    for l in failures:
        print(l)
    print(f"skipped: {skipped if 'runner' in suites else 0}" + ("" if vbe else " (no --vba-expressions)"))
    print(f"{elapsed:.1f} s" + (f", log kept at {log_path}" if "--keep-log" in argv else ""))
    if not ok:
        print("suite run did not complete:", log.strip().split("\n")[-1:] or "no output")
        return 1
    return 0 if not failures else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
