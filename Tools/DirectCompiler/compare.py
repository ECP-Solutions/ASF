"""The direct compiler (ASF_Bytecode.CompileSource: tokens -> bytecode) against the AST pipeline
(ASF_Compiler, then ASF_Bytecode.Prepare), in LibreOffice.

    python3 Tools/DirectCompiler/compare.py                      # parity on every source, then timings
    python3 Tools/DirectCompiler/compare.py -v                   # also print the code of every difference
    python3 Tools/DirectCompiler/compare.py --sources FILE       # also the sources of FILE (programs after lines =====SRC=====)
    python3 Tools/DirectCompiler/compare.py --src "let a = 1;"   # these sources only
    python3 Tools/DirectCompiler/compare.py --export             # write corpus.txt and the benchmark scripts for VBA

Parity: for every source (generated programs, the fixed ones below, the scripts of test/), each in
fresh globals, the disassembly of the direct compiler must be the one ASF_Bytecode makes from the
AST. run_tests.py writes the sources the TestRunner compiles with ASF_CAPTURE_SOURCES=FILE --fastvm.
Timings: milliseconds to compile each benchmark script, tokenizer included, on both paths.
"""
import glob, os, random, shutil, sys, time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, "libreoffice", "tools"))
sys.path.insert(0, os.path.join(ROOT, "Tools", "stdRegex"))
import build
from BenchmarkLO import start_office

HARNESS = r'''Option VBASupport 1
Function Parity(ByVal src As String) As String
    On Error GoTo Failed
    Dim phase As String
    phase = "AST"
    Dim g As New ASF_Globals: g.ASF_InitGlobals
    Dim c As New ASF_Compiler: c.SetGlobals g
    Dim idx As Long: idx = c.CompileProgram(src)
    Dim bc As New ASF_Bytecode: bc.SetGlobals g
    Dim tree As String: tree = bc.Disassemble(idx)
    phase = "direct"
    Dim g2 As New ASF_Globals: g2.ASF_InitGlobals
    Dim d As New ASF_Bytecode: d.SetGlobals g2
    Dim idx2 As Long: idx2 = d.CompileSource(src, "@anon")
    If idx2 = 0 Then
        Parity = "NOT TAKEN: " & d.Reason
        Exit Function
    End If
    Dim direct As String: direct = d.Disassemble(idx2)
    If direct = tree Then Parity = "same" Else Parity = "AST:" & vbLf & tree & vbLf & "DIRECT:" & vbLf & direct
    Exit Function
Failed:
    Parity = "ERROR " & phase & " " & Err.Number & " " & Err.Description
    If phase = "AST" Then
        ' a source ASF_Compiler rejects must go to it, which raises the error
        Dim g3 As New ASF_Globals: g3.ASF_InitGlobals
        Dim d3 As New ASF_Bytecode: d3.SetGlobals g3
        If d3.CompileSource(src, "@anon") > 0 Then Parity = "TAKEN BUT INVALID: " & Parity
    End If
End Function
Function Timing(ByVal src As String, ByVal n As Long) As String
    On Error GoTo Failed
    Dim k As Long, t0 As Long, tAst As Long, tDirect As Long, tTok As Long
    For k = 1 To n
        Dim g As ASF_Globals: Set g = New ASF_Globals: g.ASF_InitGlobals
        Dim p As ASF_Parser: Set p = New ASF_Parser: p.SetGlobals g
        t0 = GetSystemTicks(): p.Tokenize src: tTok = tTok + GetSystemTicks() - t0
        t0 = GetSystemTicks()
        Dim c As ASF_Compiler: Set c = New ASF_Compiler: c.SetGlobals g
        Dim idx As Long: idx = c.CompileProgram(src)
        Dim bc As ASF_Bytecode: Set bc = New ASF_Bytecode: bc.SetGlobals g
        bc.Prepare idx
        tAst = tAst + GetSystemTicks() - t0
        Dim g2 As ASF_Globals: Set g2 = New ASF_Globals: g2.ASF_InitGlobals
        Dim d As ASF_Bytecode: Set d = New ASF_Bytecode: d.SetGlobals g2
        t0 = GetSystemTicks(): d.CompileSource src, "@anon": tDirect = tDirect + GetSystemTicks() - t0
    Next k
    Timing = Format(tTok / n, "0.0") & "|" & Format(tAst / n, "0.0") & "|" & Format(tDirect / n, "0.0")
    Exit Function
Failed:
    Timing = "ERROR " & Err.Number & " " & Err.Description
End Function
'''

FIXED = [
    "let n = 0; n = n + 1; n += 2; return n;",
    "let a = [1, 2]; let o = {x: 1}; o.x = a[1] + o.x * -3; print(o.x, 'hi'); return a.length;",
    "let n = 0; for (let i = 1, i <= 10, i += 1) { if (i % 2 == 0) { n = n + i; } elseif (i == 3) { n = 0; } else { n -= 1; } }; while (n > 0) { n = n - 1; }; return n;",
    "let s = 2 * 3 + 4; let b = !(s > 3) && true || null ?? 5; let t = s > 1 ? 'a' : 'b'; return typeof s;",
    "let x = 1; if (x > 0 && (x < 5 || x == 7)) { x = 2; }; if (!(x > 1 && x < 3)) { x = 9; }; return x;",
    "let m = [[1, 2], [3, 4]]; m[1][2] = 9; m[2] = 5; let z = m[1][2] + m[2]; return z;",
    "let a = 'x' + 'y'; let b = -(2 * 3); let c = 2 ^ 3 ^ 2; let d = 10 - 2 - 3; return a & b & c & d;",
    "let q = 5; q *= 2 + 1; q /= 3; q -= 1; print(q); return q > 3 ? q < 9 ? 'mid' : 'big' : 'small';",
]


def gen_expr(rnd, names, depth=0):
    r = rnd.random()
    if depth > 3 or r < 0.3:
        c = rnd.random()
        if c < 0.4:
            return rnd.choice(names)
        if c < 0.8:
            return str(rnd.randint(0, 9))
        return rnd.choice(["'s'", "true", "false", "null"])
    if r < 0.4:
        return rnd.choice([f"arr[{rnd.randint(1, 3)}]", f"arr[{rnd.choice(names)} + 1]", "m[1][2]", f"o.{rnd.choice(['p', 'q'])}", "arr.length"])
    if r < 0.75:
        op = rnd.choice(["+", "-", "*", "/", "%", "<", "<=", ">", ">=", "==", "!=", "&&", "||", "&", "??"])
        return f"{gen_expr(rnd, names, depth + 1)} {op} {gen_expr(rnd, names, depth + 1)}"
    if r < 0.85:
        return f"({gen_expr(rnd, names, depth + 1)})"
    if r < 0.92:
        return f"-{gen_expr(rnd, names, depth + 3)}" if rnd.random() < 0.5 else f"!{gen_expr(rnd, names, depth + 3)}"
    return f"{gen_expr(rnd, names, depth + 1)} ? {gen_expr(rnd, names, depth + 1)} : {gen_expr(rnd, names, depth + 1)}"


def gen_block(rnd, names, depth=0, n=None):
    out = []
    for _ in range(n or rnd.randint(1, 4)):
        r = rnd.random()
        v = rnd.choice(names)
        if depth < 2 and r < 0.15:
            out.append(f"if ({gen_expr(rnd, names)}) {{ {gen_block(rnd, names, depth + 1)} }}"
                       + (f" else {{ {gen_block(rnd, names, depth + 1)} }}" if rnd.random() < 0.5 else "") + ";")
        elif depth < 2 and r < 0.22:
            out.append(f"while ({gen_expr(rnd, names)}) {{ {gen_block(rnd, names, depth + 1)} }};")
        elif depth < 2 and r < 0.3:
            out.append(f"for (let i = 1, i <= {rnd.randint(1, 5)}, i += 1) {{ {gen_block(rnd, names, depth + 1)} }};")
        elif r < 0.4:
            out.append(f"{v} {rnd.choice(['+=', '-=', '*='])} {gen_expr(rnd, names)};")
        elif r < 0.5:
            out.append(f"{v} = {v} + {rnd.randint(1, 9)};")
        elif r < 0.55:
            out.append(f"print({gen_expr(rnd, names)});")
        elif r < 0.65:
            out.append(rnd.choice([f"arr[{rnd.randint(1, 3)}] = {gen_expr(rnd, names)};", f"arr[{v} + 1] = {gen_expr(rnd, names)};",
                                   f"m[{rnd.randint(1, 2)}][2] = {gen_expr(rnd, names)};", f"o.p = {gen_expr(rnd, names)};",
                                   f"arr[2] += {rnd.randint(1, 5)};"]))
        else:
            out.append(f"{v} = {gen_expr(rnd, names)};")
    return " ".join(out)


def corpus(count=150, seed=7):
    rnd = random.Random(seed)
    progs = list(FIXED)
    for _ in range(count):
        names = rnd.sample(["a", "b", "c", "x", "y", "total"], 3)
        init = " ".join(f"let {n} = {rnd.randint(0, 5)};" for n in names) + " let arr = [1, 2, 3]; let m = [[1, 2], [3, 4]]; let o = {p: 1, q: 'z'};"
        progs.append(f"{init} {gen_block(rnd, names, 0, rnd.randint(2, 6))} return {names[0]};")
    return progs


def gen_rich_expr(rnd, names, depth=0):
    r = rnd.random()
    if depth > 2 or r < 0.25:
        return gen_expr(rnd, names, 2)
    if r < 0.35:
        return f"`t${{{gen_expr(rnd, names, 2)}}}-${{{rnd.choice(names)}}}`"
    if r < 0.45:
        return f"[{gen_rich_expr(rnd, names, depth + 1)}, ...arr, {rnd.choice(names)}]"
    if r < 0.52:
        return f"{{k: {gen_rich_expr(rnd, names, depth + 1)}, ...o, 'z': 1}}"
    if r < 0.62:
        return f"f{rnd.randint(1, 2)}({gen_rich_expr(rnd, names, depth + 1)}, {rnd.choice(names)})"
    if r < 0.7:
        return f"arr.{rnd.choice(['map', 'filter'])}(fun(v) {{ return v + {rnd.choice(names)}; }})"
    if r < 0.76:
        return f"fun(p, ...q) {{ {rnd.choice(names)} = p; return q.length + {rnd.choice(names)}; }}"
    if r < 0.82:
        return f"typeof {rnd.choice(names)} == 'number' ? {gen_rich_expr(rnd, names, depth + 1)} : null"
    if r < 0.88:
        return f"new P({gen_rich_expr(rnd, names, depth + 1)}).get({rnd.randint(1, 3)})"
    if r < 0.94:
        return f"o.p.toString().length + range(1, 3).length"
    return f"(fun() {{ let w = {rnd.choice(names)}; return fun() {{ w += 1; return w; }}; }})()()"


def gen_rich_block(rnd, names, depth=0, n=None):
    out = []
    for _ in range(n or rnd.randint(1, 4)):
        r = rnd.random()
        v = rnd.choice(names)
        e = gen_rich_expr(rnd, names)
        if depth < 2 and r < 0.1:
            out.append(f"try {{ {gen_rich_block(rnd, names, depth + 1)} if ({v}) {{ throw `x${{{v}}}`; }}; }} catch (err) {{ {v} = err; }};")
        elif depth < 2 and r < 0.18:
            out.append(f"switch ({v}) {{ case 1 {{ {gen_rich_block(rnd, names, depth + 1)} break; }} case 'a' {{ {v} = 2; }} default {{ {v} = 0; }} }};")
        elif depth < 2 and r < 0.26:
            out.append(f"for ({rnd.choice(['it', 'k'])} {rnd.choice(['of', 'in'])} {rnd.choice(['arr', 'o', 'm[1]'])}) {{ {gen_rich_block(rnd, names, depth + 1)} if (it) {{ continue; }}; }};")
        elif depth < 2 and r < 0.32:
            out.append(f"if ({gen_expr(rnd, names)}) {{ {gen_rich_block(rnd, names, depth + 1)} }} elseif ({gen_expr(rnd, names)}) {{ {v} = 1; }} else {{ {gen_rich_block(rnd, names, depth + 1)} }};")
        elif depth < 2 and r < 0.38:
            out.append(f"while ({gen_expr(rnd, names)}) {{ {gen_rich_block(rnd, names, depth + 1)} break; }};")
        elif r < 0.45:
            out.append(f"[{v}, {rnd.choice(names)}, ...rest] = {e};")
        elif r < 0.52:
            out.append(f"arr.push({e}); arr[{v} + 1] = {e}; m[1][{rnd.randint(1, 2)}] += {rnd.randint(1, 3)};")
        elif r < 0.58:
            out.append(f"o.p = {e}; o[{rnd.choice(['\'q\'', v])}] = {v};")
        elif r < 0.64:
            out.append(f"print({e}, {v});")
        elif r < 0.7:
            out.append(f"{v} {rnd.choice(['+=', '-=', '*=', '&='])} {e};")
        else:
            out.append(f"{v} = {e};")
    return " ".join(out)


def rich_corpus(count=150, seed=11):
    """Programs with functions, closures, classes, templates, try, switch and loops over values"""
    rnd = random.Random(seed)
    progs = []
    for _ in range(count):
        names = rnd.sample(["a", "b", "c", "x", "y", "total"], 3)
        head = (" ".join(f"let {n} = {rnd.randint(0, 5)};" for n in names)
                + " let arr = [1, 2, 3]; let m = [[1, 2], [3, 4]]; let o = {p: 1, q: 'z'}; let rest = [];"
                + " class P { field v = 0, w; constructor(v) { this.v = v; } get(d) { return this.v + d; } };"
                + " class Q extends P { constructor(v) { super(v * 2); } get(d) { return d; } };"
                + f" fun f1(u, v) {{ {gen_rich_block(rnd, ['u', 'v', names[0]], 1, 2)} return u; }};"
                + f" fun f2(u, ...more) {{ let s = u; for (z of more) {{ s += z; }}; return s; }};")
        progs.append(f"{head} {gen_rich_block(rnd, names, 0, rnd.randint(2, 6))} return {names[0]};")
    return progs


def test_scripts():
    """The .vas scripts of test/"""
    out = []
    for path in sorted(glob.glob(os.path.join(ROOT, "test", "**", "*.vas"), recursive=True)):
        with open(path, encoding="utf-8", errors="replace") as f:
            out.append(f.read().replace("\r\n", "\n"))
    return out


def captured(path):
    """The sources run_tests.py captured (ASF_CAPTURE_SOURCES)"""
    with open(path, encoding="utf-8", errors="replace") as f:
        text = f.read().replace("\r\n", "\n")
    return [p[1:-1] if p.startswith("\n") and p.endswith("\n") else p.strip("\n") for p in text.split("=====SRC=====") if p.strip()]


def bench_scripts(rnd=random.Random(3)):
    long_simple = "let n = 0; " + "n = n + 1; " * 300 + "return n;"
    names = ["a", "b", "c", "x", "y", "total"]
    varied = " ".join(f"let {n} = 1;" for n in names) + " " + gen_block(rnd, names, 0, 300) + " return a;"
    return [("300 x n = n + 1", long_simple), ("300 varied statements", varied)]


def export():
    """The corpus (programs separated by lines of ----) and the benchmark scripts, for DirectCompare.bas"""
    with open(os.path.join(HERE, "corpus.txt"), "w", newline="\n") as f:
        f.write("\n----\n".join(corpus() + rich_corpus()) + "\n")
    for i, (name, src) in enumerate(bench_scripts(), 1):
        with open(os.path.join(HERE, f"bench{i}.vas"), "w", newline="\n") as f:
            f.write(src + "\n")
    print("wrote corpus.txt, bench1.vas, bench2.vas")


def main():
    if "--export" in sys.argv:
        export()
        return
    verbose = "-v" in sys.argv
    ctx, proc, work = start_office()
    try:
        libs = ctx.ServiceManager.createInstanceWithContext("com.sun.star.script.ApplicationScriptLibraryContainer", ctx)
        lib = libs.createLibrary("ASF")
        for n, c, _ in build.library_modules(ROOT):
            lib.insertByName(n, c)
        lib.insertByName("DirectHarness", HARNESS)
        libs.loadLibrary("ASF")
        sp = ctx.ServiceManager.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", ctx).createScriptProvider("")
        parity = sp.getScript("vnd.sun.star.script:ASF.DirectHarness.Parity?language=Basic&location=application")
        timing = sp.getScript("vnd.sun.star.script:ASF.DirectHarness.Timing?language=Basic&location=application")
        same = taken = errors = 0
        reasons = {}
        if "--src" in sys.argv:
            progs = [a for a in sys.argv[sys.argv.index("--src") + 1:] if not a.startswith("-")]
        elif "--rich" in sys.argv:
            progs = rich_corpus(int(sys.argv[sys.argv.index("--rich") + 1]), int(sys.argv[sys.argv.index("--rich") + 2]))
        elif "--only" in sys.argv:
            progs = FIXED[:int(sys.argv[sys.argv.index("--only") + 1])]
        else:
            progs = corpus() + rich_corpus() + test_scripts()
            if "--sources" in sys.argv:
                progs += captured(sys.argv[sys.argv.index("--sources") + 1])
            progs = list(dict.fromkeys(progs))
        for src in progs:
            r = parity.invoke((src,), (), ())[0]
            if r.startswith("NOT TAKEN"):
                reasons[r] = reasons.get(r, 0) + 1
                if verbose:
                    print(r, "|", src[:100].replace("\n", " "))
                continue
            if r.startswith("TAKEN BUT INVALID"):
                print(r, "|", src[:300].replace("\n", " "))
                errors += 1
                continue
            if r.startswith("ERROR"):
                errors += 1
                if verbose or not r.startswith("ERROR AST"):
                    print(r, "|", src[:100].replace("\n", " "))
                continue
            taken += 1
            if r == "same":
                same += 1
            else:
                print("DIFF:", src)
                if verbose:
                    print(r)
        print(f"parity: {same}/{taken} identical, {len(progs) - taken - errors} not taken, {errors} errors")
        for reason, n in sorted(reasons.items(), key=lambda kv: -kv[1]):
            print(f"  {n:4d}  {reason}")
        if "--only" in sys.argv or "--src" in sys.argv: return
        print("script | tokenize | tokens -> AST -> bytecode | tokens -> bytecode (direct) | speed-up")
        for name, src in bench_scripts():
            tok, ast, direct = (float(x.replace(",", ".")) for x in timing.invoke((src, 5), (), ())[0].split("|"))
            print(f"{name} | {tok:.0f} | {ast:.0f} | {direct:.0f} | {ast / direct:.1f}x")
    finally:
        try:
            ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx).terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    main()
