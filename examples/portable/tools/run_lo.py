"""Builds the example workbooks and runs the examples in LibreOffice, headless.

    python3 examples/portable/tools/run_lo.py

Builds the ASF library from src/ (see libreoffice/tools/build.py), creates a Calc document with
one sheet per table of data/, saves it as portable-examples.ods with hosts/CalcHost.bas inside and
as portable-examples.xlsx (tables only, for Excel), runs CalcHost's RunAll on the .ods and compares
every result sheet with expected/ (see expected.py). The exit status is 0 when all of them match.
Needs LibreOffice and python3-uno.
"""
import csv, datetime, os, re, shutil, socket, subprocess, sys, tempfile, time

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)
ROOT = os.path.dirname(os.path.dirname(BASE))
sys.path.insert(0, os.path.join(ROOT, "libreoffice", "tools"))
import build  # noqa: E402

SHEETS = ["Budget", "WorkReports", "Progress", "PriceIndices", "Cleanup"]
DATE = re.compile(r"\d{4}-\d{2}-\d{2}")
NUMBER = re.compile(r"-?\d+(\.\d+)?")


def start_office():
    import uno
    from com.sun.star.connection import NoConnectException
    work = tempfile.mkdtemp(prefix="asf-portable-")
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


def prop(name, value):
    from com.sun.star.beans import PropertyValue
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def read_csv(path):
    with open(path, encoding="utf-8", newline="") as f:
        return list(csv.reader(f))


def fill_sheet(doc, sheet, rows):
    from com.sun.star.lang import Locale
    date_format = doc.getNumberFormats().getStandardFormat(2, Locale())  # NumberFormat.DATE
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = sheet.getCellByPosition(j, i)
            if i > 0 and DATE.fullmatch(v):
                c.setValue((datetime.date.fromisoformat(v) - datetime.date(1899, 12, 30)).days)
                c.NumberFormat = date_format
            elif i > 0 and NUMBER.fullmatch(v):
                c.setValue(float(v))
            elif v != "":
                c.setString(v)
    sheet.getColumns().OptimalWidth = True


def used_rows(sheet):
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    end = cursor.getRangeAddress()
    return sheet.getCellRangeByPosition(0, 0, end.EndColumn, end.EndRow).getDataArray()


def same(expected, actual):
    if NUMBER.fullmatch(expected) or re.fullmatch(r"-?\d+(\.\d+)?e[-+]?\d+", expected):
        return isinstance(actual, float) and abs(float(expected) - actual) <= 1e-9 * max(1.0, abs(actual))
    return expected == ("" if actual is None else str(actual))


def main():
    ctx, proc, work = start_office()
    failures = 0
    doc = None
    try:
        smgr = ctx.ServiceManager
        libs = smgr.createInstanceWithContext("com.sun.star.script.ApplicationScriptLibraryContainer", ctx)
        lib = libs.createLibrary(build.LIB_NAME)
        for name, code, _ in build.library_modules(ROOT):
            lib.insertByName(name, code)
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (prop("Hidden", True),))
        sheets = doc.getSheets()
        sheets.getByIndex(0).setName(SHEETS[0])
        for name in SHEETS[1:]:
            sheets.insertNewByName(name, sheets.getCount())
        for name in SHEETS:
            fill_sheet(doc, sheets.getByName(name), read_csv(os.path.join(BASE, "data", name + ".csv")))
        import uno
        doc.storeToURL(uno.systemPathToFileUrl(os.path.join(BASE, "portable-examples.xlsx")),
                       (prop("FilterName", "Calc MS Excel 2007 XML"),))
        with open(os.path.join(BASE, "hosts", "CalcHost.bas"), encoding="utf-8") as f:
            basic = doc.BasicLibraries
            std = basic.getByName("Standard") if basic.hasByName("Standard") else basic.createLibrary("Standard")
            std.insertByName("CalcHost", f.read())
        doc.storeAsURL(uno.systemPathToFileUrl(os.path.join(BASE, "portable-examples.ods")), ())
        doc.getScriptProvider().getScript(
            "vnd.sun.star.script:Standard.CalcHost.RunAll?language=Basic&location=document").invoke((), (), ())
        for name in SHEETS:
            out = name + " result"
            expected = read_csv(os.path.join(BASE, "expected", out + ".csv"))
            actual = used_rows(sheets.getByName(out)) if sheets.hasByName(out) else ()
            bad = []
            for i in range(max(len(expected), len(actual))):
                e = expected[i] if i < len(expected) else []
                a = list(actual[i]) if i < len(actual) else []
                width = max(len(e), len(a))
                e = e + [""] * (width - len(e))
                a = a + [""] * (width - len(a))
                if not all(same(x, y) for x, y in zip(e, a)):
                    bad.append(f"  row {i + 1}: expected {e}, got {a}")
            print(f"{name}: {'ok' if not bad else 'DIFFERS'} ({len(actual)} rows)")
            for b in bad[:10]:
                print(b)
            failures += bool(bad)
    finally:
        try:
            if doc is not None:
                doc.close(True)
            smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx).terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()
        shutil.rmtree(work, ignore_errors=True)
    print("all examples match" if not failures else f"{failures} example(s) differ")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
