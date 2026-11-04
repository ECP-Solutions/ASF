"""Runs every example in LibreOffice Calc, Writer and Impress, headless, and builds the documents.

    python3 examples/portable/tools/run_lo.py            # compare with expected/
    python3 examples/portable/tools/run_lo.py --update   # write expected/ from this run

Builds the ASF library from src/ (see libreoffice/tools/build.py) and runs each script of
scripts/ through hosts/LibreOfficeHost.bas on three documents: a spreadsheet holding the tables
of data/, a new text document and a new presentation. What the script wrote is turned into text
(sheets as tab-separated rows, paragraphs with their style, tables as | rows, slides with their
title and body) and compared with expected/<script>.<app>.txt. It also saves the documents users
open: portable-examples.ods (the tables and the host), .odt and .odp (empty, with the host), and
portable-examples.xlsx (the tables, for Excel). The exit status is 0 when everything matches.
Needs LibreOffice and python3-uno.
"""
import csv, datetime, os, re, shutil, socket, subprocess, sys, tempfile, time

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)
ROOT = os.path.dirname(os.path.dirname(BASE))
sys.path.insert(0, os.path.join(ROOT, "libreoffice", "tools"))
import build  # noqa: E402

SCRIPTS = ["hello.vas", "budget.vas", "progress.vas", "price-revision.vas", "workreports.vas", "cleanup.vas"]
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


def fill_sheets(doc):
    from com.sun.star.lang import Locale
    date_format = doc.getNumberFormats().getStandardFormat(2, Locale())  # NumberFormat.DATE
    sheets = doc.getSheets()
    sheets.getByIndex(0).setName(SHEETS[0])
    for name in SHEETS[1:]:
        sheets.insertNewByName(name, sheets.getCount())
    for name in SHEETS:
        sheet = sheets.getByName(name)
        with open(os.path.join(BASE, "data", name + ".csv"), encoding="utf-8", newline="") as f:
            rows = list(csv.reader(f))
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


def cell(v):
    if isinstance(v, float):
        return ("%.10g" % v)
    return str(v)


def dump_sheet(doc, title):
    sheets = doc.getSheets()
    if not sheets.hasByName(title):
        return "(no sheet %s)\n" % title
    sheet = sheets.getByName(title)
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    end = cursor.getRangeAddress()
    rows = sheet.getCellRangeByPosition(0, 0, end.EndColumn, end.EndRow).getDataArray()
    return "".join("\t".join(cell(v) for v in r).rstrip("\t") + "\n" for r in rows)


def dump_text(doc):
    out = []
    enum = doc.getText().createEnumeration()
    while enum.hasMoreElements():
        el = enum.nextElement()
        if el.supportsService("com.sun.star.text.Paragraph"):
            if el.getString() != "":
                out.append("[%s] %s" % (el.ParaStyleName, el.getString()))
        elif el.supportsService("com.sun.star.text.TextTable"):
            for r in el.getDataArray():
                out.append("| " + " | ".join(cell(v) for v in r) + " |")
    return "".join(line + "\n" for line in out)


def dump_slides(doc):
    out = []
    pages = doc.getDrawPages()
    for i in range(pages.getCount()):
        page = pages.getByIndex(i)
        texts = [page.getByIndex(k).getString() for k in range(page.getCount())]
        out.append("== slide %d: %s" % (i + 1, texts[0] if texts else ""))
        for t in texts[1:]:
            for line in t.split("\n"):
                out.append("  " + line)
    return "".join(line + "\n" for line in out)


TITLES = {"hello.vas": "Hello from ASF", "budget.vas": "Budget summary", "progress.vas": "Progress certificate",
          "price-revision.vas": "Price revision", "workreports.vas": "Work report checks", "cleanup.vas": "Item list cleanup"}


def main(argv):
    update = "--update" in argv
    ctx, proc, work = start_office()
    import uno
    failures, docs = 0, []
    try:
        smgr = ctx.ServiceManager
        libs = smgr.createInstanceWithContext("com.sun.star.script.ApplicationScriptLibraryContainer", ctx)
        lib = libs.createLibrary(build.LIB_NAME)
        for name, code, _ in build.library_modules(ROOT):
            lib.insertByName(name, code)
        with open(os.path.join(BASE, "hosts", "LibreOfficeHost.bas"), encoding="utf-8") as f:
            host = f.read()
        hostlib = libs.createLibrary("PortableHost")
        hostlib.insertByName("LibreOfficeHost", host)
        libs.loadLibrary("PortableHost")
        provider = smgr.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", ctx).createScriptProvider("")
        run = provider.getScript("vnd.sun.star.script:PortableHost.LibreOfficeHost.RunScriptOn?language=Basic&location=application")
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        new = lambda kind: desktop.loadComponentFromURL(f"private:factory/{kind}", "_blank", 0, (prop("Hidden", True),))
        folder = os.path.join(BASE, "scripts")

        # the documents users open
        calc = new("scalc")
        fill_sheets(calc)
        calc.storeToURL(uno.systemPathToFileUrl(os.path.join(BASE, "portable-examples.xlsx")), (prop("FilterName", "Calc MS Excel 2007 XML"),))
        for kind, ext in (("scalc", "ods"), ("swriter", "odt"), ("simpress", "odp")):
            d = calc if kind == "scalc" else new(kind)
            basic = d.BasicLibraries
            std = basic.getByName("Standard") if basic.hasByName("Standard") else basic.createLibrary("Standard")
            std.insertByName("LibreOfficeHost", host)
            d.storeToURL(uno.systemPathToFileUrl(os.path.join(BASE, f"portable-examples.{ext}")), ())
            if d is not calc:
                d.close(True)

        writer, impress = new("swriter"), new("simpress")
        docs = [calc, writer, impress]
        for script in SCRIPTS:
            for app, d in (("calc", calc), ("writer", writer), ("impress", impress)):
                # a fresh document for every script outside the spreadsheet
                if app != "calc":
                    d.close(True)
                    d = new("swriter" if app == "writer" else "simpress")
                    docs.append(d)
                msg = run.invoke((d, folder, script), (), ())[0]
                text = "result: %s\n" % msg
                if app == "calc":
                    text += dump_sheet(d, TITLES[script])
                elif app == "writer":
                    text += dump_text(d)
                else:
                    text += dump_slides(d)
                path = os.path.join(BASE, "expected", "%s.%s.txt" % (script[:-4], app))
                if update:
                    os.makedirs(os.path.dirname(path), exist_ok=True)
                    with open(path, "w", encoding="utf-8", newline="\n") as f:
                        f.write(text)
                    status = "written"
                else:
                    with open(path, encoding="utf-8") as f:
                        expected = f.read()
                    status = "ok" if expected == text else "DIFFERS"
                    if expected != text:
                        failures += 1
                        exp, got = expected.splitlines(), text.splitlines()
                        for k in range(max(len(exp), len(got))):
                            e = exp[k] if k < len(exp) else ""
                            g = got[k] if k < len(got) else ""
                            if e != g:
                                print(f"    line {k + 1}: expected {e!r}\n             got      {g!r}")
                                break
                print(f"{script:20s} {app:8s} {status}  ({msg})", flush=True)
                if app == "writer":
                    writer = d
                elif app == "impress":
                    impress = d
    finally:
        for d in docs:
            try:
                d.close(True)
            except Exception:
                pass
        try:
            ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx).terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()
        shutil.rmtree(work, ignore_errors=True)
    if not update:
        print("all examples match" if not failures else f"{failures} run(s) differ")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main(sys.argv[1:])
