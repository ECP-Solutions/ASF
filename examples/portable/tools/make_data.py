"""Writes scripts/data/project.vas, the sample project as an ASF module, from data/*.csv.

    python3 examples/portable/tools/make_data.py

The scripts read their tables from the spreadsheet when it has them (input(name, fallback) of
lib/office.vas) and from this module otherwise, so they also run in Word, Writer, PowerPoint and
Impress, or on an empty workbook. Dates become serial numbers, as both spreadsheets hold them.
"""
import csv, datetime, os, re

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(HERE)
TABLES = [("budget", "Budget"), ("workReports", "WorkReports"), ("progress", "Progress"),
          ("priceIndices", "PriceIndices"), ("supplierItems", "Cleanup")]


def literal(v):
    if re.fullmatch(r"-?\d+(\.\d+)?", v):
        return v
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
        return str((datetime.date.fromisoformat(v) - datetime.date(1899, 12, 30)).days)
    # other characters by their code: ASF reads modules byte by byte, so they stay plain ASCII
    parts = re.findall(r"[\x00-\x7f]+|[^\x00-\x7f]", v)
    return " + ".join(["''.fromCharCode(%d)" % ord(p) if ord(p[0]) > 127 else "'" + p.replace("\\", "\\\\").replace("'", "\\'") + "'" for p in parts]) or "''"


def main():
    out = ["// Sample data of a building project, generated from data/*.csv by tools/make_data.py.",
           "// The scripts use it when the document has no sheet with the table: dates are serial numbers.", ""]
    for name, sheet in TABLES:
        with open(os.path.join(BASE, "data", sheet + ".csv"), encoding="utf-8", newline="") as f:
            rows = list(csv.reader(f))
        out.append(f"// sheet {sheet}")
        out.append(f"let {name} = [")
        lines = []
        for i, r in enumerate(rows):
            cells = ["'" + v + "'" if i == 0 else literal(v) for v in r]
            lines.append("    [" + ", ".join(cells) + "]")
        out.append(",\n".join(lines))
        out.append("];")
        out.append("")
    out.append("export { " + ", ".join(n for n, _ in TABLES) + " };")
    path = os.path.join(BASE, "scripts", "data", "project.vas")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="ascii", newline="\n") as f:
        f.write("\n".join(out) + "\n")
    print(path)


if __name__ == "__main__":
    main()
