"""The sample data of the pipeline workbook: orders as they come from several sources, a customer
reference table and the validation rules. Shared by make_workbook.py and check.py."""

# Order | Date | Customer | Email | Product | Qty | Price | Channel, with the faults real exports have:
# numbers and dates as text in several formats, stray spaces and case, a blank row, a duplicate
# corrected later, an unknown customer, a bad email, an unknown product, a negative quantity...
RAW = [
    ["Order", "Date", "Customer", "Email", "Product", "Qty", "Price", "Channel"],
    ["SO-1001", 46023, " c001 ", "ANA@Example.com ", "widget", 3, "$12.50", "Web"],
    ["SO-1002", "2026-01-03", "C002", "bo@example.org", " Gadget ", "2", "1,250.00", "store"],
    ["SO-1003", "05/01/2026", "c003", "cy@example.com", "GIZMO", 1, "€ 9,99", "web"],
    ["SO-1004", "2026-01-06", "C004", "dee.at.example.com", "Widget", 5, 12.5, "Phone"],
    ["SO-1005", None, "C005", "ed@example.com", "Gadget", 2, "1250", "Store"],
    ["SO-1006", "31/02/2026", "C006", "fi@example.com", "Gizmo", 4, "9.99", "Web"],
    ["SO-1007", "2026-01-10", "C099", "gu@example.com", "Widget", 1, "12.50", "Web"],
    ["SO-1008", "2026-01-11", "C007", None, "Doohickey", 2, "5", "Store"],
    ["SO-1009", "2026-01-12", "C008", "hi@example.com", "Widget", -1, "12.5", "Web"],
    ["SO-1010", "2026-01-12", "C001", "ana@example.com", "Gadget", "three", "1250", "Phone"],
    ["SO-1002", "2026-01-03", "C002", "bo@example.org", "Gadget", 3, "1,250.00", "Store"],
    [None, None, None, None, None, None, None, None],
    ["so-1011", "2026-01-15", "C009", "ivy@example.com", "Gizmo", 6, "9.99", "Web"],
    ["SO-1012", "2026-01-18", "C010", "jo@example.net", "Widget", 2, "(3.00)", "Store"],
    ["SO-1013", "20/01/2026", "C011", "kai@example.com", "Gadget", 1, "1.250,00", "Phone"],
    ["SO-1014", "2026-02-02", "C012", "lu@example.com", "Widget", 10, "12.5", "Web"],
    ["SO-1015", "2026-02-03", "C001", "ana@example.com", "Gizmo", 2, "9.99", "Web"],
    ["SO-1016", "04/02/2026", "C002", "bo@example.org", "Widget", 4, "12.50", "Store"],
    ["SO-1017", "2026-02-09", "C003", "cy@example.com", "Gadget", 1, "1250", "Phone"],
    ["SO-1018", "2026-02-14", "C004", "dee@example.com", "Gizmo", 3, "9.99", "Web"],
    ["SO-1019", "2026-02-20", "C005", "ed@example.com", "Widget", 7, "12.5", "Store"],
    ["SO-1020", "2026-02-25", "C009", "ivy@example.com", "Gadget", 2, "1,250", "Web"],
    ["SO-1021", "2026-03-02", "C010", "jo@example.net", "Widget", 1, "12.5", "Phone"],
    ["SO-1022", "05/03/2026", "C011", "kai@example.com", "Gizmo", 5, "9.99", "Web"],
    ["SO-1023", "2026-03-09", "C012", "lu@example.com", "Gadget", 1, "1250.00", "Store"],
    ["SO-1024", "2026-03-15", " c006 ", "fi@example.com", "widget", 8, " 12.50 ", "web"],
]

CUSTOMERS = [
    ["CustomerID", "Name", "Region", "Segment"],
    ["C001", "Ana Lopez", "North", "Retail"],
    ["C002", "Bo Chen", "South", "Wholesale"],
    ["C003", "Cy Ortiz", "East", "Retail"],
    ["C004", "Dee Moss", "West", "Retail"],
    ["C005", "Ed Park", "North", "Wholesale"],
    ["C006", "Fi Grant", "South", "Retail"],
    ["C007", "Gus Hale", "East", "Wholesale"],
    ["C008", "Hi Tran", "West", "Retail"],
    ["C009", "Ivy Kerr", "North", "Retail"],
    ["C010", "Jo Reyes", "South", "Wholesale"],
    ["C011", "Kai Dunn", "East", "Retail"],
    ["C012", "Lu Wei", "West", "Wholesale"],
]

RULES = [
    ["Column", "Rule", "Value"],
    ["Order", "required", None],
    ["Order", "pattern", r"^SO-\d{4}$"],
    ["Order", "unique", None],
    ["Date", "required", None],
    ["Date", "date", None],
    ["Customer", "required", None],
    ["Email", "pattern", r"^[^@\s]+@[^@\s]+\.[a-z]{2,}$"],
    ["Product", "oneOf", "Widget, Gadget, Gizmo"],
    ["Qty", "required", None],
    ["Qty", "number", None],
    ["Qty", "min", 1],
    ["Price", "number", None],
    ["Price", "min", 0],
]

CLEAN_SPEC = "Order:upper, Date:date-dmy, Customer:upper, Email:lower, Product:title, Qty:number, Price:number, Channel:title"

# the script in a cell of the workbook (Scripts!B2) that adds Amount = Qty * Price
AMOUNT_SCRIPT = (
    "import { trimTable, columnIndex, column } from './tables.vas';\n"
    "let t = trimTable($1);\n"
    "let i = columnIndex(t);\n"
    "let q = column(i, 'Qty');\n"
    "let p = column(i, 'Price');\n"
    "t[1].push('Amount');\n"
    "for (let r = 2, r <= t.length, r += 1) { t[r].push(Math.round(t[r][q] * t[r][p] * 100) / 100); };\n"
    "return t;"
)
