# -*- coding: utf-8 -*-
"""
asset_register.py
=================
The laptop register: one spreadsheet saying who has which machine, what
it is, who gave it to them, and whether they still have it.

It is DERIVED, never edited. Every time a document is generated, edited
or deleted, and every time someone is marked as having left, the whole
sheet is rebuilt from the database. That is slower than patching one row
and worth it: a register that is rebuilt cannot drift from the records it
describes, and there is no state to repair when something goes wrong
halfway. It also means the very first rebuild backfills every document
ever generated, with no import step.

The trade is that hand-edits to the .xlsx are overwritten on the next
document. If a row is wrong, the record behind it is wrong - fix it in
the app and the sheet follows.

One row per assignment, so the history is kept: replacing a laptop marks
the old row Replaced and adds a new one rather than overwriting. Filter
Status = Held for the current picture.
"""

from datetime import datetime

# Only the documents that issue a computer. A headset receipt has a
# serial number too and it has no business in a laptop register.
LAPTOP_TEMPLATES = {
    # template id -> where that document's fields keep the machine it
    # ISSUES. A replacement hands over the "new_" one; the machine it
    # takes back is the row already in the register.
    "laptop_handover": "",
    "laptop_replacement": "new_",
}

COLUMNS = (
    ("Employee", 26), ("Employee code", 14), ("Department", 18), ("Email", 26),
    ("Computer name", 16), ("Model", 20), ("Serial number", 18),
    ("CPU", 10), ("RAM", 10), ("Storage", 14), ("Colour", 12),
    ("Document", 20), ("Assigned on", 13), ("Assigned by", 18),
    ("Status", 11), ("Ended on", 13), ("Note", 30),
)

HELD, REPLACED, LEFT = "Held", "Replaced", "Left"

# Held is the one worth spotting from across the room; the other two are
# history and should stay quiet.
STATUS_FILL = {
    HELD: "E9F4EE",
    REPLACED: "F1F3F9",
    LEFT: "FDECEB",
}
STATUS_TEXT = {
    HELD: "1F6B4A",
    REPLACED: "5F6577",
    LEFT: "9C2B1F",
}


def parse_date(value):
    """The "D/M/YYYY" the records store, as a date. Anything unparseable
    sorts last rather than raising - a register that refuses to build
    because one row has a typo in it is worse than one odd row."""
    text = (value or "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def show_date(value):
    """1-Sep-26, matching how the emails write dates."""
    parsed = parse_date(value)
    if parsed is None:
        return (value or "").strip()
    return f"{parsed.day}-{parsed.strftime('%b')}-{parsed.strftime('%y')}"


def assignment(record, identity):
    """One laptop handed to one person, or None if this document did not
    hand over a laptop."""
    if record.template_id not in LAPTOP_TEMPLATES:
        return None
    prefix = LAPTOP_TEMPLATES[record.template_id]
    fields = record.fields

    def field(key):
        return (fields.get(prefix + key) or "").strip()

    return {
        "identity": identity,
        "record_id": record.id,
        "name": record.name or (fields.get("name") or "").strip(),
        "code": (fields.get("code") or "").strip(),
        "department": (fields.get("department") or record.department or "").strip(),
        "email": (fields.get("email") or "").strip(),
        "computer_name": (fields.get("computer_name") or "").strip(),
        "model": field("model"),
        "serial": field("serial"),
        "cpu": field("cpu"),
        "ram": field("ram"),
        "storage": field("storage"),
        "color": field("color"),
        "document": record.template_label,
        "date": record.handover_date or "",
        "sort_date": parse_date(record.handover_date),
        "by": record.created_by or "",
        "status": HELD,
        "ended": "",
        "note": "",
    }


def build_rows(records, identity_of, departures):
    """Every laptop assignment, newest first, with a status worked out
    per person.

    `departures` maps an employee identity to a Departure-shaped object
    with `.left_on` and `.name`.
    """
    rows = []
    for record in records:
        found = assignment(record, identity_of(record))
        if found and (found["serial"] or found["model"]):
            rows.append(found)

    # Per person, oldest first: everything but their latest machine has
    # been superseded, and the replacement's own date is the day the old
    # one came back.
    by_person = {}
    for row in rows:
        by_person.setdefault(row["identity"], []).append(row)

    for identity, mine in by_person.items():
        mine.sort(key=lambda r: (r["sort_date"] or datetime.min.date(), r["record_id"]))
        for older, newer in zip(mine, mine[1:]):
            older["status"] = REPLACED
            older["ended"] = newer["date"]
            older["note"] = f"Replaced by {newer['model'] or 'a new machine'}".strip()
        gone = departures.get(identity)
        if gone is not None:
            # Leaving trumps holding: the machine is not theirs any more
            # whatever the last document said. Rows already marked
            # Replaced stay that way - they ended before the person did.
            current = mine[-1]
            current["status"] = LEFT
            current["ended"] = gone.left_on or ""
            current["note"] = "Employee left the company"

    rows.sort(key=lambda r: (r["sort_date"] or datetime.min.date(), r["record_id"]),
              reverse=True)
    return rows


def row_values(row):
    return [row["name"], row["code"], row["department"], row["email"],
            row["computer_name"], row["model"], row["serial"], row["cpu"],
            row["ram"], row["storage"], row["color"], row["document"],
            show_date(row["date"]), row["by"], row["status"],
            show_date(row["ended"]), row["note"]]


def write_workbook(rows, path):
    """The sheet itself. Written to a neighbouring temporary file and
    moved into place, so a reader never catches it half-written."""
    import os
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Laptops"

    ws.append([heading for heading, _ in COLUMNS])
    head_fill = PatternFill("solid", fgColor="E9ECF3")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = head_fill
        cell.alignment = Alignment(vertical="center")

    for row in rows:
        ws.append(row_values(row))
        status = ws.cell(row=ws.max_row, column=15)
        status.font = Font(bold=True, color=STATUS_TEXT.get(row["status"], "000000"))
        status.fill = PatternFill("solid", fgColor=STATUS_FILL.get(row["status"], "FFFFFF"))

    for i, (_, width) in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # Filter buttons on the header row and the header frozen, because the
    # first thing anyone does with this is filter Status = Held.
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{max(ws.max_row, 1)}"

    path = str(path)
    tmp = f"{path}.tmp"
    wb.save(tmp)
    os.replace(tmp, path)
    return len(rows)
