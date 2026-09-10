#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fill_logic.py
=============
The document-filling engine, shared by both the original command-line
script (fill_handover.py) and this web app. It is copied verbatim from the
already-tested fill_handover.py so the exact same behaviour (and the two
RTL/formatting bugs it fixes) carries over unchanged.

It edits only the specific data runs inside Template.docx's XML (never the
fixed Arabic/English labels around them), so the original formatting,
fonts, RTL layout and table structure are preserved exactly - only the
data changes. It also auto-detects whether each value is Arabic or Latin
script and fixes that run's right-to-left flag to match.
"""

import re

import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

WEEKDAYS_AR = ["الاثنين", "الثلاثاء", "الاربعاء", "الخميس", "الجمعة", "السبت", "الاحد"]
MONTHS_AR = [
    "يناير", "فبراير", "مارس", "ابريل", "مايو", "يونيو",
    "يوليو", "اغسطس", "سبتمبر", "اكتوبر", "نوفمبر", "ديسمبر",
]

ARABIC_RE = re.compile(r"[؀-ۿ]")

REQUIRED_FIELDS = [
    "name", "department", "role", "mobile", "email",
    "code", "govid", "serial", "model", "cpu",
]


def has_arabic(text):
    return bool(ARABIC_RE.search(text))


def set_value(run, text):
    """Set a run's text, and fix its right-to-left flag to match the
    script of the new text (Arabic vs Latin/digits)."""
    run.text = text
    rPr = run._r.find(qn("w:rPr"))
    if rPr is None:
        return
    rtl_el = rPr.find(qn("w:rtl"))
    if has_arabic(text):
        if rtl_el is None:
            rPr.append(OxmlElement("w:rtl"))
    else:
        if rtl_el is not None:
            rPr.remove(rtl_el)


def clear(run):
    run.text = ""


def strip_hyperlinks(paragraph):
    p = paragraph._p
    for hl in p.findall(qn("w:hyperlink")):
        p.remove(hl)


def runs_of(table, row, col, para=0):
    return table.rows[row].cells[col].paragraphs[para].runs


def fill_document(template_path, output_path, data):
    """data must contain: name, department, role, mobile, email, code,
    govid, serial, model, cpu, company, color, storage, ram, date_obj
    (a datetime)."""
    d = docx.Document(str(template_path))

    date_obj = data["date_obj"]
    date_str = f"{date_obj.day}/{date_obj.month}/{date_obj.year}"
    day_name = WEEKDAYS_AR[date_obj.weekday()]
    month_name = MONTHS_AR[date_obj.month - 1]

    name = data["name"]
    department = data["department"]
    role = data["role"]
    mobile = data["mobile"]
    email = data["email"]
    code = data["code"]
    govid = data["govid"]
    serial = data["serial"]
    model = data["model"]
    cpu = data["cpu"]
    company = data["company"]
    color = data["color"]
    storage = data["storage"]
    ram = data["ram"]

    t1, t2, t3, t4, t5 = d.tables[1], d.tables[2], d.tables[3], d.tables[4], d.tables[5]

    # --- Table 1: date / month / weekday line -------------------------
    r = runs_of(t1, 0, 0)
    set_value(r[0], date_str)
    clear(r[1]); clear(r[2]); clear(r[3])

    r = runs_of(t1, 0, 1)
    set_value(r[3], month_name)

    r = runs_of(t1, 0, 2)
    set_value(r[1], day_name)
    clear(r[2])

    # --- Table 2: employee details -------------------------------------
    r = runs_of(t2, 0, 0)
    set_value(r[0], role + " ")

    r = runs_of(t2, 0, 1)
    set_value(r[2], name)
    clear(r[3])

    r = runs_of(t2, 1, 0)
    set_value(r[2], " " + company)

    r = runs_of(t2, 1, 1)
    set_value(r[2], department + " ")

    p = t2.rows[2].cells[0].paragraphs[0]
    r = p.runs
    set_value(r[0], mobile)
    clear(r[1])
    set_value(r[4], email)
    strip_hyperlinks(p)

    r = runs_of(t2, 2, 1)
    set_value(r[0], code)
    clear(r[1])

    # --- Table 3: laptop specs ------------------------------------------
    set_value(runs_of(t3, 1, 0)[0], model)
    set_value(runs_of(t3, 1, 1)[0], serial)
    set_value(runs_of(t3, 1, 2)[0], color)
    set_value(runs_of(t3, 1, 3)[0], storage)
    set_value(runs_of(t3, 1, 4)[0], ram)
    r = runs_of(t3, 1, 5)
    set_value(r[0], cpu)
    if len(r) > 1:
        clear(r[1])

    # --- Table 4: national ID / acknowledgement block --------------------
    r = runs_of(t4, 0, 0)
    set_value(r[3], govid)
    clear(r[4])
    set_value(r[13], company)
    clear(r[14])

    r = runs_of(t4, 0, 1)
    set_value(r[2], name)
    clear(r[3])
    set_value(r[26], department + " ")

    # --- Table 5: signature dates -----------------------------------------
    r = runs_of(t5, 2, 0)
    set_value(r[0], date_str)
    clear(r[1]); clear(r[2]); clear(r[3])

    r = runs_of(t5, 2, 1)
    set_value(r[0], date_str)
    clear(r[1]); clear(r[2]); clear(r[3])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    d.save(str(output_path))
    return output_path
