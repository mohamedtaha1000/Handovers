#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
employees.py
============
Recognising a person across their documents, and working out what they
still hold.

There is no employee table in this app - only documents, each carrying
its own copy of the person's details. So "the same person" is a
judgement (employee_identity), "have I made this one before" is a
judgement (find_duplicate), and "which laptop do they have" means
reading every document they appear on. Those judgements live here rather
than among the routes: they are the rules most likely to need changing,
and the ones worth being able to read on their own.

No Flask. Given the rows, every function here is a pure question about
them.
"""

import asset_register
from models import Departure, Handover
from templates import EMPLOYEE_KEYS, TEMPLATES

# The fields worth showing in the history table's equipment column, in
# the order we would rather have them: what the thing is, then which one.
EQUIPMENT_KEYS = ("model", "new_model", "brand", "capacity", "old_model")
SERIAL_KEYS = ("serial", "new_serial", "sim_number", "old_serial")


# Where a computer's serial number lives, per document type. A
# replacement issues a new machine, so its NEW serial is the one the
# person is still holding on the day they leave. A headset receipt has a
# serial too and it is not the one the security team is asking about, so
# only these two count.
COMPUTER_SERIAL_KEYS = {
    "laptop_handover": "serial",
    "laptop_replacement": "new_serial",
}



def employee_identity(record):
    """What makes two records the same person. The employee code is the
    real identifier; the national ID backs it up for older records that
    predate the code field, and the name is the last resort."""
    fields = record.fields
    code = (fields.get("code") or "").strip().lower()
    govid = (record.govid or "").strip()
    return code or govid or (record.name or "").strip().lower()


def known_employees(limit=400):
    """One entry per person, taken from their most recent document.

    Most recent matters: someone who changed department should come back
    with the department they are in now, not the one they were in when
    they were first issued a laptop."""
    records = (Handover.query
               .order_by(Handover.created_at.desc(), Handover.id.desc())
               .limit(limit).all())
    people = {}
    for record in records:
        key = employee_identity(record)
        if not key or key in people:
            continue
        fields = record.fields
        entry = {k: (fields.get(k) or "") for k in EMPLOYEE_KEYS}
        entry["name"] = entry["name"] or record.name or ""
        entry["department"] = entry["department"] or record.department or ""
        entry["role"] = entry["role"] or record.role or ""
        entry["govid"] = entry["govid"] or record.govid or ""
        if entry["name"]:
            people[key] = entry
    return list(people.values())


def find_duplicate(template_id, values, exclude_id=None):
    """An earlier document of this same type for this same person, if
    there is one. Two handovers of the same thing to the same person is
    usually a mistake - or a sign the replacement template was the one
    actually wanted - so it is worth asking before generating another."""
    code = (values.get("code") or "").strip()
    govid = (values.get("govid") or "").strip()
    name = (values.get("name") or "").strip()

    query = Handover.query.filter(Handover.template_id == template_id)
    if exclude_id is not None:
        query = query.filter(Handover.id != exclude_id)
    for record in query.order_by(Handover.created_at.desc()).limit(200).all():
        fields = record.fields
        if code and (fields.get("code") or "").strip() == code:
            return record
        if govid and (record.govid or "").strip() == govid:
            return record
        if not code and not govid and name and (record.name or "").strip() == name:
            return record
    return None


def equipment_summary(record):
    """A one-line "what was handed over" for a history row: the model (or
    brand, or capacity) and the serial, drawn from whichever fields that
    document type happens to collect."""
    fields = record.fields
    def first(keys):
        for key in keys:
            value = (fields.get(key) or "").strip()
            if value:
                return value
        return ""
    what = first(EQUIPMENT_KEYS)
    brand = (fields.get("brand") or "").strip()
    if brand and what and brand != what:
        what = f"{brand} {what}"
    return {"what": what, "serial": first(SERIAL_KEYS)}


def leaver_lookup(limit=400):
    """Everyone with a document on file, shaped like the leaver form.

    Feeds the "reuse someone already on file" suggestions on that page.
    Nobody has to be here - the form is typed either way - but when the
    person does have documents this saves copying five things across.

    Employee details come from their most recent document, so a change of
    department is reflected. The serial comes from the most recent
    document that actually issued them a computer, and the computer name
    from the most recent document that recorded one.
    """
    records = (Handover.query
               .order_by(Handover.created_at.desc(), Handover.id.desc())
               .limit(limit).all())
    grouped = {}
    for record in records:
        key = employee_identity(record)
        if key:
            grouped.setdefault(key, []).append(record)

    people = []
    for mine in grouped.values():
        newest = mine[0]
        fields = newest.fields
        serial = next(
            ((r.fields.get(COMPUTER_SERIAL_KEYS[r.template_id]) or "").strip()
             for r in mine
             if r.template_id in COMPUTER_SERIAL_KEYS
             and (r.fields.get(COMPUTER_SERIAL_KEYS[r.template_id]) or "").strip()), "")
        computer = next(((r.fields.get("computer_name") or "").strip()
                         for r in mine if (r.fields.get("computer_name") or "").strip()), "")
        name = newest.name or (fields.get("name") or "").strip()
        if not name:
            continue
        people.append({
            "name": name,
            "department": (fields.get("department") or newest.department or "").strip(),
            "email": (fields.get("email") or "").strip(),
            "computer_name": computer,
            "serial": serial,
            "code": (fields.get("code") or "").strip(),
        })
    return people


def register_rows():
    """Every laptop assignment the register knows about, built from the
    documents and the people who have left. Lives here rather than in
    app.py because it is the same question matching_laptops asks."""
    records = Handover.query.all()
    departures = {d.identity: d for d in Departure.query.all()}
    return asset_register.build_rows(records, employee_identity, departures)


def matching_laptops(name="", serial="", code=""):
    """The laptop rows the register holds for whoever is typed into the
    leaver page.

    That page is free text - the person may never have had a document
    generated - so this matches on what it has: the employee code first
    because it is the real identifier, then an exact serial, then the
    name. Nothing fuzzy: marking the wrong person as gone is worse than
    finding nobody and saying so.
    """
    name, serial, code = name.strip().lower(), serial.strip().lower(), code.strip().lower()
    if not (name or serial or code):
        return []
    rows = register_rows()
    matched = []
    for row in rows:
        if code and row["code"].strip().lower() == code:
            matched.append(row)
        elif serial and row["serial"].strip().lower() == serial:
            matched.append(row)
        elif name and row["name"].strip().lower() == name:
            matched.append(row)
    return matched
