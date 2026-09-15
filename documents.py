#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
documents.py
============
Turning a submitted form into a stored Word document: validating what
was typed, naming the file, filling it, and writing the result onto a
Handover row.

No Flask. Everything here takes plain values and returns plain values or
a sentence describing what went wrong, which is what lets the same code
serve one document, an edit, and a batch of ten - and lets these rules
be exercised without starting a web server.
"""

import json
import logging
import os
import re
import secrets
from datetime import datetime

import builders
import settings
from models import Handover
from templates import EMPLOYEE_KEYS, TEMPLATES, all_fields, required_field_keys, template_path

log = logging.getLogger(__name__)

# The Arabic naming convention the team uses for the file a person
# downloads, kept separate from the ASCII-safe name on disk.
FILENAME_UNSAFE_RE = re.compile(r'[\\/:*?"<>|]')

# Asked once on the combined form and used by every document in it.
# The employee half, the date - and the computer name, because a person
# handed a laptop, a mouse and a headset on the same morning is sitting
# at one machine, and typing its name three times would only be a way to
# get it wrong twice.
SHARED_BATCH_KEYS = EMPLOYEE_KEYS + ("date", "computer_name")


def display_filename(template_id, name):
    """The clean, human-facing filename a person sees when they download -
    e.g. "استلام لابتوب(Name).docx" - no timestamps or ids in it."""
    spec = TEMPLATES.get(template_id, {})
    prefix = spec.get("filename_prefix", "مستند")
    safe_name = FILENAME_UNSAFE_RE.sub("", name).strip() or "employee"
    return f"{prefix}({safe_name}).docx"


def generated_filename(template_id, name, date_obj, exclude=None):
    """The filename a generated document is actually saved under in
    generated/ - e.g. "Yasmin Mohamed - Laptop Handover - 2026-09-10.docx" -
    so the folder is browsable on its own, not just through the site.
    Still collision-safe: if that exact name already exists (same person,
    same document type, same day), a " (2)", " (3)", ... counter is
    appended rather than overwriting an earlier document."""
    spec = TEMPLATES.get(template_id, {})
    label = spec.get("label", template_id)
    safe_name = FILENAME_UNSAFE_RE.sub("", name).strip() or "employee"
    safe_label = FILENAME_UNSAFE_RE.sub("", label).strip() or template_id
    date_str = date_obj.strftime("%Y-%m-%d")
    base = f"{safe_name} - {safe_label} - {date_str}"
    candidate = f"{base}.docx"
    n = 2
    # `exclude` is the record's own current file when re-generating after
    # an edit: without it, correcting a typo that doesn't change the name
    # or the date would see the record's existing file, decide the name
    # was taken, and save the correction as " (2)" beside the original.
    while (settings.GENERATED_DIR / candidate).exists() and candidate != exclude:
        candidate = f"{base} ({n}).docx"
        n += 1
    return candidate


def collect_values(template_id, form):
    """Read this template's fields out of a submitted form, validate them,
    and return (values, fill_data, date_obj, errors). `errors` empty means
    the submission is good."""
    spec = TEMPLATES[template_id]
    fields = all_fields(template_id)
    values = {f["key"]: form.get(f["key"], "").strip() for f in fields}
    errors = []

    # Named the way the form labels them, not by their internal keys:
    # this message is now the thing an older record shows when it is
    # opened for editing and has no computer name yet, so it has to read
    # like a sentence rather than like a database column.
    labels = {f["key"]: f["label"] for f in fields}
    missing = [labels.get(k, k) for k in required_field_keys(template_id) if not values[k]]
    if missing:
        errors.append("Please fill in all required fields: " + ", ".join(missing))

    # Format checks (mobile number, national ID, ...) - only fields whose
    # spec declares a "pattern" get checked; a value that's merely
    # non-empty but the wrong shape (too short, letters where there should
    # be digits, ...) is caught here rather than ending up wrong inside
    # the generated document.
    for f in fields:
        pattern = f.get("pattern")
        if not pattern or not values.get(f["key"]):
            continue
        if not re.fullmatch(pattern, values[f["key"]]):
            errors.append(f.get("pattern_msg") or f"{f['label']} is not the right format.")

    # Email field (only some templates have one): always force the
    # @stm.com.eg domain - only the part before an "@" (if the person
    # typed one) is kept, so it's impossible to end up with any other
    # domain.
    if "email" in values and values["email"]:
        values["email"] = values["email"].split("@")[0].strip() + "@stm.com.eg"

    date_input = form.get("date", "").strip()
    date_obj = datetime.today()
    if date_input:
        try:
            date_obj = datetime.strptime(date_input, "%Y-%m-%d")
        except ValueError:
            errors.append("Invalid date")

    fill_data = dict(values)
    fill_data["date_obj"] = date_obj
    for extra in spec.get("extra_fields", []):
        fill_data[extra["key"]] = form.get(extra["key"], "").strip() or extra.get("default", "")

    return values, fill_data, date_obj, errors


def write_document(template_id, fill_data, internal_name):
    """Generate the .docx into generated/ under `internal_name`.

    Written to a temporary file first and moved into place only once it
    has been produced in full, so a re-generation that fails part way
    through can't leave a half-written document where a good one was."""
    spec = TEMPLATES[template_id]
    doc_path = template_path(template_id)
    if not doc_path.exists():
        return f"{spec['doc_file']} is missing on the server."
    tmp_path = settings.GENERATED_DIR / f".tmp-{secrets.token_hex(8)}.docx"
    try:
        builders.fill_function(template_id)(doc_path, tmp_path, fill_data)
        os.replace(tmp_path, settings.GENERATED_DIR / internal_name)
    except Exception as failure:              # noqa: BLE001 - see below
        # A template that cannot be filled used to escape as a 500. It is
        # a real bug either way, so the traceback is logged in full - but
        # the person in front of it gets a sentence instead of a crash
        # page, and in a batch the documents that did work are kept and
        # named. Returning a problem is the whole contract save_document
        # is built on.
        tmp_path.unlink(missing_ok=True)
        log.exception("filling %s failed", template_id)
        return (f"{spec['label']} could not be generated from "
                f"{spec['doc_file']} ({failure}). Nothing was saved for it.")
    return None


def stamp_record(record, template_id, values, fill_data, date_obj, internal_name):
    """Write one generated document onto a Handover row.

    The same seven fields were being set in three places - creating one
    document, editing one, and creating a batch - which is why adding the
    register earlier meant four separate edits and adding the computer
    name meant touching each. One function, called three times.

    Takes the row rather than making it, because editing has to keep the
    existing id, created_by and created_at; only the caller knows which
    of the two it is doing.
    """
    record.template_id = template_id
    record.name = values["name"]
    record.department = values["department"]
    record.role = values["role"]
    record.govid = values.get("govid", "")
    record.handover_date = f"{date_obj.day}/{date_obj.month}/{date_obj.year}"
    record.fields_json = json.dumps(fill_data, default=str)
    record.filename = internal_name
    return record


def form_data_from_record(record):
    """Turn a saved record back into the dict form.html pre-fills from."""
    data = dict(record.fields)

    # The date is stored as a "YYYY-MM-DD HH:MM:SS" string (fields_json
    # serialises the datetime), but <input type="date"> needs the bare
    # date, so hand it one it will actually display.
    raw = str(data.pop("date_obj", "") or "")
    stamp = ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            stamp = datetime.strptime(raw[:19] if " " in raw else raw, fmt).strftime("%Y-%m-%d")
            break
        except ValueError:
            continue
    if not stamp and record.handover_date:
        try:
            day, month, year = record.handover_date.split("/")
            stamp = f"{int(year):04d}-{int(month):02d}-{int(day):02d}"
        except (ValueError, AttributeError):
            stamp = ""
    data["date"] = stamp or date.today().isoformat()

    # A field rendered with a suffix (the "@stm.com.eg" on the email box)
    # only ever shows the part before it, so strip the suffix back off -
    # otherwise editing would round-trip to "name@stm.com.eg@stm.com.eg".
    for f in all_fields(record.template_id):
        suffix = f.get("suffix")
        value = data.get(f["key"], "")
        if suffix and isinstance(value, str) and value.endswith(suffix):
            data[f["key"]] = value[: -len(suffix)]
    return data


def scoped_form(template_id, form):
    """A view of the submitted form as this one template expects it:
    shared fields as they are, per-document fields un-prefixed."""
    prefix = f"{template_id}__"
    scoped = {}
    for key in form.keys():
        if key.startswith(prefix):
            scoped[key[len(prefix):]] = form.get(key)
        elif key in SHARED_BATCH_KEYS:
            scoped[key] = form.get(key)
    return scoped


def batch_records(raw_ids):
    ids = [int(i) for i in raw_ids.split(",") if i.strip().isdigit()]
    found = {r.id: r for r in Handover.query.filter(Handover.id.in_(ids)).all()} if ids else {}
    return [found[i] for i in ids if i in found]
