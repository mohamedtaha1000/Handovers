#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
STM Handover Documents — web app
==================================
A small internal Flask site that fills STM's handover/receipt Word
documents from a web form instead of the command line, and keeps a
searchable, deletable history of every document it has generated.

Supports multiple document templates (laptop handover, laptop
replacement, keyboard/mouse receipt, screen handover, ...) - see
fill_logic.py's TEMPLATES registry, which is the single source of truth
for what documents exist and what fields each one's form collects. Adding
a new document type later means adding one entry there plus a .docx file
in doc_templates/ - nothing in this file needs to change.

Run locally:
    pip install -r requirements.txt
    python app.py
    -> open http://127.0.0.1:5000

See README.md for environment variables and deployment notes.
"""

import json
import os
import re
import secrets
from datetime import datetime, date, timedelta
from functools import wraps
from urllib.parse import quote, urlencode
from pathlib import Path

from dotenv import load_dotenv
from flask import (
    Flask, render_template, request, redirect, url_for,
    session, send_from_directory, flash, abort,
)
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import inspect, text

from fill_logic import (
    TEMPLATES, TEMPLATE_GROUPS,
    all_fields, required_field_keys, template_path,
)
import notify_email
import leaver_email
import asset_register

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")  # reads SECRET_KEY / TEAM_PASSWORD from .env if present
GENERATED_DIR = BASE_DIR / "generated"
INSTANCE_DIR = BASE_DIR / "instance"
GENERATED_DIR.mkdir(exist_ok=True)
INSTANCE_DIR.mkdir(exist_ok=True)

app = Flask(__name__)

# --- Configuration (override these via environment variables when you
# deploy - see README.md) ------------------------------------------------
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
app.config["SQLALCHEMY_DATABASE_URI"] = os.environ.get(
    "DATABASE_URL", f"sqlite:///{INSTANCE_DIR / 'handovers.db'}"
)
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

# Who the "this has been handed over" email is addressed to. Kept in the
# environment rather than the code so a change of owner - or of person in
# the role - does not need an edit and a redeploy.
NOTIFY_TO = os.environ.get("HANDOVER_NOTIFY_TO", "").strip()
NOTIFY_NAME = os.environ.get("HANDOVER_NOTIFY_NAME", "Eng. Hegazy").strip()

# Where the two leaver messages go. Separate from NOTIFY_TO because they
# are a different conversation with a different team; either left unset
# just means the message opens with an empty To line, which is still
# faster than writing it out by hand.
EMS_TO = os.environ.get("HANDOVER_EMS_TO", "").strip()
LEAVER_TO = os.environ.get("HANDOVER_LEAVER_TO", "").strip()
LEAVER_RECIPIENTS = {"ems": EMS_TO, "resignation": LEAVER_TO}

# The laptop register. Beside the app by default, so on a machine where
# the folder is synced it simply appears - no download step, no second
# copy to go stale. Somewhere else via HANDOVER_REGISTER_PATH.
REGISTER_PATH = Path(os.environ.get(
    "HANDOVER_REGISTER_PATH", str(BASE_DIR / "laptop_register.xlsx")))

TEAM_PASSWORD = os.environ.get("TEAM_PASSWORD") or "changeme"
if TEAM_PASSWORD == "changeme":
    app.logger.warning(
        "TEAM_PASSWORD is empty or not set - using the insecure default "
        "'changeme'. Set a real TEAM_PASSWORD in your .env file (or as an "
        "environment variable when you deploy)."
    )

db = SQLAlchemy(app)


class Handover(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    template_id = db.Column(db.String(60), nullable=False, default="laptop_handover")
    name = db.Column(db.String(200), nullable=False)
    department = db.Column(db.String(120))
    role = db.Column(db.String(120))
    govid = db.Column(db.String(40))
    handover_date = db.Column(db.String(20))
    # Every field the form collected for this document (mobile, email,
    # code, and whatever device-specific fields that template has) -
    # kept as JSON since different templates collect very different
    # fields. name/department/role/govid above are duplicated out as
    # real columns just so the history list can show and search them
    # without needing to parse JSON for every row.
    fields_json = db.Column(db.Text)
    filename = db.Column(db.String(300), nullable=False)
    created_by = db.Column(db.String(120))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    # Set only once a record has actually been corrected, so "never
    # edited" stays distinguishable from "edited by the same person who
    # created it, straight away".
    updated_by = db.Column(db.String(120))
    updated_at = db.Column(db.DateTime)

    @property
    def fields(self):
        try:
            return json.loads(self.fields_json) if self.fields_json else {}
        except (TypeError, ValueError):
            return {}

    @property
    def template_label(self):
        spec = TEMPLATES.get(self.template_id)
        return spec["label"] if spec else self.template_id


class Departure(db.Model):
    """Someone who has left. Kept per person rather than per document:
    they hand back everything at once, and the register works out which
    rows that touches."""
    id = db.Column(db.Integer, primary_key=True)
    # employee_identity(): the employee code where there is one, the
    # national ID behind it, the name as a last resort.
    identity = db.Column(db.String(160), unique=True, index=True)
    name = db.Column(db.String(200))
    left_on = db.Column(db.String(20))
    recorded_by = db.Column(db.String(120))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


with app.app_context():
    db.create_all()

    # Lightweight migration: add any columns older databases don't have
    # yet, without touching (or losing) existing rows. SQLite's ALTER
    # TABLE only supports adding columns, which is all we need here.
    inspector = inspect(db.engine)
    if "handover" in inspector.get_table_names():
        existing_cols = {c["name"] for c in inspector.get_columns("handover")}
        with db.engine.begin() as conn:
            if "template_id" not in existing_cols:
                conn.execute(text(
                    "ALTER TABLE handover ADD COLUMN template_id VARCHAR(60) "
                    "DEFAULT 'laptop_handover'"
                ))
            if "fields_json" not in existing_cols:
                conn.execute(text("ALTER TABLE handover ADD COLUMN fields_json TEXT"))
            if "govid" not in existing_cols:
                conn.execute(text("ALTER TABLE handover ADD COLUMN govid VARCHAR(40)"))
            if "updated_by" not in existing_cols:
                conn.execute(text("ALTER TABLE handover ADD COLUMN updated_by VARCHAR(120)"))
            if "updated_at" not in existing_cols:
                conn.execute(text("ALTER TABLE handover ADD COLUMN updated_at DATETIME"))


@app.context_processor
def notify_details():
    """Who the handover email goes to, available to every template so the
    buttons and the optional section can name them rather than saying
    "the asset owner"."""
    return {"notify_name": NOTIFY_NAME, "notify_to": NOTIFY_TO}


@app.template_filter("initials")
def initials(value):
    """One or two initials for the topbar badge. Takes the first letter
    of the first and last word, which works for an Arabic name as well as
    a Latin one."""
    parts = [p for p in str(value or "").split() if p]
    if not parts:
        return "?"
    if len(parts) == 1:
        return parts[0][:2].upper()
    return (parts[0][0] + parts[-1][0]).upper()


@app.template_filter("mask_id")
def mask_id(value):
    """Show only the last 4 characters of a sensitive value (national ID)
    in list views, so it isn't fully exposed to everyone who can see the
    history page. The full value is still in the generated .docx and on
    the confirmation page right after creating it."""
    if not value:
        return ""
    value = str(value)
    if len(value) <= 4:
        return "•" * len(value)
    return "•" * (len(value) - 4) + value[-4:]


# ----------------------------------------------------------------------
# Auth (single shared team password - see README for stronger options)
# ----------------------------------------------------------------------

def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        password = request.form.get("password", "")
        display_name = request.form.get("your_name", "").strip()
        if not display_name:
            error = "Please enter your name"
        elif not secrets.compare_digest(password, TEAM_PASSWORD):
            error = "Wrong password"
        else:
            session["logged_in"] = True
            session["display_name"] = display_name
            nxt = request.args.get("next") or url_for("index")
            return redirect(nxt)
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ----------------------------------------------------------------------
# Template picker (home page)
# ----------------------------------------------------------------------

@app.route("/")
@login_required
def index():
    # ?employee=<record id> carries "another document for this person"
    # through the picker, so whichever type is chosen next opens with the
    # employee half already filled in.
    return render_template(
        "picker.html", templates=TEMPLATES, groups=grouped_templates(),
        employee=request.args.get("employee", type=int),
    )


# ----------------------------------------------------------------------
# Document form (one per template)
# ----------------------------------------------------------------------

FILENAME_UNSAFE_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


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
    while (GENERATED_DIR / candidate).exists() and candidate != exclude:
        candidate = f"{base} ({n}).docx"
        n += 1
    return candidate


# ----------------------------------------------------------------------
# What the screens need to know beyond the records themselves
# ----------------------------------------------------------------------

def grouped_templates():
    """The ten document types arranged for the picker, in group order."""
    groups = []
    for name in TEMPLATE_GROUPS:
        members = [(tid, spec) for tid, spec in TEMPLATES.items()
                   if spec.get("group") == name]
        if members:
            groups.append((name, members))
    # A type whose group was mistyped still has to appear somewhere.
    loose = [(tid, spec) for tid, spec in TEMPLATES.items()
             if spec.get("group") not in TEMPLATE_GROUPS]
    if loose:
        groups.append(("Other", loose))
    return groups


# The fields worth showing in the history table's equipment column, in
# the order we would rather have them: what the thing is, then which one.
EQUIPMENT_KEYS = ("model", "new_model", "brand", "capacity", "old_model")
SERIAL_KEYS = ("serial", "new_serial", "sim_number", "old_serial")


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


def active_filters(q, type_filter, date_from, date_to):
    """The filters currently narrowing the history, each as a label plus
    the query string that removes just that one - so they can be read back
    in words and cleared individually."""
    current = {"q": q, "type": type_filter, "from": date_from, "to": date_to}
    chips = []
    def add(key, label):
        remaining = {k: v for k, v in current.items() if v and k != key}
        chips.append({"label": label, "remove": url_for("history", **remaining)})
    if q:
        add("q", f'"{q}"')
    if type_filter and type_filter in TEMPLATES:
        add("type", TEMPLATES[type_filter]["label"])
    if date_from and date_to:
        chips.append({"label": f"{date_from} to {date_to}",
                      "remove": url_for("history", **{k: v for k, v in current.items()
                                                      if v and k not in ("from", "to")})})
    elif date_from:
        add("from", f"from {date_from}")
    elif date_to:
        add("to", f"until {date_to}")
    return chips


# ----------------------------------------------------------------------
# Employees already on file
#
# Every template asks for the same seven employee fields, and somebody
# collecting a laptop usually collects a mouse, a keyboard and a headset
# too - so the same national ID was being typed out four times, each one
# a fresh chance to get a digit wrong on a document that gets signed.
# The details are already in the history; these helpers hand them back.
# ----------------------------------------------------------------------

# Fields that describe the person rather than the equipment. Only these
# are ever copied from a previous document - the serial number of the
# laptop they were given last year must not follow them onto a new one.
EMPLOYEE_KEYS = ("name", "department", "role", "mobile", "email", "code", "govid")


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


@app.route("/api/employees")
@login_required
def api_employees():
    """Feeds the "reuse a previous employee" suggestions on the form."""
    query = request.args.get("q", "").strip().lower()
    people = known_employees()
    if query:
        people = [p for p in people
                  if query in p["name"].lower()
                  or query in p["code"].lower()
                  or query in p["department"].lower()]
    return {"employees": people[:8]}


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


# ----------------------------------------------------------------------
# Shared form handling
#
# Creating a document and correcting one are the same job apart from what
# happens at the end, so both routes go through the helpers below rather
# than each carrying their own copy of the rules. That matters more than
# it saves typing: a validation rule or a normalisation that lived in
# only one of the two would mean a value the form rejects on the way in
# could still be edited back in afterwards.
# ----------------------------------------------------------------------

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


def render_form(template_id, data, record=None, duplicate=None):
    spec = TEMPLATES[template_id]
    return render_template(
        "form.html", spec=spec, template_id=template_id,
        data=data, today=date.today().isoformat(), record=record,
        duplicate=duplicate,
    )


def write_document(template_id, fill_data, internal_name):
    """Generate the .docx into generated/ under `internal_name`.

    Written to a temporary file first and moved into place only once it
    has been produced in full, so a re-generation that fails part way
    through can't leave a half-written document where a good one was."""
    spec = TEMPLATES[template_id]
    doc_path = template_path(template_id)
    if not doc_path.exists():
        return f"{spec['doc_file']} is missing on the server."
    tmp_path = GENERATED_DIR / f".tmp-{secrets.token_hex(8)}.docx"
    try:
        spec["fill"](doc_path, tmp_path, fill_data)
        os.replace(tmp_path, GENERATED_DIR / internal_name)
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise
    return None


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


# ----------------------------------------------------------------------
# Create a document
# ----------------------------------------------------------------------

@app.route("/new/<template_id>", methods=["GET", "POST"])
@login_required
def new_document(template_id):
    spec = TEMPLATES.get(template_id)
    if spec is None:
        abort(404)

    if request.method == "POST":
        values, fill_data, date_obj, errors = collect_values(template_id, request.form)
        if errors:
            for message in errors:
                flash(message, "error")
            return render_form(template_id, request.form)

        # Stored on disk under a human-readable, collision-safe name (see
        # generated_filename()) so the generated/ folder is browsable on
        # its own - e.g. "Yasmin Mohamed - Laptop Handover - 2026-09-10.docx".
        # The name a person sees when they download from the site is
        # computed separately in display_filename() and can differ (it
        # follows STM's Arabic document-naming convention).
        # Ask once before making a second document of the same type for
        # the same person. The answer travels back in a hidden field, so
        # confirming doesn't lose anything already typed.
        if not request.form.get("confirm_duplicate"):
            duplicate = find_duplicate(template_id, values)
            if duplicate is not None:
                return render_form(template_id, request.form, duplicate=duplicate)

        internal_name = generated_filename(template_id, values["name"], date_obj)
        problem = write_document(template_id, fill_data, internal_name)
        if problem:
            flash(problem, "error")
            return render_form(template_id, request.form)

        record = Handover(
            template_id=template_id,
            name=values["name"], department=values["department"], role=values["role"],
            govid=values.get("govid", ""),
            handover_date=f"{date_obj.day}/{date_obj.month}/{date_obj.year}",
            fields_json=json.dumps(fill_data, default=str),
            filename=internal_name,
            created_by=session.get("display_name", "-"),
        )
        db.session.add(record)
        db.session.commit()
        register_updated()

        return redirect(url_for("done", record_id=record.id))

    # "Another document for this person" arrives as ?employee=<record id>,
    # which pre-fills the employee half of the form and leaves the
    # equipment half empty.
    prefill = {}
    source_id = request.args.get("employee", type=int)
    if source_id:
        source = db.session.get(Handover, source_id)
        if source is not None:
            fields = source.fields
            prefill = {k: fields.get(k, "") for k in EMPLOYEE_KEYS if fields.get(k)}
            prefill.setdefault("name", source.name or "")
            prefill.setdefault("department", source.department or "")
            prefill.setdefault("role", source.role or "")
            prefill.setdefault("govid", source.govid or "")
            for f in all_fields(template_id):
                suffix = f.get("suffix")
                value = prefill.get(f["key"], "")
                if suffix and isinstance(value, str) and value.endswith(suffix):
                    prefill[f["key"]] = value[: -len(suffix)]
    return render_form(template_id, prefill)


# ----------------------------------------------------------------------
# Correct a document that has already been generated
#
# Re-generates the .docx in place rather than adding a second one: a
# correction is the same handover, not a new one, so the record keeps a
# single current document and the history keeps a single row. The old
# file is removed when the correction changes the employee name or the
# date, since those are what the filename is built from.
# ----------------------------------------------------------------------

@app.route("/edit/<int:record_id>", methods=["GET", "POST"])
@login_required
def edit_document(record_id):
    record = Handover.query.get_or_404(record_id)
    template_id = record.template_id
    if template_id not in TEMPLATES:
        flash("That document was made from a template this site no longer has.", "error")
        return redirect(url_for("history"))

    if request.method == "POST":
        values, fill_data, date_obj, errors = collect_values(template_id, request.form)
        if errors:
            for message in errors:
                flash(message, "error")
            return render_form(template_id, request.form, record=record)

        previous_name = record.filename
        internal_name = generated_filename(
            template_id, values["name"], date_obj, exclude=previous_name)
        problem = write_document(template_id, fill_data, internal_name)
        if problem:
            flash(problem, "error")
            return render_form(template_id, request.form, record=record)

        if previous_name != internal_name:
            old_file = GENERATED_DIR / previous_name
            if old_file.exists():
                old_file.unlink()

        record.name = values["name"]
        record.department = values["department"]
        record.role = values["role"]
        record.govid = values.get("govid", "")
        record.handover_date = f"{date_obj.day}/{date_obj.month}/{date_obj.year}"
        record.fields_json = json.dumps(fill_data, default=str)
        record.filename = internal_name
        record.updated_by = session.get("display_name", "-")
        record.updated_at = datetime.utcnow()
        db.session.commit()
        register_updated()

        flash(f"Saved. The document for {record.name} has been generated again.", "success")
        return redirect(url_for("done", record_id=record.id))

    return render_form(template_id, form_data_from_record(record), record=record)


# ----------------------------------------------------------------------
# Several documents for one person, in one pass
#
# A new joiner typically collects a laptop, a mouse, a keyboard and a
# headset on their first morning - four documents whose employee half is
# identical. This collects that half once and each piece of equipment
# separately, then generates the lot.
#
# Device fields are namespaced per template ("headset_handover__model")
# because the templates genuinely collide: nearly all of them ask for a
# "model", "serial" and "color", and without the prefix the headset's
# serial would overwrite the laptop's.
# ----------------------------------------------------------------------

# Asked once on the combined form and used by every document in it.
# The employee half, the date - and the computer name, because a person
# handed a laptop, a mouse and a headset on the same morning is sitting
# at one machine, and typing its name three times would only be a way to
# get it wrong twice.
SHARED_BATCH_KEYS = EMPLOYEE_KEYS + ("date", "computer_name")


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


@app.route("/new", methods=["POST"])
@login_required
def new_batch_start():
    """Picker -> combined form. A single tick just goes to the normal
    one-document form, which is a better page for that job."""
    chosen = [t for t in request.form.getlist("template_id") if t in TEMPLATES]
    if not chosen:
        flash("Pick at least one document to create.", "error")
        return redirect(url_for("index"))
    employee = request.args.get("employee", type=int)
    if len(chosen) == 1:
        return redirect(url_for("new_document", template_id=chosen[0], employee=employee))
    return redirect(url_for("new_batch", types=",".join(chosen), employee=employee))


@app.route("/new-batch", methods=["GET", "POST"])
@login_required
def new_batch():
    chosen = [t for t in request.values.get("types", "").split(",") if t in TEMPLATES]
    if len(chosen) < 2:
        return redirect(url_for("index"))

    def show(data, duplicates=None, per_template_errors=None):
        return render_template(
            "batch_form.html", chosen=chosen, templates=TEMPLATES,
            data=data, today=date.today().isoformat(),
            types=",".join(chosen), duplicates=duplicates or {},
            errors=per_template_errors or {},
        )

    if request.method == "POST":
        collected, errors, duplicates = {}, {}, {}
        for template_id in chosen:
            values, fill_data, date_obj, problems = collect_values(
                template_id, scoped_form(template_id, request.form))
            if problems:
                errors[template_id] = problems
            collected[template_id] = (values, fill_data, date_obj)
            if not request.form.get("confirm_duplicate"):
                found = find_duplicate(template_id, values)
                if found is not None:
                    duplicates[template_id] = found

        if errors:
            # The employee half is shared, so the same missing name would
            # otherwise be reported once per document.
            seen = set()
            for template_id, problems in errors.items():
                for message in problems:
                    if message not in seen:
                        seen.add(message)
                        flash(message, "error")
            return show(request.form, duplicates=None, per_template_errors=errors)
        if duplicates:
            return show(request.form, duplicates=duplicates)

        created = []
        for template_id in chosen:
            values, fill_data, date_obj = collected[template_id]
            internal_name = generated_filename(template_id, values["name"], date_obj)
            problem = write_document(template_id, fill_data, internal_name)
            if problem:
                flash(problem, "error")
                return show(request.form)
            record = Handover(
                template_id=template_id,
                name=values["name"], department=values["department"],
                role=values["role"], govid=values.get("govid", ""),
                handover_date=f"{date_obj.day}/{date_obj.month}/{date_obj.year}",
                fields_json=json.dumps(fill_data, default=str),
                filename=internal_name,
                created_by=session.get("display_name", "-"),
            )
            db.session.add(record)
            db.session.flush()
            created.append(record.id)
        db.session.commit()
        register_updated()
        return redirect(url_for("done_batch", ids=",".join(str(i) for i in created)))

    prefill = {}
    source_id = request.args.get("employee", type=int)
    if source_id:
        source = db.session.get(Handover, source_id)
        if source is not None:
            fields = source.fields
            prefill = {k: fields.get(k, "") for k in EMPLOYEE_KEYS if fields.get(k)}
            if prefill.get("email", "").endswith("@stm.com.eg"):
                prefill["email"] = prefill["email"].split("@")[0]
    return show(prefill)


def batch_records(raw_ids):
    ids = [int(i) for i in raw_ids.split(",") if i.strip().isdigit()]
    found = {r.id: r for r in Handover.query.filter(Handover.id.in_(ids)).all()} if ids else {}
    return [found[i] for i in ids if i in found]


@app.route("/done-batch")
@login_required
def done_batch():
    records = batch_records(request.args.get("ids", ""))
    if not records:
        return redirect(url_for("history"))
    return render_template("done_batch.html", records=records,
                           ids=request.args.get("ids", ""),
                           mailto=mailto_link(records))


@app.route("/files.zip")
@login_required
def download_batch():
    """All of a batch's documents in one download, named the way a single
    download names them."""
    import zipfile
    from io import BytesIO
    from flask import send_file

    records = batch_records(request.args.get("ids", ""))
    if not records:
        abort(404)

    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as bundle:
        used = set()
        for record in records:
            path = GENERATED_DIR / record.filename
            if not path.exists():
                continue
            name = display_filename(record.template_id, record.name)
            # Two documents of the same type for the same person would
            # otherwise collide inside the zip and one would be lost.
            stem, n = name[:-5], 2
            while name in used:
                name = f"{stem} ({n}).docx"
                n += 1
            used.add(name)
            bundle.write(path, name)
    buf.seek(0)

    safe = FILENAME_UNSAFE_RE.sub("", records[0].name).strip() or "documents"
    return send_file(buf, as_attachment=True, download_name=f"{safe}.zip",
                     mimetype="application/zip")


# ----------------------------------------------------------------------
# Telling the asset owner
#
# A mailto: link rather than a file to download: one click opens Outlook's
# new-message window with the address, subject and details already in it,
# and the person sends it from their own mailbox. Plain text is all a
# mailto can carry - no ruled table, no attachment - which is the trade
# for not having to open a downloaded file first. notify_email lays the
# details out as labelled sections, which read the same in any font.
# ----------------------------------------------------------------------

# Outlook on Windows stops honouring a mailto around 2,000 characters.
# One handover encodes to well under half that even with an Arabic name;
# a batch of seven or more is what can reach it.
MAILTO_LIMIT = 1900


def build_mailto(records, detailed):
    # No sender is passed and no sign-off is written: the draft opens in
    # whoever's Outlook clicked the link, and their own signature goes
    # under it.
    body = notify_email.build_body(
        records, greeting_name=NOTIFY_NAME, detailed=detailed)
    query = urlencode({"subject": notify_email.subject_for(records), "body": body},
                      quote_via=quote)
    return f"mailto:{quote(NOTIFY_TO, safe='@.')}?{query}"


def mailto_link(records):
    """The mailto: URL for these documents, or None if there is nothing to
    describe.

    A long batch loses detail blocks from the end until the URL fits -
    the message then names those items instead of describing them, which
    is better than handing the mail client a URL it cuts off mid-word.
    """
    if not records:
        return None
    link = build_mailto(records, records)
    detailed = list(records)
    while len(link) > MAILTO_LIMIT and len(detailed) > 1:
        detailed.pop()
        link = build_mailto(records, detailed)
    return link


@app.route("/done/<int:record_id>")
@login_required
def done(record_id):
    record = Handover.query.get_or_404(record_id)
    return render_template("done.html", record=record,
                           mailto=mailto_link([record]))


@app.route("/file/<int:record_id>")
@login_required
def get_file(record_id):
    record = Handover.query.get_or_404(record_id)
    file_path = GENERATED_DIR / record.filename
    if not file_path.exists():
        abort(404)
    return send_from_directory(
        GENERATED_DIR, record.filename,
        as_attachment=True, download_name=display_filename(record.template_id, record.name),
    )


# ----------------------------------------------------------------------
# History (search + permanent delete)
# ----------------------------------------------------------------------

HISTORY_PAGE_SIZE = 50


@app.route("/history")
@login_required
def history():
    q = request.args.get("q", "").strip()
    type_filter = request.args.get("type", "").strip()
    date_from = request.args.get("from", "").strip()
    date_to = request.args.get("to", "").strip()
    try:
        page = max(1, int(request.args.get("page", 1)))
    except ValueError:
        page = 1

    query = Handover.query.order_by(Handover.created_at.desc())
    if q:
        like = f"%{q}%"
        query = query.filter(
            db.or_(
                Handover.name.like(like),
                Handover.department.like(like),
                Handover.role.like(like),
            )
        )
    if type_filter and type_filter in TEMPLATES:
        query = query.filter(Handover.template_id == type_filter)
    # Date range filters on created_at (a real datetime column) rather
    # than handover_date (a free-form "D/M/YYYY" string the person typed
    # in the form, not reliably sortable/comparable) - this is "when the
    # document was generated", which is what a person browsing history
    # usually means by a date range anyway.
    if date_from:
        try:
            query = query.filter(Handover.created_at >= datetime.strptime(date_from, "%Y-%m-%d"))
        except ValueError:
            date_from = ""
    if date_to:
        try:
            end = datetime.strptime(date_to, "%Y-%m-%d") + timedelta(days=1)
            query = query.filter(Handover.created_at < end)
        except ValueError:
            date_to = ""

    total = query.count()
    pages = max(1, (total + HISTORY_PAGE_SIZE - 1) // HISTORY_PAGE_SIZE)
    page = min(page, pages)
    records = query.offset((page - 1) * HISTORY_PAGE_SIZE).limit(HISTORY_PAGE_SIZE).all()

    return render_template(
        "history.html", records=records, q=q, type_filter=type_filter,
        date_from=date_from, date_to=date_to, templates=TEMPLATES,
        page=page, pages=pages, total=total, page_size=HISTORY_PAGE_SIZE,
        equipment={r.id: equipment_summary(r) for r in records},
        mailto={r.id: mailto_link([r]) for r in records},
        filters=active_filters(q, type_filter, date_from, date_to),
        matching=total,
        # `total` is what the current filters match; the heading wants the
        # size of the whole log, or it reads as though filtering deleted
        # everything else.
        grand_total=Handover.query.count(),
    )


def _filtered_history_query():
    """Builds the same filtered (but unpaginated) query used by both the
    history page and the Excel export, so the two can never drift apart -
    whatever's currently filtered/searched on screen is exactly what gets
    exported."""
    q = request.args.get("q", "").strip()
    type_filter = request.args.get("type", "").strip()
    date_from = request.args.get("from", "").strip()
    date_to = request.args.get("to", "").strip()

    query = Handover.query.order_by(Handover.created_at.desc())
    if q:
        like = f"%{q}%"
        query = query.filter(
            db.or_(
                Handover.name.like(like),
                Handover.department.like(like),
                Handover.role.like(like),
            )
        )
    if type_filter and type_filter in TEMPLATES:
        query = query.filter(Handover.template_id == type_filter)
    if date_from:
        try:
            query = query.filter(Handover.created_at >= datetime.strptime(date_from, "%Y-%m-%d"))
        except ValueError:
            pass
    if date_to:
        try:
            end = datetime.strptime(date_to, "%Y-%m-%d") + timedelta(days=1)
            query = query.filter(Handover.created_at < end)
        except ValueError:
            pass
    return query


# ----------------------------------------------------------------------
# The laptop register
#
# A spreadsheet rebuilt from the database whenever anything changes, so
# it can never drift from the records behind it. See asset_register.py
# for the shape of it.
# ----------------------------------------------------------------------

def register_rows():
    records = Handover.query.all()
    departures = {d.identity: d for d in Departure.query.all()}
    return asset_register.build_rows(records, employee_identity, departures)


def refresh_register():
    """Rebuild the sheet. Returns None on success, or a sentence saying
    why not.

    This is called from the middle of generating a document, so it must
    never raise: a register that could not be written is a nuisance, and
    losing the document that was just signed is not. The usual cause on
    Windows is the file being open in Excel, which locks it - worth
    saying plainly rather than reporting as an error.
    """
    try:
        asset_register.write_workbook(register_rows(), REGISTER_PATH)
        return None
    except PermissionError:
        return (f"The laptop register ({REGISTER_PATH.name}) is open in Excel, "
                f"so it could not be updated. Close it and the next document "
                f"will bring it up to date.")
    except Exception as problem:            # noqa: BLE001 - never block the document
        app.logger.exception("register rebuild failed")
        return f"The laptop register could not be updated: {problem}"


def register_updated():
    """Rebuild, and flash only if something went wrong. Success is
    silent: it happens on every single document and a message saying so
    every time would be noise the person learns to ignore."""
    problem = refresh_register()
    if problem:
        flash(problem, "info")


@app.route("/register.xlsx")
@login_required
def download_register():
    """The file itself. It lives beside the app, but the app may be on a
    different machine from whoever wants to read it."""
    from flask import send_file
    if not REGISTER_PATH.exists():
        problem = refresh_register()
        if problem:
            flash(problem, "error")
            return redirect(url_for("history"))
    return send_file(
        REGISTER_PATH, as_attachment=True,
        download_name=f"laptop-register-{date.today().isoformat()}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/register/rebuild", methods=["POST"])
@login_required
def rebuild_register():
    problem = refresh_register()
    if problem:
        flash(problem, "error")
    else:
        flash(f"Laptop register rebuilt — {len(register_rows())} assignments.",
              "success")
    return redirect(request.form.get("next") or url_for("history"))


# ----------------------------------------------------------------------
# Someone is leaving
#
# Two messages go out when an employee resigns. Every detail is typed
# here rather than looked up: people leave who never had a document
# generated for them, and a page that could only describe employees on
# file would be useless exactly when it was needed.
#
# The form is plain GET, so it works with JavaScript off - the details
# come back in the URL and the links are rebuilt server-side. With
# JavaScript on, the links carry a token per field and are rewritten as
# the boxes are typed, so the buttons are live and nothing has to be
# submitted at all.
# ----------------------------------------------------------------------

# Where a computer's serial number lives, per document type. A
# replacement issues a new machine, so its NEW serial is the one the
# person is still holding on the day they leave. A headset receipt has a
# serial too and it is not the one the security team is asking about, so
# only these two count.
COMPUTER_SERIAL_KEYS = {
    "laptop_handover": "serial",
    "laptop_replacement": "new_serial",
}


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


@app.route("/api/leavers")
@login_required
def api_leavers():
    """Feeds the suggestions on the leaver page. Deliberately a different
    endpoint from /api/employees: that one must never hand a serial
    number to a handover form, and this one exists to hand one over."""
    query = request.args.get("q", "").strip().lower()
    people = leaver_lookup()
    if query:
        people = [p for p in people
                  if query in p["name"].lower()
                  or query in p["code"].lower()
                  or query in p["department"].lower()]
    return {"employees": people[:8]}


def leaver_links(person):
    """A mailto: for each of the two messages, from one set of details."""
    links = {}
    for message in leaver_email.MESSAGES:
        to = LEAVER_RECIPIENTS.get(message["key"], "")
        query = urlencode({"subject": message["subject"](person),
                           "body": message["body"](person)}, quote_via=quote)
        links[message["key"]] = f"mailto:{quote(to, safe='@.')}?{query}"
    return links


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


@app.route("/api/holdings")
@login_required
def api_holdings():
    """What the leaver page shows under its two email buttons, refreshed
    as the boxes are typed."""
    rows = matching_laptops(request.args.get("name", ""),
                            request.args.get("serial", ""),
                            request.args.get("code", ""))
    return {
        "laptops": [{"model": r["model"], "serial": r["serial"],
                     "status": r["status"], "assigned_on": asset_register.show_date(r["date"]),
                     "assigned_by": r["by"], "name": r["name"]}
                    for r in rows],
        "open": sum(1 for r in rows if r["status"] == asset_register.HELD),
    }


@app.route("/leaver/left", methods=["POST"])
@login_required
def mark_left():
    """Record that someone has gone, and rebuild the register around it.

    Stored against the same identity the rest of the app uses, so it
    follows the person rather than one document: everything they were
    ever issued flips to Left together.
    """
    # The page's own button asks for JSON: it is opening Outlook twice in
    # the same click and must not navigate away to a redirect. The plain
    # form POST (no JavaScript) still gets the redirect and the flash.
    wants_json = request.form.get("format") == "json"

    # The page greys its buttons out until all five details are there.
    # That is a courtesy, not a guard: anything a browser enforces can be
    # turned off in the developer tools, and this request writes to the
    # register. So the same rule is checked here, where it cannot be
    # edited away, and the request is refused rather than half-applied.
    typed = {f["key"]: request.form.get(f["key"], "").strip()
             for f in leaver_email.FORM_FIELDS}
    missing = [f["label"] for f in leaver_email.FORM_FIELDS if not typed[f["key"]]]
    if missing:
        message = ("Nothing was changed: " + ", ".join(missing)
                   + (" is" if len(missing) == 1 else " are")
                   + " still empty, and all five are needed before anyone can "
                     "be marked as left.")
        if wants_json:
            return {"marked": 0, "message": message, "incomplete": missing}, 400
        flash(message, "error")
        return redirect(url_for("leaver", **request.form.to_dict(flat=True)))

    name, serial = typed["name"], typed["serial"]
    code = request.form.get("code", "").strip()
    rows = matching_laptops(name, serial, code)
    if not rows:
        message = ("Nothing on file matches those details, so there was "
                   "nothing to mark in the register.")
        if wants_json:
            return {"marked": 0, "message": message}
        flash(message + " The two emails still work.", "info")
        return redirect(url_for("leaver", **request.form.to_dict(flat=True)))

    identities = {r["identity"] for r in rows}
    today = date.today()
    left_on = f"{today.day}/{today.month}/{today.year}"
    for identity in identities:
        existing = Departure.query.filter_by(identity=identity).first()
        if existing is None:
            db.session.add(Departure(
                identity=identity, name=rows[0]["name"], left_on=left_on,
                recorded_by=session.get("display_name", "-")))
        else:
            existing.left_on = left_on
            existing.recorded_by = session.get("display_name", "-")
    db.session.commit()

    problem = refresh_register()
    machines = ", ".join(f"{r['model']} ({r['serial']})" for r in rows if r["serial"])
    message = (f"{rows[0]['name']} marked as left. "
               f"{len(rows)} laptop{'' if len(rows) == 1 else 's'} in the register "
               f"updated{': ' + machines if machines else ''}.")
    if wants_json:
        return {"marked": len(rows), "message": message, "problem": problem or ""}
    if problem:
        flash(problem, "info")
    flash(message, "success")
    return redirect(url_for("leaver", **request.form.to_dict(flat=True)))


@app.route("/leaver")
@login_required
def leaver():
    typed = {f["key"]: request.args.get(f["key"], "").strip()
             for f in leaver_email.FORM_FIELDS}
    return render_template(
        "leaver.html",
        fields=leaver_email.FORM_FIELDS, typed=typed,
        holdings=matching_laptops(typed.get("name", ""), typed.get("serial", "")),
        messages=leaver_email.MESSAGES, recipients=LEAVER_RECIPIENTS,
        links=leaver_links(typed),
        # The same two links with a token wherever a value goes, for the
        # browser to fill in as the boxes are typed.
        token_links=leaver_links({**leaver_email.TOKENS,
                                  "dept_clause": leaver_email.CLAUSE_TOKEN}),
        tokens=leaver_email.TOKENS,
        clause_token=leaver_email.CLAUSE_TOKEN,
        clause_template=leaver_email.CLAUSE_TEMPLATE,
    )


@app.route("/delete/<int:record_id>", methods=["POST"])
@login_required
def delete_history(record_id):
    record = Handover.query.get_or_404(record_id)
    file_path = GENERATED_DIR / record.filename
    if file_path.exists():
        file_path.unlink()
    name = record.name
    db.session.delete(record)
    db.session.commit()
    register_updated()
    flash(f"Deleted the record for {name}, and its generated file.", "success")
    # Preserve whatever search/filter/page the delete was performed from,
    # so deleting a row from page 3 of a filtered view doesn't bounce the
    # person back to an unfiltered page 1.
    keep = {}
    for key in ("q", "type", "from", "to", "page"):
        val = request.form.get(key, "")
        if val:
            keep[key] = val
    return redirect(url_for("history", **keep))


if __name__ == "__main__":
    debug = os.environ.get("FLASK_DEBUG", "1") == "1"
    app.run(debug=debug, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
