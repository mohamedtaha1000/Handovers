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
templates.py's TEMPLATES registry, which is the single source of truth
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

from templates import (
    EMPLOYEE_KEYS, TEMPLATES, TEMPLATE_GROUPS,
    all_fields, required_field_keys, template_path,
)
import builders
import notify_email
import leaver_email
import asset_register
import documents
import employees
import settings
from documents import (
    batch_records, collect_values, display_filename, form_data_from_record,
    generated_filename, scoped_form, stamp_record, write_document,
)
from employees import (
    employee_identity, equipment_summary, find_duplicate,
    known_employees, leaver_lookup, matching_laptops, register_rows,
)
from models import Handover, Departure, create_all_and_migrate, db

app = Flask(__name__)

# Everything configurable lives in settings.py. Read through the module
# (settings.X) rather than importing the names, so a test that points the
# app at a temporary folder is actually seen by the code that writes
# there.
app.config["SECRET_KEY"] = settings.SECRET_KEY or secrets.token_hex(32)
app.config["SQLALCHEMY_DATABASE_URI"] = settings.DATABASE_URL
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

# The two leaver addresses as one lookup, so a message can ask for its
# own recipient by name. Built here rather than in settings.py because
# it pairs a setting with a message key, which is an app-level idea.
LEAVER_RECIPIENTS = {"ems": settings.EMS_TO, "resignation": settings.LEAVER_TO}

if settings.TEAM_PASSWORD == "changeme":
    app.logger.warning(
        "TEAM_PASSWORD is empty or not set - using the insecure default "
        "'changeme'. Set a real TEAM_PASSWORD in your .env file (or as an "
        "environment variable when you deploy)."
    )

db.init_app(app)
with app.app_context():
    create_all_and_migrate()


@app.context_processor
def notify_details():
    """Who the handover email goes to, available to every template so the
    buttons and the optional section can name them rather than saying
    "the asset owner"."""
    return {"notify_name": settings.NOTIFY_NAME, "notify_to": settings.NOTIFY_TO}


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
        elif not secrets.compare_digest(password, settings.TEAM_PASSWORD):
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



def save_document(template_id, values, fill_data, date_obj, record=None):
    """Generate the .docx and store it, as one step.

    Returns (record, problem). `problem` is a sentence to show the person
    and means nothing was saved - the file is written before the row is
    touched, so a template that fails to fill cannot leave a row pointing
    at a document that was never made.

    Pass `record` to regenerate an existing one; leave it out to make a
    new one.
    """
    editing = record is not None
    internal_name = generated_filename(
        template_id, values["name"], date_obj,
        exclude=record.filename if editing else None)

    problem = write_document(template_id, fill_data, internal_name)
    if problem:
        return None, problem

    if editing:
        # The old file is only removed once the new one exists, and only
        # if the name actually changed.
        if record.filename and record.filename != internal_name:
            old_file = settings.GENERATED_DIR / record.filename
            if old_file.exists():
                old_file.unlink()
        stamp_record(record, template_id, values, fill_data, date_obj, internal_name)
        record.updated_by = session.get("display_name", "-")
        record.updated_at = datetime.utcnow()
    else:
        record = stamp_record(Handover(), template_id, values, fill_data,
                              date_obj, internal_name)
        record.created_by = session.get("display_name", "-")
        db.session.add(record)

    db.session.commit()
    register_updated()
    return record, None


def render_form(template_id, data, record=None, duplicate=None):
    spec = TEMPLATES[template_id]
    return render_template(
        "form.html", spec=spec, template_id=template_id,
        data=data, today=date.today().isoformat(), record=record,
        duplicate=duplicate,
    )




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

        record, problem = save_document(template_id, values, fill_data, date_obj)
        if problem:
            flash(problem, "error")
            return render_form(template_id, request.form)

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

        record, problem = save_document(template_id, values, fill_data, date_obj,
                                        record=record)
        if problem:
            flash(problem, "error")
            return render_form(template_id, request.form, record=record)

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

        # Each document is saved as it is made, rather than all of them at
        # the end. It means a batch that fails on its third document
        # keeps the first two - file and row together - instead of
        # leaving two .docx files on disk that no record points at, which
        # is what the single commit at the end used to do. The person is
        # told exactly where it stopped so they can finish the rest.
        created = []
        for template_id in chosen:
            values, fill_data, date_obj = collected[template_id]
            record, problem = save_document(template_id, values, fill_data, date_obj)
            if problem:
                flash(problem, "error")
                if created:
                    done = ", ".join(TEMPLATES[t]["label"] for t in chosen[:len(created)])
                    flash(f"{len(created)} of {len(chosen)} were made and saved "
                          f"({done}). Only the rest still need doing.", "info")
                return show(request.form)
            created.append(record.id)
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



@app.route("/done-batch")
@login_required
def done_batch():
    records = batch_records(request.args.get("ids", ""))
    if not records:
        return redirect(url_for("history"))
    return render_template("done_batch.html", records=records,
                           ids=request.args.get("ids", ""),
                           # What each document actually handed over, so
                           # the list reads as equipment rather than as
                           # ten repetitions of the same date.
                           equipment={r.id: equipment_summary(r) for r in records},
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
            path = settings.GENERATED_DIR / record.filename
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
        records, greeting_name=settings.NOTIFY_NAME, detailed=detailed)
    query = urlencode({"subject": notify_email.subject_for(records), "body": body},
                      quote_via=quote)
    return f"mailto:{quote(settings.NOTIFY_TO, safe='@.')}?{query}"


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
    file_path = settings.GENERATED_DIR / record.filename
    if not file_path.exists():
        abort(404)
    return send_from_directory(
        settings.GENERATED_DIR, record.filename,
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
        asset_register.write_workbook(register_rows(), settings.REGISTER_PATH)
        return None
    except PermissionError:
        return (f"The laptop register ({settings.REGISTER_PATH.name}) is open in Excel, "
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
    if not settings.REGISTER_PATH.exists():
        problem = refresh_register()
        if problem:
            flash(problem, "error")
            return redirect(url_for("history"))
    return send_file(
        settings.REGISTER_PATH, as_attachment=True,
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
    file_path = settings.GENERATED_DIR / record.filename
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



