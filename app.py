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
from pathlib import Path

from dotenv import load_dotenv
from flask import (
    Flask, render_template, request, redirect, url_for,
    session, send_from_directory, flash, abort,
)
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import inspect, text

from fill_logic import TEMPLATES, all_fields, required_field_keys, template_path

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
    return render_template("picker.html", templates=TEMPLATES)


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


def generated_filename(template_id, name, date_obj):
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
    while (GENERATED_DIR / candidate).exists():
        candidate = f"{base} ({n}).docx"
        n += 1
    return candidate


@app.route("/new/<template_id>", methods=["GET", "POST"])
@login_required
def new_document(template_id):
    spec = TEMPLATES.get(template_id)
    if spec is None:
        abort(404)

    if request.method == "POST":
        fields = all_fields(template_id)
        field_keys = [f["key"] for f in fields]
        values = {k: request.form.get(k, "").strip() for k in field_keys}

        missing = [k for k in required_field_keys(template_id) if not values[k]]
        if missing:
            flash("Please fill in all required fields: " + ", ".join(missing))
            return render_template(
                "form.html", spec=spec, template_id=template_id,
                data=request.form, today=date.today().isoformat(),
            )

        # Format checks (mobile number, national ID, ...) - only fields
        # whose spec declares a "pattern" get checked; a value that's
        # merely non-empty but the wrong shape (too short, letters where
        # there should be digits, ...) is caught here rather than ending
        # up wrong inside the generated document.
        format_errors = []
        for f in fields:
            pattern = f.get("pattern")
            if not pattern or not values.get(f["key"]):
                continue
            if not re.fullmatch(pattern, values[f["key"]]):
                format_errors.append(f.get("pattern_msg") or f"{f['label']} is not the right format.")
        if format_errors:
            for msg in format_errors:
                flash(msg)
            return render_template(
                "form.html", spec=spec, template_id=template_id,
                data=request.form, today=date.today().isoformat(),
            )

        # Email field (only some templates have one): always force the
        # @stm.com.eg domain - only the part before an "@" (if the person
        # typed one) is kept, so it's impossible to end up with any other
        # domain.
        if "email" in values:
            email_local = values["email"].split("@")[0].strip()
            values["email"] = f"{email_local}@stm.com.eg"

        date_input = request.form.get("date", "").strip()
        if date_input:
            try:
                date_obj = datetime.strptime(date_input, "%Y-%m-%d")
            except ValueError:
                flash("Invalid date")
                return render_template(
                    "form.html", spec=spec, template_id=template_id,
                    data=request.form, today=date.today().isoformat(),
                )
        else:
            date_obj = datetime.today()

        fill_data = dict(values)
        fill_data["date_obj"] = date_obj
        for extra in spec.get("extra_fields", []):
            fill_data[extra["key"]] = request.form.get(extra["key"], "").strip() or extra.get("default", "")

        doc_path = template_path(template_id)
        if not doc_path.exists():
            flash(f"{spec['doc_file']} is missing on the server.")
            return render_template(
                "form.html", spec=spec, template_id=template_id,
                data=request.form, today=date.today().isoformat(),
            )

        # Stored on disk under a human-readable, collision-safe name (see
        # generated_filename()) so the generated/ folder is browsable on
        # its own - e.g. "Yasmin Mohamed - Laptop Handover - 2026-09-10.docx".
        # The name a person sees when they download from the site is
        # computed separately in display_filename() and can differ (it
        # follows STM's Arabic document-naming convention).
        internal_name = generated_filename(template_id, values["name"], date_obj)
        output_path = GENERATED_DIR / internal_name
        spec["fill"](doc_path, output_path, fill_data)

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

        return redirect(url_for("done", record_id=record.id))

    return render_template(
        "form.html", spec=spec, template_id=template_id,
        data={}, today=date.today().isoformat(),
    )


@app.route("/done/<int:record_id>")
@login_required
def done(record_id):
    record = Handover.query.get_or_404(record_id)
    return render_template("done.html", record=record)


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


@app.route("/history/export.xlsx")
@login_required
def export_history():
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from io import BytesIO

    records = _filtered_history_query().all()

    wb = Workbook()
    ws = wb.active
    ws.title = "Handover history"

    headers = ["Name", "Document type", "Department", "Position", "National ID",
               "Handover date", "Generated by", "Generated at"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for r in records:
        ws.append([
            # National ID stays masked here too, same as the on-screen
            # history table - exporting the full number would undo the
            # point of masking it there in the first place.
            r.name, r.template_label, r.department, r.role, mask_id(r.govid),
            r.handover_date, r.created_by,
            r.created_at.strftime("%Y-%m-%d %H:%M") if r.created_at else "",
        ])

    # Reasonable column widths rather than Excel's cramped default.
    widths = [22, 26, 18, 20, 16, 14, 16, 18]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = f"handover-history-{date.today().isoformat()}.xlsx"
    from flask import send_file
    return send_file(
        buf, as_attachment=True, download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
    flash(f"Deleted the record for {name}.")
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
