#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
STM Laptop Handover — web app
==============================
A small internal Flask site that fills the "محضر تسليم لاب توب" Word
document from a web form instead of the command line, and keeps a
searchable history of every document it has generated.

Run locally:
    pip install -r requirements.txt
    python app.py
    -> open http://127.0.0.1:5000

See README.md for environment variables and deployment notes.
"""

import os
import re
import secrets
import uuid
from datetime import datetime, date
from functools import wraps
from pathlib import Path

from dotenv import load_dotenv
from flask import (
    Flask, render_template, request, redirect, url_for,
    session, send_from_directory, flash, abort,
)
from flask_sqlalchemy import SQLAlchemy

from fill_logic import fill_document, REQUIRED_FIELDS

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")  # reads SECRET_KEY / TEAM_PASSWORD from .env if present
GENERATED_DIR = BASE_DIR / "generated"
INSTANCE_DIR = BASE_DIR / "instance"
TEMPLATE_PATH = BASE_DIR / "Template.docx"
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
    name = db.Column(db.String(200), nullable=False)
    department = db.Column(db.String(120))
    role = db.Column(db.String(120))
    mobile = db.Column(db.String(40))
    email = db.Column(db.String(200))
    code = db.Column(db.String(40))
    govid = db.Column(db.String(40))
    serial = db.Column(db.String(80))
    model = db.Column(db.String(80))
    cpu = db.Column(db.String(40))
    handover_date = db.Column(db.String(20))
    filename = db.Column(db.String(300), nullable=False)
    created_by = db.Column(db.String(120))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


with app.app_context():
    db.create_all()


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
# Main form
# ----------------------------------------------------------------------

FILENAME_UNSAFE_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def display_filename(name):
    """The clean, human-facing filename a person sees when they download -
    always "استلام لابتوب(Name).docx", no timestamps or ids in it."""
    safe_name = FILENAME_UNSAFE_RE.sub("", name).strip() or "employee"
    return f"استلام لابتوب({safe_name}).docx"


@app.route("/", methods=["GET", "POST"])
@login_required
def index():
    if request.method == "POST":
        values = {k: request.form.get(k, "").strip() for k in REQUIRED_FIELDS}

        missing = [k for k in REQUIRED_FIELDS if not values[k]]
        if missing:
            flash("Please fill in all required fields: " + ", ".join(missing))
            return render_template("form.html", data=request.form, today=date.today().isoformat())

        # Always force the @stm.com.eg domain - only the part before an "@"
        # (if the person typed one) is kept, so it's impossible to end up
        # with any other domain.
        email_local = values["email"].split("@")[0].strip()
        values["email"] = f"{email_local}@stm.com.eg"

        date_input = request.form.get("date", "").strip()
        if date_input:
            try:
                date_obj = datetime.strptime(date_input, "%Y-%m-%d")
            except ValueError:
                flash("Invalid date")
                return render_template("form.html", data=request.form, today=date.today().isoformat())
        else:
            date_obj = datetime.today()

        fill_data = dict(values)
        fill_data["date_obj"] = date_obj
        fill_data["color"] = request.form.get("color", "").strip() or "BLACK"
        fill_data["storage"] = request.form.get("storage", "").strip() or "512SSD"
        fill_data["ram"] = request.form.get("ram", "").strip() or "24GB"
        fill_data["company"] = request.form.get("company", "").strip() or "اس تي ام للاستثمار"

        # Stored on disk under an internal, collision-proof name (so two
        # people named the same thing - or the same person generated
        # twice - never overwrite each other's file or corrupt an older
        # history entry). The name a person actually sees when they
        # download is computed separately in display_filename(), with no
        # id or timestamp in it.
        internal_name = f"{uuid.uuid4().hex}.docx"
        output_path = GENERATED_DIR / internal_name

        if not TEMPLATE_PATH.exists():
            flash("Template.docx is missing on the server.")
            return render_template("form.html", data=request.form, today=date.today().isoformat())

        fill_document(TEMPLATE_PATH, output_path, fill_data)

        record = Handover(
            name=values["name"], department=values["department"], role=values["role"],
            mobile=values["mobile"], email=values["email"], code=values["code"],
            govid=values["govid"], serial=values["serial"], model=values["model"], cpu=values["cpu"],
            handover_date=f"{date_obj.day}/{date_obj.month}/{date_obj.year}",
            filename=internal_name,
            created_by=session.get("display_name", "-"),
        )
        db.session.add(record)
        db.session.commit()

        return redirect(url_for("done", record_id=record.id))

    return render_template("form.html", data={}, today=date.today().isoformat())


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
        as_attachment=True, download_name=display_filename(record.name),
    )


@app.route("/history")
@login_required
def history():
    q = request.args.get("q", "").strip()
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
    records = query.limit(300).all()
    return render_template("history.html", records=records, q=q)


if __name__ == "__main__":
    debug = os.environ.get("FLASK_DEBUG", "1") == "1"
    app.run(debug=debug, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
