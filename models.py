#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
models.py
=========
What the app stores, and nothing else: two tables, and the small
migration that keeps an older database readable.

Split out of app.py so "what is a Handover" is answerable without
reading past thirty routes, and so the domain modules can be handed rows
to work on without importing the web application. The SQLAlchemy object
is created unbound here and attached to the app in app.py, which is what
lets this module be imported on its own.
"""

import json
from datetime import datetime

from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import inspect, text

from templates import TEMPLATES

db = SQLAlchemy()


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


def create_all_and_migrate():
    """Create the tables, then add any column an older database is
    missing. Called once from app.py inside an application context.

    SQLite's ALTER TABLE only supports adding columns, which is all this
    has ever needed - no row is rewritten and nothing is dropped, so
    running it against an up-to-date database does nothing at all.
    """
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
