#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
settings.py
===========
Every knob in one place: where files live, who the emails go to, what
the shared password is.

All of it comes from the environment (a .env file beside the app, or
real environment variables when deployed) so a change of owner, address
or folder needs no edit and no redeploy. Import the MODULE, not the
names - the tests point these at a temporary folder by reassigning
them, and that only works if they are read at call time:

    import settings
    settings.GENERATED_DIR / name        # picks up a reassignment
    from settings import GENERATED_DIR   # frozen at import - don't
"""

import os
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

GENERATED_DIR = BASE_DIR / "generated"
INSTANCE_DIR = BASE_DIR / "instance"
DOC_TEMPLATES_DIR = BASE_DIR / "doc_templates"
GENERATED_DIR.mkdir(exist_ok=True)
INSTANCE_DIR.mkdir(exist_ok=True)

DATABASE_URL = os.environ.get("DATABASE_URL", f"sqlite:///{INSTANCE_DIR / 'handovers.db'}")

TEAM_PASSWORD = os.environ.get("TEAM_PASSWORD") or "changeme"
SECRET_KEY = os.environ.get("SECRET_KEY")

# Who the "this has been handed over" email is addressed to.
NOTIFY_TO = os.environ.get("HANDOVER_NOTIFY_TO", "").strip()
NOTIFY_NAME = os.environ.get("HANDOVER_NOTIFY_NAME", "Eng. Hegazy").strip()

# Where the two leaver messages go. Separate from NOTIFY_TO because they
# are a different conversation with a different team; either left unset
# just means that message opens with an empty To line, which is still
# faster than writing it out by hand.
EMS_TO = os.environ.get("HANDOVER_EMS_TO", "").strip()
LEAVER_TO = os.environ.get("HANDOVER_LEAVER_TO", "").strip()

# The laptop register. Beside the app by default, so on a machine where
# the folder is synced it simply appears - no download step, no second
# copy to go stale.
REGISTER_PATH = Path(os.environ.get(
    "HANDOVER_REGISTER_PATH", str(BASE_DIR / "laptop_register.xlsx")))
