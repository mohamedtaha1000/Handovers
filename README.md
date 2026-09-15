# STM Handover Documents

An internal Flask web application for generating STM's handover and
receipt Word documents from a browser form, replacing manual editing of
`.docx` files. Every generated document is logged to a searchable,
filterable history with export and print support.

## Contents

- [Overview](#overview)
- [Requirements](#requirements)
- [Running locally](#running-locally)
- [Configuration](#configuration)
- [Project structure](#project-structure)
- [Features](#features)
- [Deployment](#deployment)
- [Security notes](#security-notes)
- [Adding a new document type](#adding-a-new-document-type)
- [Possible future enhancements](#possible-future-enhancements)

## Overview

Users log in with a shared team password, choose a document type, fill
in a form, and download the completed `.docx`. Supported document types:

| Document type              | Description                              |
|------------------------------|-------------------------------------------|
| Laptop handover              | Issuing a laptop to an employee            |
| Laptop replacement           | Swapping an employee's laptop for a new one (records both the returned and issued device) |
| Keyboard & mouse handover     | Issuing a keyboard-and-mouse combo kit to an employee |
| Mouse handover                | Issuing a mouse (on its own) to an employee |
| Screen handover               | Issuing a monitor to an employee           |
| Router handover                | Issuing a SIM/data router to an employee   |
| Headset handover               | Issuing a headset to an employee           |
| Printer handover                | Issuing a printer to an employee          |
| Flash drive handover             | Issuing a USB flash drive to an employee |
| Hard disk handover               | Issuing an external hard disk to an employee |

Each document type is defined in a single registry
(`fill_logic.py::TEMPLATES`), which drives the home page, the form for
each document type, and the underlying `.docx` template — adding a new
type does not require changing the Flask routes or HTML pages (see
[Adding a new document type](#adding-a-new-document-type)).

## Requirements

- Python 3.10+
- The packages listed in `requirements.txt`

## Running locally

```bash
cd webapp
pip install -r requirements.txt
```

Copy `.env.example` to `.env` (if not already present) and set two
values:

```
SECRET_KEY=<a random long string>
TEAM_PASSWORD=<a password to share with your team>
```

Generate a secure `SECRET_KEY` with:

```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

Start the app:

```bash
python app.py
```

Open **http://127.0.0.1:5000**, log in with any display name and the
team password, and generate a document. `.env` is loaded automatically
on startup — no environment variables need to be set manually in the
shell.

On first run this creates:
- `instance/handovers.db` — a local SQLite database of generated documents
- `generated/` — the generated `.docx` files, named
  `<Name> - <Document type> - <Date>.docx` (e.g.
  `Yasmin Mohamed - Laptop handover - 2026-09-10.docx`); if the same
  person generates the same document type again on the same day, `(2)`,
  `(3)`, etc. is appended instead of overwriting the earlier file

Both directories are excluded from version control (see `.gitignore`).

## Configuration

| Variable        | Required | Description                                            |
|------------------|----------|----------------------------------------------------------|
| `SECRET_KEY`      | Yes      | Signs Flask session cookies. Any long random string.       |
| `TEAM_PASSWORD`   | Yes      | Shared password required to log in.                        |
| `DATABASE_URL`    | No       | Overrides the default SQLite database (e.g. a Postgres connection string) with no code changes required. |
| `PORT`            | No       | Port to listen on when run directly with `python app.py` (default `5000`). |
| `FLASK_DEBUG`      | No       | Set to `0` to disable Flask's debug/auto-reload mode.      |
| `HANDOVER_NOTIFY_TO`   | No   | Who the "this has been handed over" email is addressed to. Unset means the draft opens with an empty To line. |
| `HANDOVER_NOTIFY_NAME` | No   | How that email greets its recipient (default `Eng. Hegazy`). |
| `HANDOVER_EMS_TO`      | No   | Who the leaver's EMS deactivation email is addressed to. |
| `HANDOVER_LEAVER_TO`   | No   | Who the leaver's resignation email is addressed to. |
| `HANDOVER_REGISTER_PATH` | No | Where the laptop register `.xlsx` is written (default `laptop_register.xlsx` beside the app). Rebuilt automatically on every change. |

The app refuses to start with the placeholder `TEAM_PASSWORD` value —
it must be set before running.

## Project structure

```
app.py                  Flask application: routes, authentication, database
fill_logic.py            Document-filling engine and the TEMPLATES registry
                         (one fill function per document type, plus the
                         field list each type's form is generated from)
doc_templates/            One placeholder Word template per document type:
                             Laptop Handover Template.docx
                             Laptop Replacement Template.docx
                             Keyboard and Mouse Handover Template.docx
                             Mouse Receipt Template.docx
                             Screen Handover Template.docx
                             Router Handover Template.docx
                             Headset Handover Template.docx
                             Printer Handover Template.docx
                             Flash Drive Handover Template.docx
                             Hard Disk Handover Template.docx
templates/               HTML pages (login, picker, form, done, history)
static/style.css          Stylesheet
static/stm-logo.png        Logo, used in the header and favicon
requirements.txt          Python dependencies
Procfile                  Start command for hosting platforms (gunicorn)
.env                       Local secrets (not committed to version control)
.env.example                Blank reference copy of .env, safe to commit
```

## Features

### Document generation

Each document type has its own form, generated from its field list in
`fill_logic.py`. Required fields are validated server-side; mobile
number and National ID fields are additionally checked against a
regular expression (11-digit Egyptian mobile number, 14-digit national
ID) before a document is generated, with the specific problem shown
inline if a value doesn't match.

### History

Every generated document is recorded with the submitting user's name,
department, position, national ID (masked to the last 4 digits in list
views), document type, and the date and time it was generated. The
History page supports:

- **Search** — matches name, department, or position
- **Filters** — by document type, and by a from/to date range (based on
  generation date)
- **Pagination** — 50 records per page
- **Print** — a print-friendly view of the current filtered results,
  with navigation and controls hidden
- **Export to Excel** — downloads the current filtered results as an
  `.xlsx` file (national ID masked the same way as on screen)
- **Delete** — permanently removes a record and its generated `.docx`
  file, after a confirmation prompt. Any logged-in user can delete any
  record.

Search, filters, and pagination are reflected in the URL, so a filtered
view can be bookmarked or shared.

## Deployment

The application is a standard Flask app and can be deployed to any
platform that runs Python (Render, Railway, PythonAnywhere, an internal
server, etc.). Example using **Render**:

1. Push this folder to its own **private** GitHub repository —
   `doc_templates/` contains STM's internal document layouts and should
   not be made public.
2. On [render.com](https://render.com), create a **New → Web Service**
   and connect the repository.
3. Set the build and start commands:
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn app:app`
4. Set the environment variables `SECRET_KEY` and `TEAM_PASSWORD` (see
   [Configuration](#configuration)).
5. Deploy. Render provides a URL (e.g.
   `https://stm-handover.onrender.com`) to share with the team, along
   with the `TEAM_PASSWORD`.

**Note on free hosting tiers:** platforms like Render's free tier use an
*ephemeral* filesystem — the SQLite database and generated files can be
lost on restart or redeploy. For reliable long-term history, either:

- attach a persistent disk (available on paid plans), or
- point `DATABASE_URL` at a managed database (e.g. Render's free
  Postgres tier) — no code changes required.

Railway and PythonAnywhere follow a similar process: connect the
repository, set the two required environment variables, and point the
start command at `gunicorn app:app`.

## Security notes

This application collects national ID numbers and other personal
employee data. Before deploying it somewhere reachable outside STM's
internal network, review the following:

- **Authentication** is a single shared password for the whole team —
  it gates every page, but does not distinguish between individual
  users beyond the display name they type at login.
- **National IDs are masked** in list views (last 4 digits only). The
  full number is still present in the generated `.docx` file and
  briefly on the confirmation page immediately after generation, since
  it is required in the document itself.
- **Change `SECRET_KEY` and `TEAM_PASSWORD`** from any placeholder
  values before deploying.
- Consider checking with STM's IT/security function before hosting real
  employee national IDs anywhere reachable from outside the office
  network, even behind a password. At minimum, use a host that provides
  HTTPS by default, and do not share the URL or password outside the
  team.
- If access is only needed from the office, hosting on an internal
  server reachable only over the company network/VPN removes the public
  exposure entirely.

## Adding a new document type

1. Obtain a real, filled-in example of the new document.
2. Sanitize it into a placeholder template (`FILL_*` placeholders
   replacing real values) and add it to `doc_templates/`.
3. Write a corresponding fill function in `fill_logic.py`.
4. Add an entry to the `TEMPLATES` dict describing the document's label,
   template file, fill function, and field list.

No other file needs to change — the home page and the document's form
are generated from the registry.

## Possible future enhancements

- Individual per-user logins in place of the single shared password
- Rate limiting, CSRF tokens, and other production-hardening measures
- Restricting who can delete a given history entry to its creator
