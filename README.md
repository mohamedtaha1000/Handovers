# STM Laptop Handover — web app

A small internal website that fills the "محضر تسليم لاب توب" (laptop
handover) Word document from a web form, instead of your teammate having
to run the Python script from a terminal. Every document it generates is
also saved to a searchable history page.

It's built on the exact same tested document-filling code as
`fill_handover.py` (now living in `fill_logic.py`) — same fixes for the
RTL/slash and Arabic-label-gluing bugs, same formatting preservation.

## ⚠️ Before you deploy this publicly — read this

This form collects **national ID numbers** and other personal employee
data. You told me you want it hosted publicly online (reachable from
anywhere), which is the easiest option but also the one with the most
exposure for this kind of data. A few things I built in to reduce the
risk, and a few you should still decide on:

- **Login is required for every page.** Nobody can see the form, the
  history, or download a file without your team's shared password.
- **National IDs are masked** on the history list (only the last 4 digits
  show). The full number is still in the generated `.docx` file itself and
  briefly on the confirmation page right after creating it — that's
  unavoidable, since the ID has to be in the actual document.
- **Change `TEAM_PASSWORD` and `SECRET_KEY`** before you deploy (see
  below) — the app refuses to run safely with the placeholder values.
- I'd strongly recommend checking with whoever handles IT/security policy
  at STM before putting real employee national IDs on a public host,
  even a password-protected one. If that's not possible, at minimum use a
  host that gives you HTTPS by default (all three suggested below do) and
  don't share the URL or password outside your team.
- If it turns out you only need your teammate(s) to reach it from the
  office, hosting it on an internal server reachable only over your
  company network/VPN instead of the public internet would be safer —
  happy to help set that up instead if you change your mind.

## What's in this folder

```
app.py               the Flask application (routes, login, database)
fill_logic.py         the tested document-filling engine (same as fill_handover.py)
Template.docx         the placeholder-based template (FILL_NAME etc. already removed)
templates/            the HTML pages (login, form, done, history)
static/style.css      styling
static/stm-logo.png    the STM logo (pulled from Template.docx) used in the header/favicon
requirements.txt      Python dependencies
Procfile              tells hosting platforms how to start the app (gunicorn)
.env                   your two secrets go here (see below) - keep this file private
.env.example           a blank reference copy of .env, safe to check into git
```

## Running it on your own computer first

```
cd webapp
pip install -r requirements.txt
```

Open `.env` in Notepad and fill in your two values:

```
SECRET_KEY=<a random long string>
TEAM_PASSWORD=<a password you'll share with your teammate>
```

Generate a good `SECRET_KEY` with:
```
python -c "import secrets; print(secrets.token_hex(32))"
```

Then just run:
```
python app.py
```

The app reads `.env` automatically now, so there's nothing to type in
PowerShell before this — no `$env:` lines needed.

Open **http://127.0.0.1:5000** — log in with any name + the team
password, fill the form, and you'll get a download link. Everything you
generate also shows up on the **History** page.

This creates a local `instance/handovers.db` (SQLite database) and saves
generated files in `generated/`. Both are ignored by git (see
`.gitignore`) so they don't get bundled up if you put this in a repo.

## Deploying it publicly

You need to pick a hosting platform and create an account there yourself
— I can't do that part for you, but here's the easiest path with
**Render** (free tier available, gives you HTTPS automatically):

1. Put this `webapp` folder in its own GitHub repository (private
   repository — don't make it public, since it contains `Template.docx`
   with your company's document layout).
2. Go to [render.com](https://render.com), sign up, and click
   **New → Web Service**, then connect that GitHub repo.
3. Render will detect it's a Python app. Set:
   - **Build command:** `pip install -r requirements.txt`
   - **Start command:** `gunicorn app:app`
4. Under **Environment**, add the environment variables:
   - `SECRET_KEY` = (a random value, see above)
   - `TEAM_PASSWORD` = (your team's password)
5. Deploy. Render gives you a URL like
   `https://stm-handover.onrender.com` — that's what you share with your
   teammate.

**Important limitation on Render's free tier:** its filesystem is
*ephemeral* — the SQLite database and generated files can be wiped
whenever the service restarts or redeploys (this happens on the free
tier after periods of inactivity). That's fine for trying it out, but if
you want the history to reliably persist long-term, either:
  - upgrade to a Render paid plan and attach a **persistent disk**, or
  - use a proper database instead of SQLite — Render's free **Postgres**
    tier works well; set the `DATABASE_URL` environment variable to its
    connection string and the app will use it automatically (no code
    changes needed).

**Railway** and **PythonAnywhere** are two other easy options if you'd
rather not use Render — the steps are very similar (connect a repo or
upload the folder, set the same two environment variables, point the
start command at `gunicorn app:app`).

## Giving your teammate access

Once it's deployed, just share:
1. The URL Render (or whichever host) gives you.
2. The `TEAM_PASSWORD` you set.

They type in their own name when logging in (so the history shows who
generated each document) and the shared password.

## If you want to add more people later without sharing one password

Right now everyone shares one team password. If down the line you want
individual logins (so people can't see each other's password, or so you
can revoke one person's access without changing it for everyone), that's
a straightforward upgrade to the login system — just ask and I can add
it.
