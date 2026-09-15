# -*- coding: utf-8 -*-
"""
leaver_email.py
===============
The two messages that go out when someone leaves, built from the same
records the handover documents were generated from.

Both are delivered the same way as the handover notification - a mailto:
link that opens a new Outlook message for the person to read and send -
so this module only has to produce a subject and a plain-text body.

The wording is the team's own, kept as it was given. Two small things
were normalised: a stray double space after "Email:", and the missing
blank line under "Dear Team," in the resignation note. The department in
the resignation sentence is the employee's own rather than a fixed
"Development", since the same message is used for every department.

Everything is typed on the page rather than looked up: someone can leave
without ever having had a document generated for them, so requiring a
record on file would make the buttons useless exactly when they matter.

TOKENS is what lets the two links stay live as the boxes are typed in.
Each link is rendered once with a token where each value goes, and the
browser swaps the typed values in - so the message wording lives here,
in Python, and the page only does substitution. A link that reached
Outlook still carrying a token would be a bug, so the page substitutes
every token on every keystroke, empty values included.
"""

# What the page asks for, in the order it asks.
FORM_FIELDS = (
    {"key": "name", "label": "Full name", "placeholder": "Abdelrahman Ashraf Sarour",
     "wide": True},
    {"key": "department", "label": "Department", "placeholder": "Technical Office"},
    {"key": "email", "label": "Email", "placeholder": "asarour@stm.com.eg"},
    {"key": "computer_name", "label": "Computer name", "placeholder": "STM-LT-0142"},
    {"key": "serial", "label": "Serial number", "placeholder": "PK5PJ7T"},
)

FIELD_KEYS = tuple(f["key"] for f in FORM_FIELDS)
TOKENS = {key: f"__{key.upper()}__" for key in FIELD_KEYS}

# The department is named twice in the resignation note - once in the
# sentence, once in the details - and the sentence has to lose its whole
# clause when there is no department rather than read "from the
# department". So the clause gets a token of its own, and this is the
# shape the page fills in: its wording stays here rather than in the
# browser.
CLAUSE_TOKEN = "__DEPT_CLAUSE__"
CLAUSE_TEMPLATE = f" from the {TOKENS['department']} department"

# The five lines both messages share, in the order the team writes them.
DETAIL_LINES = (
    ("Name", "name"),
    ("Department", "department"),
    ("Email", "email"),
    ("Computer Name", "computer_name"),
    ("Serial Number", "serial"),
)


def detail_block(person):
    """The bulleted employee details. Every line is present even when its
    value is missing: these go to a team that reads the same five lines
    every time, and a gap is a clearer prompt to fill something in than a
    silently absent row."""
    return [f"* {label}: {(person.get(key) or '').strip()}"
            for label, key in DETAIL_LINES]


def dept_clause(department):
    """" from the Technical Office department", or nothing at all."""
    value = (department or "").strip()
    return CLAUSE_TEMPLATE.replace(TOKENS["department"], value) if value else ""


def ems_subject(person):
    return f"EMS deactivation — {person.get('name', '').strip()}".strip(" —")


def ems_body(person):
    return "\n".join([
        "Dear Team,",
        "",
        "Kindly provide the FortiClient deactivation code.",
        "",
        "Employee Details:",
        "",
        *detail_block(person),
        "",
        "Your support is highly appreciated.",
    ])


def resignation_subject(person):
    return f"Resignation — {person.get('name', '').strip()}".strip(" —")


def resignation_body(person):
    name = (person.get("name") or "").strip()
    where = person.get("dept_clause")
    if where is None:
        where = dept_clause(person.get("department"))
    return "\n".join([
        "Dear Team,",
        "",
        f"I would like to inform you that the employee {name}{where} "
        f"has left the company.",
        "",
        "Employee Details:",
        "",
        *detail_block(person),
    ])


# What each button sends, so the route and the page can loop over one
# list instead of repeating themselves per message.
MESSAGES = (
    {
        "key": "ems",
        "label": "EMS deactivation",
        "hint": "Asks for the FortiClient deactivation code",
        "subject": ems_subject,
        "body": ems_body,
    },
    {
        "key": "resignation",
        "label": "Resignation",
        "hint": "Tells the team the employee has left",
        "subject": resignation_subject,
        "body": resignation_body,
    },
)
