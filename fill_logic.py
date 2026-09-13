#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fill_logic.py
=============
The document-filling engine for every handover template. Started as a
single hardcoded function for the laptop-handover document; now holds one
fill function per template plus a TEMPLATES registry that the web app
uses to build the "pick a document" screen and each one's own form.

Every fill function edits only the specific data runs inside its
template's XML (never the fixed Arabic/English labels around them), so
the original formatting, fonts, RTL layout and table structure are
preserved exactly - only the data changes. It also auto-detects whether
each value is Arabic or Latin script and fixes that run's right-to-left
flag to match (this is what fixes the original "/" and Arabic-label-
gluing bugs - see fill_handover_script.md in the project for the story).
"""

import copy
import re
from pathlib import Path

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Emu
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE_DIR = Path(__file__).resolve().parent
DOC_TEMPLATES_DIR = BASE_DIR / "doc_templates"

WEEKDAYS_AR = ["الاثنين", "الثلاثاء", "الاربعاء", "الخميس", "الجمعة", "السبت", "الاحد"]
MONTHS_AR = [
    "يناير", "فبراير", "مارس", "ابريل", "مايو", "يونيو",
    "يوليو", "اغسطس", "سبتمبر", "اكتوبر", "نوفمبر", "ديسمبر",
]

ARABIC_RE = re.compile(r"[؀-ۿ]")

DEFAULT_COMPANY = "اس تي ام للاستثمار"


# ----------------------------------------------------------------------
# Low-level run helpers (shared by every fill function)
# ----------------------------------------------------------------------

def has_arabic(text):
    return bool(ARABIC_RE.search(text))


def ensure_bidi(p_el):
    """Make sure a paragraph's <w:pPr> carries <w:bidi/>, adding it (in
    the right schema position, before jc/rPr) if it's missing.

    Inspecting a document the user actually generated and flagged found
    the real root cause of the "mobile number before its label" bug:
    that paragraph was right-aligned (<w:jc w:val="right"/>) but had NO
    <w:bidi/> - so its base paragraph direction was left-to-right even
    though it's full of Arabic text. With an LTR base direction, the
    Unicode bidi algorithm treats the Arabic label as the *first*
    logical segment (placed at the left of the line) and the embedded
    LTR mobile-number run as the *second* (placed at the right, i.e.
    right up against the right-aligned margin) - so reading right to
    left, as any Arabic reader does, the number is exactly what's
    encountered first. Word reproduced this; LibreOffice happened not
    to. Most label/value cells across these templates turned out to
    share this same jc="right"-without-bidi combination (evidently how
    the original source documents were authored) - so it isn't unique to
    the mobile field, which is why it needs to be fixed everywhere at
    once rather than patched for one field. See set_value(), which calls
    this automatically for any paragraph containing Arabic text."""
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_el.insert(0, pPr)
    if pPr.find(qn("w:bidi")) is None:
        pPr.insert(0, OxmlElement("w:bidi"))


def set_value(run, text):
    """Set a run's text, and fix its right-to-left flag to match the
    script of the new text (Arabic vs Latin/digits).

    Defensive: a handful of source runs in the original templates had no
    <w:rPr> at all, which meant they silently inherited the document's
    docDefaults font size (12pt) instead of the 9pt used everywhere else
    in these tables - that mismatch is what caused values like a long
    job title to wrap awkwardly instead of fitting the column. If a run
    is missing rPr, give it one at the same 9pt size the rest of the
    table uses instead of leaving it to chance.

    Also defensive: some of the newer source documents have leftover
    placeholder runs colored white-on-white (a cursor-position spacer the
    original author never meant to be read) - writing a real value into
    one of those would silently vanish. Any run this function fills in
    is meant to be visible, so an explicit white/near-white w:color is
    dropped (falling back to the document's normal black) rather than
    trusted.

    A user reported a real (MS Word) document where a mobile number ended
    up displayed on the wrong side of its Arabic label - reading
    right-to-left, the digits came before "رقم الموبيل/" instead of
    after it. Inspecting the actual generated file found the paragraph
    was missing <w:bidi/> (see ensure_bidi() below, which this function
    now calls automatically) - that was the real cause. On top of that
    fix, two more defensive changes close every other gap the same class
    of bug could hide in: an explicit w:val="0" is now always written
    (never just left "absent") for a non-Arabic run's own w:rtl flag, and
    the text itself is wrapped in Left-to-Right Marks (U+200E, invisible)
    so the run's position is also pinned by the Unicode bidi algorithm
    directly, independent of any single flag or engine's handling of it."""
    is_arabic = has_arabic(text)
    if not is_arabic and text:
        text = "‎" + text + "‎"
    run.text = text
    p_el = run._r.getparent()
    while p_el is not None and p_el.tag != qn("w:p"):
        p_el = p_el.getparent()
    if p_el is not None:
        para_text = "".join(t.text or "" for t in p_el.iter(qn("w:t")))
        if is_arabic or has_arabic(para_text):
            ensure_bidi(p_el)
    rPr = run._r.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        run._r.insert(0, rPr)
        for tag, attrib in (
            ("w:rFonts", {"w:ascii": "Calibri", "w:hAnsi": "Calibri", "w:cs": "Calibri"}),
            ("w:sz", {"w:val": "18"}),
            ("w:szCs", {"w:val": "18"}),
        ):
            el = OxmlElement(tag)
            for k, v in attrib.items():
                el.set(qn(k), v)
            rPr.append(el)
    color_el = rPr.find(qn("w:color"))
    if color_el is not None:
        val = (color_el.get(qn("w:val")) or "").upper()
        if val in ("FFFFFF", "FFF", "FEFEFE") or val.startswith("FFFFFF"):
            rPr.remove(color_el)
    rtl_el = rPr.find(qn("w:rtl"))
    if rtl_el is not None:
        rPr.remove(rtl_el)
    if is_arabic:
        rPr.append(OxmlElement("w:rtl"))
    else:
        # Always write an explicit w:val="0" rather than just leaving the
        # element absent - an absent w:rtl is *supposed* to mean off, but
        # this is exactly the flag implicated in the bug described above,
        # so leave nothing for any renderer to have to infer.
        rtl_el = OxmlElement("w:rtl")
        rtl_el.set(qn("w:val"), "0")
        rPr.append(rtl_el)


def match_size(run, ref_run):
    """Copy just the font size (not color/family) from ref_run's rPr
    onto run's rPr. A few leftover placeholder runs in the newer
    templates carry an oddly large font size relative to their sibling
    runs in the same cell (an editing artifact, not intentional) - large
    enough that the value wraps onto its own line instead of sitting
    inline after its label like every other field. Call this right
    after set_value() when that's the case, passing the cell's label
    run (whose size is the one actually intended) as ref_run."""
    ref_rPr = ref_run._r.find(qn("w:rPr"))
    rPr = run._r.find(qn("w:rPr"))
    if ref_rPr is None or rPr is None:
        return
    for tag in ("w:sz", "w:szCs"):
        ref_el = ref_rPr.find(qn(tag))
        if ref_el is None:
            continue
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), ref_el.get(qn("w:val")))


def match_font(run, ref_run):
    """Copy w:rFonts from ref_run's rPr onto run's rPr. The laptop
    template's device-spec table has every cell in Calibri except the
    Serial cell, which was left in a different font (Aptos Narrow) - a
    leftover from whatever the value was last edited in. Aptos Narrow
    renders noticeably larger than Calibri at the same declared point
    size, which is what actually made the serial number look oversized
    next to its row (match_size() alone matched the point size but not
    the font, so the visual mismatch remained). Call alongside
    match_size() when a run's font family, not just its size, needs to
    follow its row's convention."""
    ref_rPr = ref_run._r.find(qn("w:rPr"))
    rPr = run._r.find(qn("w:rPr"))
    if ref_rPr is None or rPr is None:
        return
    ref_fonts = ref_rPr.find(qn("w:rFonts"))
    if ref_fonts is None:
        return
    fonts = rPr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rPr.append(fonts)
    for attr, val in ref_fonts.attrib.items():
        fonts.set(attr, val)


def clear(run):
    run.text = ""


def append_value(paragraph, label_run_count, text):
    """Write a value into a cell whose label has no reserved value run
    of its own (a bare label like "الوظيفة/" with nothing after it in
    the source template). label_run_count is how many runs the label
    itself is split into in the *pristine* template.

    The first time a given template is filled - including the one-time
    pass that fills it with FILL_* placeholders to produce the shipped
    template in the first place - this appends a new run after the
    label. But a fill function has to keep working on a template that's
    already been through that once (that's exactly what every shipped
    template already has), so if a run beyond the label is already
    there, reuse it instead of stacking a second value run next to the
    first: naively calling paragraph.add_run() unconditionally left the
    template's own "FILL_X" placeholder text sitting right next to the
    real value, e.g. "FILL_ROLESupport Engineer", when a field like this
    was filled directly against a doc_templates/ file rather than the
    example the function was first written against."""
    runs = paragraph.runs
    if len(runs) > label_run_count:
        run = runs[label_run_count]
        for extra in runs[label_run_count + 1:]:
            clear(extra)
    else:
        run = paragraph.add_run()
    set_value(run, text)
    return run


def align_top(cell):
    """Remove a cell's <w:vAlign> override, if any, so it falls back to
    the table's default top alignment. In the printer template the
    header row's cells all agree, but on the value row Brand/Model are
    left at the (default) top alignment while Serial Number/Color were
    left centered - on a row that's taller than one line (left over
    from the header's own multi-line-sized formatting) that makes two
    of the four values visibly sit lower than the other two. Call this
    on the mismatched cells to put the whole row back on one baseline."""
    tcPr = cell._tc.find(qn("w:tcPr"))
    if tcPr is None:
        return
    vAlign = tcPr.find(qn("w:vAlign"))
    if vAlign is not None:
        tcPr.remove(vAlign)


def clear_indent(paragraph):
    """Remove a paragraph's <w:ind> override, if any. One cell in the
    printer template carries a stray negative left/right indent left
    over from earlier editing that its sibling cells in the same row
    don't have - harmless while the cell was empty, but it pushes a
    value written into it onto its own indented line instead of
    sitting inline like every other cell. Call this on a paragraph
    before writing into it when that's the case."""
    pPr = paragraph._p.find(qn("w:pPr"))
    if pPr is None:
        return
    ind = pPr.find(qn("w:ind"))
    if ind is not None:
        pPr.remove(ind)


def strip_hyperlinks(paragraph):
    p = paragraph._p
    for hl in p.findall(qn("w:hyperlink")):
        p.remove(hl)


def fix_table_width(table, target_width_twips):
    """A couple of the newer source tables declare column widths that
    add up to well more than the page's usable width (up to ~11+in of
    columns on a ~7.4in content area) - Word/LibreOffice then silently
    overlaps or clips whichever columns don't fit, which is how a value
    can end up positioned right on top of its neighbour and never
    actually show, even though the text is genuinely there in the XML.
    Force tblLayout=fixed (so the widths set here are honored rather
    than auto-recalculated) and scale every column - both the tblGrid
    and each row's own per-cell tcW, which can otherwise disagree - down
    proportionally so the whole table fits within target_width_twips."""
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblPr.append(layout)
    layout.set(qn("w:type"), "fixed")
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(target_width_twips))

    grid = tbl.find(qn("w:tblGrid"))
    cols = grid.findall(qn("w:gridCol"))
    old_widths = [int(c.get(qn("w:w"))) for c in cols]
    factor = target_width_twips / sum(old_widths)
    new_widths = [round(w * factor) for w in old_widths]

    for col, w in zip(cols, new_widths):
        col.set(qn("w:w"), str(w))
    for row in table.rows:
        for cell, w in zip(row.cells, new_widths):
            tcPr = cell._tc.find(qn("w:tcPr"))
            if tcPr is None:
                continue
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is not None:
                tcW.set(qn("w:w"), str(w))
    return new_widths


def runs_of(table, row, col, para=0):
    """A cell's runs. Some of the newer templates have genuinely empty
    value cells - no <w:r> at all, not even an empty one - since nothing
    was ever typed into them. Create a single empty run first in that
    case, so callers can always safely write into runs_of(...)[0]
    without special-casing "brand new" vs "previously typed-in" cells."""
    p = table.rows[row].cells[col].paragraphs[para]
    if not p.runs:
        p.add_run()
    return p.runs


def date_parts(date_obj):
    date_str = f"{date_obj.day}/{date_obj.month}/{date_obj.year}"
    day_name = WEEKDAYS_AR[date_obj.weekday()]
    month_name = MONTHS_AR[date_obj.month - 1]
    return date_str, day_name, month_name


# ----------------------------------------------------------------------
# 2026-09 template refresh (v4) helpers
#
# Every template's "الموافق/ [date]" cell was rebuilt to one shared
# layout, copied from the arrangement the user made by hand in Word in
# "_استلام لابتوب template 3" and confirmed renders correctly there:
#
#     run[0] = the date value, left-to-right (no w:rtl)
#     run[1] = the Arabic label "الموافق/" (w:rtl, cs font hint)
#     run[2] = a trailing space
#
# in a jc="right" paragraph that deliberately has NO <w:bidi/>. That
# paragraph's base direction being LTR is exactly what makes the
# value-then-label order come out right, so - unlike every other field -
# this one must NOT go through set_value()/ensure_bidi(), which would
# flip the base direction back to RTL and undo the user's arrangement.
# set_date_cell() below writes it verbatim instead.
#
# The other two additions are ordinary label-then-value fields and do go
# through set_value(): a date after "التاريخ/" and "تاريخ الاستلام/" in
# the signature block (those paragraphs already carry <w:bidi/>, which
# is why label-first is correct there), and the new
# "البريد الالكتروني/" row in the employee table.
# ----------------------------------------------------------------------

def set_date_cell(table, date_str, row=0, col=0):
    """Write the date into the normalised الموافق/ cell, preserving the
    hand-verified run order and leaving the paragraph's LTR base
    direction alone. See the note above for why this bypasses
    set_value()."""
    runs = runs_of(table, row, col)
    runs[0].text = date_str + " "
    for extra in runs[3:]:
        clear(extra)


def set_after_label(table, row, col, value):
    """Write `value` into the run that follows a cell's Arabic label,
    finding it by the label's own "/" separator instead of by a fixed
    run index.

    Word re-splits runs freely whenever a document is edited and saved -
    the refreshed templates alone have the same "شهر/" label split five
    different ways ('شهر' + '/ ', 'شهر' + '/' + ' ', and so on), and the
    day cell likewise. Addressing those by index meant every template
    needed its own hand-counted offsets, and any future edit in Word
    silently shifted them. Locating the last run that carries the "/"
    and writing into the one after it (creating it when the label is the
    cell's only run) handles every layout, and keeps handling them after
    the next edit."""
    p = table.rows[row].cells[col].paragraphs[0]
    if not p.runs:
        p.add_run()
    runs = p.runs
    label_end = 0
    for i, r in enumerate(runs):
        if "/" in r.text:
            label_end = i
    if label_end + 1 >= len(runs) or "\n" in (runs[label_end + 1].text or ""):
        # Either the label is the cell's only run, or the run right after
        # it is a <w:br/> holding the cell's blank lines open (several
        # templates end "الكود الوظيفي/" that way). Both need a real run
        # inserted directly after the label - appending to the end of the
        # paragraph instead is what used to drop the employee code two
        # lines below its own label.
        new_r = OxmlElement("w:r")
        src_rPr = runs[label_end]._r.find(qn("w:rPr"))
        if src_rPr is not None:
            new_r.append(copy.deepcopy(src_rPr))
        runs[label_end]._r.addnext(new_r)
        runs = p.runs
    if not runs[label_end].text.endswith(" "):
        value = " " + value
    set_value(runs[label_end + 1], value)
    for extra in runs[label_end + 2:]:
        # Stop at a line break - several cells end with <w:br/> runs that
        # hold the row's blank lines open; clearing those collapses the
        # cell's height.
        if "\n" in (extra.text or ""):
            break
        clear(extra)
    return runs[label_end + 1]


def set_last_slash_run(table, row, col, text):
    """For the handful of cells authored value-first with the "/" at the
    END ("الموظف فى إدارة Facility Management/"), where the value and
    the separator share one run. Rewrites that run and clears whatever
    followed it."""
    p = table.rows[row].cells[col].paragraphs[0]
    if not p.runs:
        p.add_run()
    runs = p.runs
    idx = None
    for i, r in enumerate(runs):
        if "/" in r.text:
            idx = i
    if idx is None:
        idx = len(runs) - 1
    set_value(runs[idx], text)
    for extra in runs[idx + 1:]:
        if "\n" in (extra.text or ""):
            break
        clear(extra)
    return runs[idx]


def find_row(table, needle):
    """Index of the first row whose text contains `needle`, or None."""
    for ri, row in enumerate(table.rows):
        if any(needle in c.text for c in row.cells):
            return ri
    return None


def set_email(table, email):
    """Fill the "البريد الالكتروني/" row added to the employee table.
    The row is a single full-width merged cell holding a label run and
    an empty value run."""
    ri = find_row(table, "البريد الالكتروني")
    if ri is None:
        return False
    runs = runs_of(table, ri, 0)
    if len(runs) < 2:
        table.rows[ri].cells[0].paragraphs[0].add_run()
        runs = runs_of(table, ri, 0)
    set_value(runs[1], " " + email)
    for extra in runs[2:]:
        clear(extra)
    return True


def set_signature_dates(doc, date_str):
    """Fill the date after "التاريخ/" and "تاريخ الاستلام/" in the
    signature block (the last table's last row) - the reference
    template now carries the date there, so every document should."""
    for table in reversed(doc.tables):
        text = " ".join(c.text for row in table.rows for c in row.cells)
        if "التاريخ" not in text or "تاريخ الاستلام" not in text:
            continue
        row = table.rows[-1]
        seen = set()
        for cell in row.cells:
            if cell._tc in seen:
                continue
            seen.add(cell._tc)
            for p in cell.paragraphs:
                runs = p.runs
                label_at = next(
                    (i for i, r in enumerate(runs)
                     if "التاريخ" in r.text or "تاريخ الاستلام" in r.text),
                    None,
                )
                if label_at is None or label_at + 1 >= len(runs):
                    continue
                set_value(runs[label_at + 1], date_str)
                # The reference template carries a real date here, left
                # over from the document it was made from - clear it so
                # the generated document doesn't print both dates.
                for extra in runs[label_at + 2:]:
                    if "\n" in (extra.text or ""):
                        break
                    clear(extra)
                break
        return True
    return False


# ----------------------------------------------------------------------
# Placeholder engine (2026-09 template refresh)
#
# The laptop templates were re-authored in Word so that every value the
# app fills now sits in the document as an explicit "FILL_SOMETHING"
# token (FILL_NAME, FILL_SERIAL, ...) instead of being addressed by
# hand-counted run indexes. That is what broke the old code: every time
# the source .docx was re-saved from Word, the runs were re-split and
# the counted offsets pointed at the wrong text.
#
# Matching on the token itself is immune to that - Word can re-split the
# runs however it likes, the token still says what it is. These helpers
# do the replacement; each fill function below just supplies the mapping.
# ----------------------------------------------------------------------

PLACEHOLDER_RE = re.compile(r"FILL_[A-Z0-9_]+")


def iter_paragraphs(doc):
    """Every paragraph in the document - body, tables (including nested
    tables), and headers/footers - so a placeholder is found wherever the
    person who edited the template happened to put it."""

    def walk_table(table):
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
                for inner in cell.tables:
                    for p in walk_table(inner):
                        yield p

    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for p in walk_table(table):
            yield p
    for section in doc.sections:
        for part in (section.header, section.footer,
                     section.even_page_header, section.even_page_footer,
                     section.first_page_header, section.first_page_footer):
            if part is None:
                continue
            for p in part.paragraphs:
                yield p
            for table in part.tables:
                for pp in walk_table(table):
                    yield pp


def replace_in_paragraph(paragraph, token, value):
    """Replace one token with a value inside a paragraph, even if Word
    has split the token across several runs.

    The value always ends up in a run of its own so that set_value()'s
    right-to-left isolation actually applies to it - writing it back into
    a run that also holds part of a label was how "01211057730" ended up
    printed hard against "رقم الموبيل/" with no gap and, worse, on the
    wrong side of it.

    A separating space is inserted automatically when the template butts
    the token straight up against its label ("FILL_DATE" immediately
    followed by "الموافق /"). Templates are hand-edited in Word, so that
    missing space is a recurring authoring slip rather than a one-off,
    and it is cheaper to absorb it here than to police every .docx."""
    replaced = False
    while True:
        runs = paragraph.runs
        texts = [r.text or "" for r in runs]
        full = "".join(texts)
        start = full.find(token)
        if start == -1:
            return replaced
        end = start + len(token)

        before = full[start - 1] if start else ""
        after = full[end] if end < len(full) else ""
        pad_left = " " if value and before and not before.isspace() else ""
        pad_right = " " if value and after and not after.isspace() else ""

        pos = 0
        target = None
        for run, text in zip(runs, texts):
            r_start, r_end = pos, pos + len(text)
            pos = r_end
            if r_end <= start or r_start >= end:
                continue
            local_s = max(start, r_start) - r_start
            local_e = min(end, r_end) - r_start
            if target is None:
                target = (run, text[:local_s], text[local_e:])
            else:
                # A continuation run: drop just the matched slice, keep
                # whatever else it held.
                run.text = text[:local_s] + text[local_e:]

        run, prefix, suffix = target
        prefix += pad_left
        suffix = pad_right + suffix
        # Snapshot the run's formatting before set_value() edits it, so
        # the prefix/suffix runs keep the label's own direction and font
        # rather than inheriting the value's.
        pristine = copy.deepcopy(run._r)

        from docx.text.run import Run
        if prefix:
            run.text = prefix
            value_el = run_like(pristine, "")
            run._r.addnext(value_el)
        else:
            value_el = run._r
            run.text = ""
        set_value(Run(value_el, paragraph), value)
        if suffix:
            value_el.addnext(run_like(pristine, suffix))
        replaced = True


def fill_placeholders(doc, mapping):
    """Apply a {token: value} mapping to the whole document, then blank
    out any token the mapping didn't cover - a leftover "FILL_XYZ"
    printed on a signed handover document is worse than an empty slot.
    Longer tokens are done first so FILL_DAYNAME is never eaten by
    FILL_DAY."""
    tokens = sorted(mapping, key=len, reverse=True)
    for paragraph in iter_paragraphs(doc):
        if "FILL_" not in paragraph.text:
            continue
        for token in tokens:
            if token in paragraph.text:
                replace_in_paragraph(paragraph, token, mapping[token] or "")
        for leftover in set(PLACEHOLDER_RE.findall(paragraph.text)):
            replace_in_paragraph(paragraph, leftover, "")


def spec_table(doc, *required_headers):
    """Find a device-spec table by its header row rather than by index,
    so inserting or removing a table elsewhere in the document doesn't
    silently shift which table gets written to."""
    wanted = [h.lower() for h in required_headers]
    for table in doc.tables:
        if len(table.rows) < 2:
            continue
        header = " ".join(c.text.strip().lower() for c in table.rows[0].cells)
        if all(w in header for w in wanted):
            return table
    return None


def set_spec_cell(table, column_header, value):
    """Write a value into a spec table's data row by column *name*.

    Some columns in the refreshed laptop-handover template hold a literal
    default (Color "BLACK", Hard "512SSD", RAM "24GB") rather than a
    FILL_ token, because they rarely change. They are still real form
    fields though, so whatever the person typed has to win - otherwise
    a laptop issued with 1TB would print as 512SSD."""
    if table is None or not value:
        return False
    header_cells = table.rows[0].cells
    for idx, cell in enumerate(header_cells):
        if cell.text.strip().lower() == column_header.strip().lower():
            data_cell = table.rows[1].cells[idx]
            paragraph = data_cell.paragraphs[0]
            if not paragraph.runs:
                return False
            set_value(paragraph.runs[0], value)
            for extra in paragraph.runs[1:]:
                clear(extra)
            for extra_p in data_cell.paragraphs[1:]:
                for run in extra_p.runs:
                    clear(run)
            return True
    return False


# ----------------------------------------------------------------------
# Employee-details table tidy-up
#
# The "بيانات الموظف المستلم" table in the laptop templates was edited by
# hand in Word over a long time and its six label/value cells drifted
# badly out of step with each other: one cell puts the value *before* its
# label and the rest put it after, two cells start with stray tab runs
# that shove the value halfway across the column, the labels sit at
# different font sizes (some bold, some not), and nothing lines up
# vertically because the cells disagree on alignment. Filling values into
# that as-is just prints a mess, however correct the values are.
#
# So rather than write into the drift, these helpers rebuild each
# label/value pair in that one table to a single canonical shape -
# "الاسم بالكامل/ <value>", right-aligned in a proper RTL paragraph, label
# and value sharing the table's dominant font - which is what makes the
# two columns line up into two clean vertical rails.
# ----------------------------------------------------------------------

EMPLOYEE_TABLE_MARKER = "الاسم بالكامل"

# Word writes these inside a run; they are structure, not text, so they
# are dropped when a run is rebuilt from a reference run's formatting.
_RUN_CONTENT_TAGS = ("w:t", "w:br", "w:tab", "w:cr", "w:noBreakHyphen")


def arabic_weight(text):
    return sum(1 for ch in (text or "") if ARABIC_RE.match(ch))


def dominant_format(table):
    """The font size and family most of this table's Arabic labels are
    already using, weighted by how much text each run actually carries -
    so one stray oversized leftover run can't win the vote. Rebuilt runs
    adopt it, which is what removes the size/bold jumble between cells."""
    from collections import Counter
    sizes, fonts = Counter(), Counter()
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    weight = arabic_weight(run.text)
                    if not weight:
                        continue
                    rPr = run._r.find(qn("w:rPr"))
                    if rPr is None:
                        continue
                    sz = rPr.find(qn("w:sz"))
                    if sz is not None and sz.get(qn("w:val")):
                        sizes[sz.get(qn("w:val"))] += weight
                    rf = rPr.find(qn("w:rFonts"))
                    if rf is not None:
                        fonts[(rf.get(qn("w:ascii")), rf.get(qn("w:hAnsi")),
                               rf.get(qn("w:cs")))] += weight
    return (sizes.most_common(1)[0][0] if sizes else None,
            fonts.most_common(1)[0][0] if fonts else None)


def apply_format(run, size_val, font_names):
    """Force one run onto the table's dominant size/family, and clear the
    bold that a few values carried and their own labels didn't."""
    rPr = run._r.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        run._r.insert(0, rPr)
    if size_val:
        for tag in ("w:sz", "w:szCs"):
            el = rPr.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                rPr.append(el)
            el.set(qn("w:val"), size_val)
    if font_names and any(font_names):
        rf = rPr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rPr.insert(0, rf)
        for attr, val in zip(("w:ascii", "w:hAnsi", "w:cs"), font_names):
            if val:
                rf.set(qn(attr), val)
    for tag in ("w:b", "w:bCs"):
        el = rPr.find(qn(tag))
        if el is not None:
            rPr.remove(el)


def run_like(ref_el, text):
    """A new run element carrying ref_el's formatting and nothing but the
    given text."""
    el = copy.deepcopy(ref_el)
    for child in list(el):
        if child.tag in [qn(t) for t in _RUN_CONTENT_TAGS]:
            el.remove(child)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    el.append(t)
    return el


def line_segments(paragraph):
    """Split a paragraph into the lines Word draws inside it (runs can
    contain <w:br/>, which python-docx surfaces as "\n"), as lists of
    (run, text-slice) pairs. Each line is one label/value pair in these
    cells - the mobile cell, for instance, holds mobile and email as two
    lines of a single paragraph."""
    segments = [[]]
    for run in paragraph.runs:
        parts = (run.text or "").split("\n")
        for i, part in enumerate(parts):
            if i:
                segments.append([])
            if part:
                segments[-1].append((run, part))
    return segments


def tidy_label(text):
    """Collapse the tabs and doubled spaces out of a label, and move a
    leading "/" to the end where it belongs ("/Email" -> "Email/") - the
    template has it leading because the value used to be typed first."""
    label = re.sub(r"\s+", " ", text).strip()
    if label.startswith("/") and not label.endswith("/"):
        label = label[1:].strip() + "/"
    return label


# A label on its own, with the placeholder that belongs after it. These
# templates get re-authored in Word regularly, and a placeholder is one
# stray keystroke away from being deleted - which is exactly what
# happened to FILL_NAME in the "الاسم بالكامل/" cell: the label survived,
# the token did not, and the document printed with a blank where the
# employee's name should be. Nothing errored, because FILL_NAME still
# existed further down the page in the declaration, so a whole-document
# check for missing tokens saw nothing wrong.
#
# So a label found sitting alone, with no value of any kind after it, is
# filled from this map as if its token were still there. The "with no
# value after it" part is what makes this safe: a line that already has
# content is never touched, so a label whose value is deliberately
# literal text in the template (the company name in the laptop-handover
# document) is left exactly as the author wrote it.
LABEL_TOKENS = {
    "الاسم بالكامل/": "FILL_NAME",
    "أقر انا/": "FILL_NAME",
    "الوظيفة/": "FILL_ROLE",
    "الى إدارة/": "FILL_DEPARTMENT",
    "تابع الى إدارة/": "FILL_DEPARTMENT",
    "الموظف فى إدارة/": "FILL_DEPARTMENT",
    "رقم الموبيل/": "FILL_MOBILE",
    "الكود الوظيفي/": "FILL_CODE",
    "البريد الالكتروني/": "FILL_EMAIL",
    "البريد الالكترونى/": "FILL_EMAIL",
    "Email/": "FILL_EMAIL",
    "رقم قومي/": "FILL_GOVID",
    "رقم قومى/": "FILL_GOVID",
    "التاريخ/": "FILL_DATE",
    "تاريخ الاستلام/": "FILL_DATE",
}


def normalise_label(label):
    """Fold the spelling variants that make two identical-looking Arabic
    labels compare unequal: the alef forms (أ إ آ ٱ), the two yeh forms,
    the two teh-marbuta/heh forms, tatweel padding, and diacritics."""
    for group, canonical in (("أإآٱ", "ا"), ("ىي", "ي"), ("ةه", "ه")):
        for ch in group:
            label = label.replace(ch, canonical)
    label = re.sub(r"[\u0640\u064b-\u0652]", "", label)
    return re.sub(r"\s+", " ", label).strip()


NORMALISED_LABEL_TOKENS = {normalise_label(k): v for k, v in LABEL_TOKENS.items()}


def token_for_orphan_label(text):
    """The placeholder a bare label should have been followed by, or None
    if this isn't a bare label."""
    label = tidy_label(text)
    if not label:
        return None
    return NORMALISED_LABEL_TOKENS.get(normalise_label(label))


def rebuild_pairs(paragraph, mapping, size_val, font_names, trim_keep=False):
    """Rewrite every "label + FILL_TOKEN" line in this paragraph as a
    clean "label/ value", leaving any line that has no token (or more
    than one) exactly as the template author wrote it.

    trim_keep additionally strips the leading and trailing spaces off
    those left-alone lines. Those spaces were the template's way of
    positioning text by hand, and in a right-aligned right-to-left line
    the leading ones sit at the *right* - which is why the literal
    company name in the handover declaration still hung 15pt short of
    the national-ID rail above it after everything else had been lined
    up. Only pass it when the caller is setting the alignment itself,
    so hand-positioned lines elsewhere keep their spacing."""
    segments = line_segments(paragraph)
    if not any(PLACEHOLDER_RE.search("".join(t for _, t in seg))
               or token_for_orphan_label("".join(t for _, t in seg))
               for seg in segments):
        return False

    plans = []
    for seg in segments:
        full = "".join(t for _, t in seg)
        tokens = PLACEHOLDER_RE.findall(full)
        if len(tokens) == 1 and seg:
            token = tokens[0]
            ref = max(seg, key=lambda pair: arabic_weight(pair[1]))[0]
            plans.append(("pair", ref._r, tidy_label(full.replace(token, "")),
                          mapping.get(token, "") or ""))
        elif not tokens and seg and token_for_orphan_label(full):
            # A label whose placeholder was lost in a Word edit - see
            # LABEL_TOKENS above. Fill it as though the token were there.
            token = token_for_orphan_label(full)
            ref = max(seg, key=lambda pair: arabic_weight(pair[1]))[0]
            plans.append(("pair", ref._r, tidy_label(full),
                          mapping.get(token, "") or ""))
        else:
            pieces = [(run._r, text) for run, text in seg]
            if trim_keep:
                # Walk in from both ends: the padding is usually spread
                # over several runs ("    ", " ", "لدى شركة", ...), so
                # stripping only the first and last leaves a space still
                # sitting at the line's right edge.
                for i in range(len(pieces)):
                    ref, text = pieces[i]
                    pieces[i] = (ref, text.lstrip())
                    if pieces[i][1]:
                        break
                for i in range(len(pieces) - 1, -1, -1):
                    ref, text = pieces[i]
                    pieces[i] = (ref, text.rstrip())
                    if pieces[i][1]:
                        break
            plans.append(("keep", pieces))

    while len(plans) > 1 and plans[-1][0] == "keep" and not any(
            text.strip() for _, text in plans[-1][1]):
        plans.pop()

    p_el = paragraph._p
    for run_el in p_el.findall(qn("w:r")):
        p_el.remove(run_el)

    from docx.text.run import Run
    first = True
    for plan in plans:
        if not first:
            br_run = OxmlElement("w:r")
            br_run.append(OxmlElement("w:br"))
            p_el.append(br_run)
        first = False
        if plan[0] == "keep":
            for ref_el, text in plan[1]:
                p_el.append(run_like(ref_el, text))
            continue
        _, ref_el, label, value = plan
        if label:
            el = run_like(ref_el, label + " ")
            p_el.append(el)
            run = Run(el, paragraph)
            set_value(run, label + " ")
            apply_format(run, size_val, font_names)
        el = run_like(ref_el, "")
        p_el.append(el)
        run = Run(el, paragraph)
        set_value(run, value)
        apply_format(run, size_val, font_names)
    return True


def table_with(doc, *markers):
    """The first table containing all of the given marker strings."""
    for table in doc.tables:
        text = "\n".join(c.text for row in table.rows for c in row.cells)
        if all(m in text for m in markers):
            return table
    return None


def align_to_start(paragraph):
    """Put an Arabic paragraph hard against the right edge of its cell.

    Deliberately by *removing* <w:jc>, not by setting it to "right".
    Inside a right-to-left paragraph, Word and LibreOffice both read
    w:jc="right" as "the end of the line", which in RTL is the LEFT of
    the cell - measuring the rendered PDF showed exactly that: setting
    RIGHT pushed every label to the left edge, while no w:jc at all
    (the default, "start") put them where they belong. The paragraphs in
    this table inherit no alignment from their style, so dropping the
    override is safe and is also what Word itself writes for a normal
    RTL paragraph."""
    ensure_bidi(paragraph._p)
    pPr = paragraph._p.find(qn("w:pPr"))
    if pPr is None:
        return
    jc = pPr.find(qn("w:jc"))
    if jc is not None:
        pPr.remove(jc)


def tidy_table(table, mapping, realign="none"):
    """Rebuild a table's label/value pairs into one canonical shape, and
    optionally pin them to their cell's right edge so the labels form
    straight vertical rails instead of sitting at whatever indent each
    cell happened to be left on.

    realign:
      "all"   - every paragraph in the table. The employee-details table,
                where the whole point is two clean rails. It has to
                include the paragraphs holding no placeholder at all (the
                company name is literal text in the laptop-handover
                template) - leaving those on their original alignment is
                what left "لدى شركة/" sticking out past the labels above
                and below it.
      "pairs" - only the paragraphs that actually carry a placeholder.
                Used for the declaration and signature blocks, where the
                filled lines drift out of line with the fixed headings
                above them ("التاريخ/ 13/9/2026" sat well left of the
                "التوقيع/" it belongs under) but the surrounding prose -
                the accessories list, in particular - is laid out
                deliberately and must not be touched.
      "none"  - order and spacing only. The date header row is evenly
                spaced across the page by design, so it is left alone."""
    if table is None:
        return False
    size_val, font_names = dominant_format(table)
    for row in table.rows:
        for cell in row.cells:
            if realign == "all":
                align_top(cell)
            for paragraph in cell.paragraphs:
                rebuilt = rebuild_pairs(paragraph, mapping, size_val, font_names,
                                        trim_keep=realign != "none")
                if realign == "all" or (realign == "pairs" and rebuilt):
                    clear_indent(paragraph)
                    align_to_start(paragraph)
                elif rebuilt:
                    ensure_bidi(paragraph._p)
    return True


def trim_blank_lines(cell):
    """Drop the empty leading/trailing lines inside a cell.

    The old-device model cell in the replacement template starts with a
    line break before its placeholder, so the model printed on the second
    line of the row while the serial, colour, hard, RAM and CPU beside it
    printed on the first - the value was right, it just sat a line lower
    than everything else. Trailing breaks likewise pad a row taller than
    its content needs."""
    for paragraph in cell.paragraphs:
        nodes = [n for n in paragraph._p.iter() if n.tag in (qn("w:br"), qn("w:t"))]
        filled = [i for i, n in enumerate(nodes)
                  if n.tag == qn("w:t") and (n.text or "").strip()]
        if not filled:
            continue
        first, last = filled[0], filled[-1]
        for i, node in enumerate(nodes):
            if node.tag == qn("w:br") and (i < first or i > last):
                node.getparent().remove(node)


# Every placeholder that names a property of the device being handed
# over, as opposed to the employee receiving it. Used to recognise the
# device-spec table in a document without having to know which columns
# that particular document happens to have.
DEVICE_TOKENS = (
    "FILL_MODEL", "FILL_SERIAL", "FILL_BRAND", "FILL_COLOR", "FILL_CAPACITY",
    "FILL_SIM_NUMBER", "FILL_STORAGE", "FILL_HARD", "FILL_RAM", "FILL_CPU",
    "FILL_OLD_MODEL", "FILL_OLD_SERIAL", "FILL_OLD_COLOR", "FILL_OLD_STORAGE",
    "FILL_OLD_HARD", "FILL_OLD_RAM", "FILL_OLD_CPU",
    "FILL_NEW_MODEL", "FILL_NEW_SERIAL", "FILL_NEW_COLOR", "FILL_NEW_STORAGE",
    "FILL_NEW_HARD", "FILL_NEW_RAM", "FILL_NEW_CPU",
)


def find_spec_tables(doc):
    """The device-spec tables, recognised by the device placeholders they
    carry rather than by their column headings.

    Must be called BEFORE filling, while the tokens are still there. The
    earlier version looked for a table with both "Serial" and "CPU"
    headings, which quietly skipped the documents that have neither - the
    mouse receipt only has Model and Brand, so its row never got levelled
    and the model printed bold beside a non-bold brand."""
    return [table for table in doc.tables
            if len(table.rows) >= 2
            and any(token in "\n".join(c.text for row in table.rows for c in row.cells)
                    for token in DEVICE_TOKENS)]


def drop_empty_columns(table):
    """Remove a spec table's filler columns.

    These tables were drawn with six columns whatever the device needs,
    so a hard disk's four real columns sit beside two empty ones. The
    empty columns still claim their width, which squeezes the real ones -
    a 14-digit SIM number was wrapping onto a second line inside a column
    narrower than the blank one next to it. A column is only removed when
    both its heading and its value are blank, so nothing that carries
    content is ever dropped."""
    grid = table._tbl.find(qn("w:tblGrid"))
    if grid is None or len(table.rows) < 2:
        return
    columns = grid.findall(qn("w:gridCol"))
    blank = [i for i in range(len(columns))
             if i < len(table.rows[0].cells) and i < len(table.rows[1].cells)
             and not table.rows[0].cells[i].text.strip()
             and not table.rows[1].cells[i].text.strip()]
    if not blank or len(blank) == len(columns):
        return
    # A merged cell spans columns, so removing one would corrupt the row.
    # Note a <w:gridSpan> of 1 is not a merge - these templates carry them
    # on ordinary cells, and treating their mere presence as a merge made
    # this bail out on every table it was meant to fix.
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.find(qn("w:tcPr"))
            if tcPr is None:
                continue
            span = tcPr.find(qn("w:gridSpan"))
            if span is not None and int(span.get(qn("w:val")) or 1) > 1:
                return
    spare = sum(int(columns[i].get(qn("w:w")) or 0) for i in blank)
    keep = [i for i in range(len(columns)) if i not in blank]
    kept_total = sum(int(columns[i].get(qn("w:w")) or 0) for i in keep) or 1
    for i in keep:
        width = int(columns[i].get(qn("w:w")) or 0)
        columns[i].set(qn("w:w"), str(width + round(spare * width / kept_total)))
    for i in reversed(blank):
        grid.remove(columns[i])
        for row in table.rows:
            cells = row.cells
            if i < len(cells):
                cells[i]._tc.getparent().remove(cells[i]._tc)
    # The per-cell widths have to agree with the new grid.
    widths = [int(c.get(qn("w:w")) or 0) for c in grid.findall(qn("w:gridCol"))]
    for row in table.rows:
        for cell, width in zip(row.cells, widths):
            tcPr = cell._tc.find(qn("w:tcPr"))
            if tcPr is None:
                continue
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is not None:
                tcW.set(qn("w:w"), str(width))


def fit_to_page(doc, tables):
    """Shrink any of these tables that is declared wider than the page.

    Nearly every device-spec table in these documents claims more column
    width than the page has (one asked for 16425 twips of columns on a
    10656-twip page). Word and LibreOffice deal with that by overlapping
    or clipping whichever columns don't fit, which is how a header ended
    up printed on top of its neighbour and a Color value fell off the
    right edge entirely. Scaling the columns down proportionally keeps
    the intended relative widths and puts every column back on the page.

    Tables that already fit are left completely alone."""
    section = doc.sections[0]
    usable = Emu(section.page_width - section.left_margin - section.right_margin).twips
    for table in tables:
        grid = table._tbl.find(qn("w:tblGrid"))
        if grid is None:
            continue
        columns = grid.findall(qn("w:gridCol"))
        total = sum(int(c.get(qn("w:w")) or 0) for c in columns)
        if total > usable:
            fix_table_width(table, usable)


def level_spec_rows(tables):
    """Put every value in a device-spec row on one font size and family,
    and drop the bold a few of them carry and their neighbours don't.

    Each cell gets one vote for the size it is already using and the
    majority wins, so this corrects the odd one out without inventing a
    size the document never had. A tie goes to the smaller size, which
    can only ever help a value fit its column."""
    from collections import Counter
    for table in tables:
        sizes, fonts, runs = Counter(), Counter(), []
        for cell in table.rows[1].cells:
            cell_sizes, cell_fonts = set(), set()
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    if not (run.text or "").strip():
                        continue
                    runs.append(run)
                    rPr = run._r.find(qn("w:rPr"))
                    if rPr is None:
                        continue
                    sz = rPr.find(qn("w:sz"))
                    if sz is not None and sz.get(qn("w:val")):
                        cell_sizes.add(sz.get(qn("w:val")))
                    rf = rPr.find(qn("w:rFonts"))
                    if rf is not None:
                        cell_fonts.add((rf.get(qn("w:ascii")), rf.get(qn("w:hAnsi")),
                                        rf.get(qn("w:cs"))))
            # One vote per cell, so a long value can't outweigh a short one.
            for value in cell_sizes:
                sizes[value] += 1
            for value in cell_fonts:
                fonts[value] += 1

        def winner(counter, numeric=False):
            if not counter:
                return None
            key = (lambda kv: (-kv[1], int(kv[0]))) if numeric else (lambda kv: -kv[1])
            return sorted(counter.items(), key=key)[0][0]

        size_val = winner(sizes, numeric=True)
        font_names = winner(fonts)
        for run in runs:
            apply_format(run, size_val, font_names)
        for idx, cell in enumerate(table.rows[1].cells):
            align_top(cell)
            trim_blank_lines(cell)
            # Centre every heading and every value, so each value sits
            # under the column it belongs to. The templates had these
            # cells on whatever alignment the author last dragged them
            # to - some centred, some not, some simply inheriting - so
            # "Logitech" printed off to the left of the "Brand" above it
            # while "SN-77421" hugged the right of its own column.
            for target in (table.rows[0].cells[idx], cell):
                for paragraph in target.paragraphs:
                    # The headings carry leftover hand-positioning
                    # indents (one column's heading was pushed 2126
                    # twips in, the next 3126) - centring a heading
                    # inside an indent centres it in the wrong place,
                    # which is why "Serial Number" floated right of the
                    # value underneath it.
                    clear_indent(paragraph)
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


def tidy_document(doc, mapping):
    """Apply the tidy-up to the four tables that hold label/value pairs:
    the date header, the employee details, the declaration the employee
    signs, and the signature block."""
    tidy_table(table_with(doc, EMPLOYEE_TABLE_MARKER), mapping, realign="all")
    tidy_table(table_with(doc, "الموافق", "شهر"), mapping)
    tidy_table(table_with(doc, "أقر انا"), mapping, realign="pairs")
    tidy_table(table_with(doc, "التوقيع"), mapping, realign="pairs")


# ----------------------------------------------------------------------
# Laptop handover
# ----------------------------------------------------------------------

def fill_laptop_handover(template_path, output_path, data):
    """Laptop handover.

    Template refreshed again 2026-09 (v5): the source .docx is now
    placeholder-driven (FILL_NAME, FILL_SERIAL, ...), so this function no
    longer walks tables by row/column/run index. That indexing is exactly
    what broke when the document was last re-saved from Word - the labels
    and values were re-split into different runs and the hand-counted
    offsets wrote into the wrong places (or raised IndexError).

    Color / Hard / RAM are not tokens in this template - they carry the
    common defaults as literal text - so they're written by column name
    afterwards from whatever the form actually submitted."""
    d = docx.Document(str(template_path))
    date_str, day_name, month_name = date_parts(data["date_obj"])

    email = (data.get("email") or "").strip()
    company = data.get("company") or DEFAULT_COMPANY

    values = {
        "FILL_DATE": date_str,
        "FILL_MONTH": month_name,
        "FILL_DAY": day_name,
        "FILL_DAYNAME": day_name,
        "FILL_NAME": data["name"],
        "FILL_ROLE": data["role"],
        "FILL_DEPARTMENT": data["department"],
        "FILL_MOBILE": data["mobile"],
        "FILL_EMAIL": email,
        "FILL_CODE": data["code"],
        "FILL_GOVID": data["govid"],
        "FILL_COMPANY": company,
        "FILL_MODEL": data["model"],
        "FILL_SERIAL": data["serial"],
        "FILL_CPU": data["cpu"],
        # Present in some revisions of the template, literal defaults in
        # others - mapped either way so both revisions fill correctly.
        "FILL_COLOR": data.get("color", ""),
        "FILL_HARD": data.get("storage", ""),
        "FILL_STORAGE": data.get("storage", ""),
        "FILL_RAM": data.get("ram", ""),
    }
    specs = find_spec_tables(d)
    tidy_document(d, values)
    fill_placeholders(d, values)
    level_spec_rows(specs)
    for spec in specs:
        drop_empty_columns(spec)
    fit_to_page(d, specs)

    specs = spec_table(d, "serial", "cpu")
    set_spec_cell(specs, "Color", data.get("color", ""))
    set_spec_cell(specs, "Hard", data.get("storage", ""))
    set_spec_cell(specs, "RAM", data.get("ram", ""))

    # The accessories note under the signature block lists the laptop
    # charger, bag and mouse. Older source documents hardcoded one real
    # mouse brand/model there, which would print on every document
    # regardless of what was issued - strip it if it comes back.
    for p in iter_paragraphs(d):
        for run in p.runs:
            if "Logitech" in run.text or "M171" in run.text:
                run.text = ""

    output_path.parent.mkdir(parents=True, exist_ok=True)
    d.save(str(output_path))
    return output_path


# ----------------------------------------------------------------------
# Laptop replacement
# ----------------------------------------------------------------------

def fill_laptop_replacement(template_path, output_path, data):
    """Laptop replacement (old device returned + new device issued).

    Same story as the handover above - the refreshed source document is
    placeholder-driven, so the old index-based table walking is gone. It
    also quietly fixes a swap that the index version had wrong: in this
    revision the FIRST spec table is the OLD device and the second is the
    NEW one, the opposite of the previous file's order, which is how old
    and new specs ended up printed on each other's rows."""
    d = docx.Document(str(template_path))
    date_str, day_name, month_name = date_parts(data["date_obj"])

    email = (data.get("email") or "").strip()
    company = data.get("company") or DEFAULT_COMPANY

    values = {
        "FILL_DATE": date_str,
        "FILL_MONTH": month_name,
        "FILL_DAYNAME": day_name,
        "FILL_DAY": day_name,
        "FILL_NAME": data["name"],
        "FILL_ROLE": data["role"],
        "FILL_DEPARTMENT": data["department"],
        "FILL_MOBILE": data["mobile"],
        "FILL_EMAIL": email,
        "FILL_CODE": data["code"],
        "FILL_GOVID": data["govid"],
        "FILL_COMPANY": company,
        "FILL_OLD_MODEL": data["old_model"],
        "FILL_OLD_SERIAL": data["old_serial"],
        "FILL_OLD_COLOR": data["old_color"],
        "FILL_OLD_STORAGE": data["old_storage"],
        "FILL_OLD_HARD": data["old_storage"],
        "FILL_OLD_RAM": data["old_ram"],
        "FILL_OLD_CPU": data["old_cpu"],
        "FILL_NEW_MODEL": data["new_model"],
        "FILL_NEW_SERIAL": data["new_serial"],
        "FILL_NEW_COLOR": data["new_color"],
        "FILL_NEW_STORAGE": data["new_storage"],
        "FILL_NEW_HARD": data["new_storage"],
        "FILL_NEW_RAM": data["new_ram"],
        "FILL_NEW_CPU": data["new_cpu"],
    }
    specs = find_spec_tables(d)
    tidy_document(d, values)
    fill_placeholders(d, values)
    level_spec_rows(specs)
    for spec in specs:
        drop_empty_columns(spec)
    fit_to_page(d, specs)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    d.save(str(output_path))
    return output_path


# ----------------------------------------------------------------------
# The eight accessory documents (mouse, keyboard & mouse, screen, router,
# headset, printer, flash drive, hard disk)
#
# These were converted to the same placeholder form as the laptop
# templates, and they all share one six-table skeleton - title, date
# header, employee details, device spec, declaration, signatures. That
# makes one fill function enough for all eight: every difference between
# them (a router has a SIM number, a flash drive has a capacity, a mouse
# has neither) is expressed by which placeholders that document contains,
# so the mapping below can simply offer all of them and let each template
# take the ones it has. Adding a ninth accessory document later needs a
# registry entry and a .docx, and no new code at all.
# ----------------------------------------------------------------------

def fill_accessory_handover(template_path, output_path, data):
    """Fill any of the accessory handover documents."""
    d = docx.Document(str(template_path))
    date_str, day_name, month_name = date_parts(data["date_obj"])

    values = {
        "FILL_DATE": date_str,
        "FILL_MONTH": month_name,
        "FILL_DAY": day_name,
        "FILL_DAYNAME": day_name,
        "FILL_NAME": data["name"],
        "FILL_ROLE": data["role"],
        "FILL_DEPARTMENT": data["department"],
        "FILL_MOBILE": data["mobile"],
        "FILL_EMAIL": (data.get("email") or "").strip(),
        "FILL_CODE": data["code"],
        "FILL_GOVID": data["govid"],
        "FILL_COMPANY": data.get("company") or DEFAULT_COMPANY,
    }
    # Device fields, offered for every template; each document fills in
    # only the ones its own spec table actually asks for.
    for key, token in (
        ("model", "FILL_MODEL"), ("serial", "FILL_SERIAL"),
        ("brand", "FILL_BRAND"), ("color", "FILL_COLOR"),
        ("capacity", "FILL_CAPACITY"), ("sim_number", "FILL_SIM_NUMBER"),
        ("storage", "FILL_STORAGE"), ("ram", "FILL_RAM"), ("cpu", "FILL_CPU"),
    ):
        if key in data:
            values[token] = data[key]

    specs = find_spec_tables(d)
    tidy_document(d, values)
    fill_placeholders(d, values)
    level_spec_rows(specs)
    for spec in specs:
        drop_empty_columns(spec)
    fit_to_page(d, specs)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    d.save(str(output_path))
    return output_path


# ----------------------------------------------------------------------
# Template registry - the single source of truth for the "pick a
# document" screen and each one's own form fields.
# ----------------------------------------------------------------------

# Field specs shared by (almost) every template.
#
# "pattern" is a regex (matched with re.fullmatch) used for both the
# server-side check in app.py and the HTML5 pattern= attribute in
# form.html (so most people see the browser's inline error before ever
# submitting); "pattern_msg" is the human-readable explanation shown by
# both when a value doesn't match.
_NAME = {"key": "name", "label": "Full name", "placeholder": "e.g. Ahmed Ali Mohamed", "required": True}
_DEPARTMENT = {"key": "department", "label": "Department", "placeholder": "IT, Finance, Sales…", "required": True}
_ROLE = {"key": "role", "label": "Position", "placeholder": "Software Engineer", "required": True}
_MOBILE = {
    "key": "mobile", "label": "Mobile", "placeholder": "01xxxxxxxxx", "required": True,
    "inputmode": "numeric", "maxlength": 11,
    "pattern": r"01[0125][0-9]{8}",
    "hint": "11 digits",
    "pattern_msg": "Enter an 11-digit Egyptian mobile number starting with 010, 011, 012, or 015 (e.g. 01012345678).",
}
_EMAIL = {"key": "email", "label": "Email", "required": True, "suffix": "@stm.com.eg"}
_CODE = {"key": "code", "label": "Employee code", "required": True}
_GOVID = {
    "key": "govid", "label": "National ID", "required": True, "maxlength": 14,
    "inputmode": "numeric", "placeholder": "", "wide": True, "hint": "14 digits",
    "pattern": r"[0-9]{14}",
    "pattern_msg": "National ID must be exactly 14 digits.",
}

EMPLOYEE_FIELDS_WITH_EMAIL = [_NAME, _DEPARTMENT, _ROLE, _MOBILE, _EMAIL, _CODE, _GOVID]
EMPLOYEE_FIELDS_NO_EMAIL = [_NAME, _DEPARTMENT, _ROLE, _MOBILE, _CODE, _GOVID]

TEMPLATES = {
    "laptop_handover": {
        "label": "Laptop handover",
        "label_ar": "محضر تسليم لاب توب",
        "group": "Computers",
        "category": "computer",
        "short": "Laptop",
        "doc_file": "laptop_handover.docx",
        "filename_prefix": "استلام لابتوب",
        "fill": fill_laptop_handover,
        "employee_fields": EMPLOYEE_FIELDS_WITH_EMAIL,
        "device_title": "Laptop details",
        "device_fields": [
            {"key": "model", "label": "Laptop model", "placeholder": "E14", "required": True},
            {"key": "serial", "label": "Serial number", "placeholder": "PF5K...", "required": True},
            {"key": "cpu", "label": "CPU", "placeholder": "I5 / I7", "required": True},
        ],
        "extra_fields": [
            {"key": "color", "label": "Color", "placeholder": "BLACK", "default": "BLACK"},
            {"key": "storage", "label": "Storage", "placeholder": "512SSD", "default": "512SSD"},
            {"key": "ram", "label": "RAM", "placeholder": "24GB", "default": "24GB"},
            {"key": "company", "label": "Company", "placeholder": DEFAULT_COMPANY, "default": DEFAULT_COMPANY},
        ],
    },
    "laptop_replacement": {
        "label": "Laptop replacement",
        "label_ar": "محضر استبدال جهاز كمبيوتر",
        "group": "Computers",
        "category": "computer",
        "short": "Replacement",
        "doc_file": "laptop_replacement.docx",
        "filename_prefix": "استبدال لابتوب",
        "fill": fill_laptop_replacement,
        "employee_fields": EMPLOYEE_FIELDS_WITH_EMAIL,
        "device_title": "Old device (being returned)",
        "device_fields": [
            {"key": "old_model", "label": "Old model", "placeholder": "E14", "required": True},
            {"key": "old_serial", "label": "Old serial number", "placeholder": "PF5K...", "required": True},
            {"key": "old_color", "label": "Old color", "placeholder": "BLACK", "required": True},
            {"key": "old_storage", "label": "Old storage", "placeholder": "512SSD", "required": True},
            {"key": "old_ram", "label": "Old RAM", "placeholder": "24GB", "required": True},
            {"key": "old_cpu", "label": "Old CPU", "placeholder": "I5 / I7", "required": True},
        ],
        "device_title_2": "New device (being issued)",
        "device_fields_2": [
            {"key": "new_model", "label": "New model", "placeholder": "E14", "required": True},
            {"key": "new_serial", "label": "New serial number", "placeholder": "PF5K...", "required": True},
            {"key": "new_color", "label": "New color", "placeholder": "BLACK", "required": True},
            {"key": "new_storage", "label": "New storage", "placeholder": "512SSD", "required": True},
            {"key": "new_ram", "label": "New RAM", "placeholder": "24GB", "required": True},
            {"key": "new_cpu", "label": "New CPU", "placeholder": "I5 / I7", "required": True},
        ],
        "extra_fields": [
            {"key": "company", "label": "Company", "placeholder": DEFAULT_COMPANY, "default": DEFAULT_COMPANY},
        ],
    },
    "keyboard_mouse_handover": {
        "label": "Keyboard & mouse handover",
        "label_ar": "محضر تسليم كيبورد وماوس",
        "group": "Peripherals",
        "category": "peripheral",
        "short": "Keyboard & mouse",
        "doc_file": "keyboard_mouse_handover.docx",
        "filename_prefix": "استلام كيبورد وماوس",
        "fill": fill_accessory_handover,
        "employee_fields": EMPLOYEE_FIELDS_WITH_EMAIL,
        "device_title": "Keyboard & mouse details",
        "device_fields": [
            {"key": "brand", "label": "Brand", "placeholder": "Logitech", "required": True},
            {"key": "serial", "label": "Serial number", "placeholder": "KM123456", "required": True},
            {"key": "model", "label": "Model", "placeholder": "MK270", "required": True},
            {"key": "color", "label": "Color", "placeholder": "Black", "required": True},
        ],
        "extra_fields": [
            {"key": "company", "label": "Company", "placeholder": DEFAULT_COMPANY, "default": DEFAULT_COMPANY},
        ],
    },
    "mouse_receipt": {
        "label": "Mouse handover",
        "label_ar": "محضر استلام ماوس",
        "group": "Peripherals",
        "category": "peripheral",
        "short": "Mouse",
        "doc_file": "mouse_receipt.docx",
        "filename_prefix": "استلام ماوس",
        "fill": fill_accessory_handover,
        "employee_fields": EMPLOYEE_FIELDS_WITH_EMAIL,
        "device_title": "Mouse details",
        "device_fields": [
            {"key": "model", "label": "Model", "placeholder": "M171", "required": True},
            {"key": "brand", "label": "Brand", "placeholder": "Logitech", "required": True},
        ],
        "extra_fields": [
            {"key": "company", "label": "Company", "placeholder": DEFAULT_COMPANY, "default": DEFAULT_COMPANY},
        ],
    },
    "screen_handover": {
        "label": "Screen handover",
        "label_ar": "محضر تسليم شاشة",
        "group": "Peripherals",
        "category": "peripheral",
        "short": "Screen",
        "doc_file": "screen_handover.docx",
        "filename_prefix": "تسليم شاشة",
        "fill": fill_accessory_handover,
        "employee_fields": EMPLOYEE_FIELDS_WITH_EMAIL,
        "device_title": "Screen details",
        "device_fields": [
            {"key": "model", "label": "Screen model", "placeholder": "SE2725HMc (Dell 27 Monitor)", "required": True},
            {"key": "serial", "label": "Serial number", "placeholder": "CN-...", "required": True},
            {"key": "color", "label": "Color", "placeholder": "Black", "required": True},
        ],
        "extra_fields": [
            {"key": "company", "label": "Company", "placeholder": DEFAULT_COMPANY, "default": DEFAULT_COMPANY},
        ],
    },
    "router_handover": {
        "label": "Router handover",
        "label_ar": "محضر تسليم راوتر بشريحة",
        "group": "Storage & network",
        "category": "network",
        "short": "Router",
        "doc_file": "router_handover.docx",
        "filename_prefix": "استلام راوتر",
        "fill": fill_accessory_handover,
        "employee_fields": EMPLOYEE_FIELDS_WITH_EMAIL,
        "device_title": "Router details",
        "device_fields": [
            {"key": "brand", "label": "Brand", "placeholder": "Huawei", "required": True},
            {"key": "model", "label": "Model", "placeholder": "B535", "required": True},
            {"key": "serial", "label": "Serial number", "placeholder": "RT123456", "required": True},
            {"key": "sim_number", "label": "Data SIM number", "placeholder": "01234567890123", "required": True},
            {"key": "color", "label": "Color", "placeholder": "White", "required": True},
        ],
        "extra_fields": [
            {"key": "company", "label": "Company", "placeholder": DEFAULT_COMPANY, "default": DEFAULT_COMPANY},
        ],
    },
    "headset_handover": {
        "label": "Headset handover",
        "label_ar": "محضر تسليم سماعة",
        "group": "Peripherals",
        "category": "peripheral",
        "short": "Headset",
        "doc_file": "headset_handover.docx",
        "filename_prefix": "استلام سماعة",
        "fill": fill_accessory_handover,
        "employee_fields": EMPLOYEE_FIELDS_WITH_EMAIL,
        "device_title": "Headset details",
        "device_fields": [
            {"key": "model", "label": "Headset model", "placeholder": "HS-200", "required": True},
            {"key": "serial", "label": "Serial number", "placeholder": "HSN123456", "required": True},
            {"key": "color", "label": "Color", "placeholder": "Black", "required": True},
        ],
        "extra_fields": [
            {"key": "company", "label": "Company", "placeholder": DEFAULT_COMPANY, "default": DEFAULT_COMPANY},
        ],
    },
    "printer_handover": {
        "label": "Printer handover",
        "label_ar": "محضر تسليم طابعة",
        "group": "Peripherals",
        "category": "peripheral",
        "short": "Printer",
        "doc_file": "printer_handover.docx",
        "filename_prefix": "استلام طابعة",
        "fill": fill_accessory_handover,
        "employee_fields": EMPLOYEE_FIELDS_WITH_EMAIL,
        "device_title": "Printer details",
        "device_fields": [
            {"key": "brand", "label": "Brand", "placeholder": "HP", "required": True},
            {"key": "model", "label": "Model", "placeholder": "LaserJet-M111", "required": True},
            {"key": "serial", "label": "Serial number", "placeholder": "SN-PRT123456", "required": True},
            {"key": "color", "label": "Color", "placeholder": "White", "required": True},
        ],
        "extra_fields": [
            {"key": "company", "label": "Company", "placeholder": DEFAULT_COMPANY, "default": DEFAULT_COMPANY},
        ],
    },
    "flash_handover": {
        "label": "Flash drive handover",
        "label_ar": "محضر تسليم فلاشة",
        "group": "Storage & network",
        "category": "storage",
        "short": "Flash drive",
        "doc_file": "flash_handover.docx",
        "filename_prefix": "استلام فلاشة",
        "fill": fill_accessory_handover,
        "employee_fields": EMPLOYEE_FIELDS_WITH_EMAIL,
        "device_title": "Flash drive details",
        "device_fields": [
            {"key": "brand", "label": "Brand", "placeholder": "Kingston", "required": True},
            {"key": "capacity", "label": "Capacity", "placeholder": "32 GB", "required": True},
            {"key": "model", "label": "Model", "placeholder": "DTSE9/32GB", "required": True},
            {"key": "color", "label": "Color", "placeholder": "Black", "required": True},
        ],
        "extra_fields": [
            {"key": "company", "label": "Company", "placeholder": DEFAULT_COMPANY, "default": DEFAULT_COMPANY},
        ],
    },
    "hard_handover": {
        "label": "Hard disk handover",
        "label_ar": "محضر تسليم هارد",
        "group": "Storage & network",
        "category": "storage",
        "short": "Hard disk",
        "doc_file": "hard_handover.docx",
        "filename_prefix": "استلام هارد",
        "fill": fill_accessory_handover,
        "employee_fields": EMPLOYEE_FIELDS_WITH_EMAIL,
        "device_title": "Hard disk details",
        "device_fields": [
            {"key": "brand", "label": "Brand", "placeholder": "Seagate", "required": True},
            {"key": "model", "label": "Model", "placeholder": "Expansion", "required": True},
            {"key": "serial", "label": "Serial number", "placeholder": "HD-SN123456", "required": True},
            {"key": "color", "label": "Color", "placeholder": "Black", "required": True},
        ],
        "extra_fields": [
            {"key": "company", "label": "Company", "placeholder": DEFAULT_COMPANY, "default": DEFAULT_COMPANY},
        ],
    },
}


# The picker shows the ten documents in these groups, in this order, and
# the history table tints each document's chip by its category. Purely
# presentational - the fill functions never read them - but they live
# beside the labels so there is one place to describe a document type.
TEMPLATE_GROUPS = ("Computers", "Peripherals", "Storage & network")


def all_fields(template_id):
    """Every field key this template's form should collect (employee +
    device + device_2 if any), in order."""
    spec = TEMPLATES[template_id]
    fields = list(spec["employee_fields"]) + list(spec["device_fields"])
    if "device_fields_2" in spec:
        fields += list(spec["device_fields_2"])
    return fields


def required_field_keys(template_id):
    return [f["key"] for f in all_fields(template_id) if f.get("required")]


def template_path(template_id):
    return DOC_TEMPLATES_DIR / TEMPLATES[template_id]["doc_file"]
