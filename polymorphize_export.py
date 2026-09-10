#!/usr/bin/env python3
"""
polymorphize_export — the field mapping document, written back out.

Version 6.6.

The tool reads a mapping workbook and writes a specification. This module
writes the third artefact: a field mapping document in the API COE's own
format, produced from the same run that produced the specification, so the two
agree by construction rather than by an analyst keeping them in step by hand.

Three things it is for
======================

**Documentation.** The specification states what the interface is. The
spreadsheet states it in the format the API COE circulates, which is the
format the people who have to agree it actually read.

**The difference between what was asked for and what was published.** Three
columns to the right of the format's own nine say, for every element the
workbook declared, whether it reached the interface and if not why not. Those
are the rows an analyst has to act on, and they are the reason this is not
merely a format converter.

**Iteration.** The exported workbook is itself a valid input. An analyst
corrects the rows the difference columns flag, feeds the workbook back in, and
the next specification is closer. For that loop to be safe the export has to
carry the whole of the original document, including the rows this tool never
reads, so nothing is quietly lost on each pass.

Round trip fidelity, and its one limit
======================================

Where the source was a field mapping document, every cell of every row is
written back **verbatim** from the row it came from, including header rows,
parameter rows, any Parameter Type the reader ignores, and any column beyond
the nine the format defines. The tool does not rewrite an analyst's Usage or
Schema wording into its own spelling, so a comparison between two rounds shows
only what actually changed. What is added is the three difference columns and
the provenance sheet.

Where the source was a Level format workbook there is nothing to carry, and
the sheet is composed from the tree instead. That case is a **conversion**
rather than a round trip and it has two known losses, both stated on the
provenance sheet: the Remarks and Required in Swagger columns have no source
in the Level format and are left empty, and the header block is supplied from
the constant below rather than from the workbook. A Level sheet that failed to
read is not exported at all, because there is no tree to compose from; the
provenance sheet names it.

The standard header block
=========================

The Level format carries no header rows. The reference field mapping document
carries 220 of them across its 22 sheets, and they are seven distinct headers
repeated: the same block on every operation. So the block is a constant here,
taken from that document.

Writing them from a constant also settles an inconsistency in the source.
There ``x-BDO-Application-Id`` is declared both ``string`` and ``string(10)``,
three headers appear once with a length and a Usage and once with neither, and
three rows carry ``CVV``, ``auxiliaryPan`` and ``auxiliaryExpiry`` in the
Required in Swagger column, which is paste drift from the neighbouring column.
The constant uses the most complete declaration of each.

One sheet per SOR endpoint
==========================

The format has a single SOR API Field Name column, so a Level sheet naming
several SOR endpoints becomes several sheets, one per endpoint, each a
complete document for that endpoint. A field mapping sheet already has one, so
its sheet name is preserved unchanged.

Not on the command line
=======================

By instruction. The export is available in the desktop window and in the batch
runner. ``polymorphize_cli`` says so rather than failing silently.
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import os
import re

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import polymorphize_showcase as show
from polymorphize_workbook import USAGE_CONDITIONAL, USAGE_REQUIRED, text

__version__ = "6.6"

# --------------------------------------------------------------------------- #
# The format
# --------------------------------------------------------------------------- #

#: The nine banner rows, in order, exactly as the reference document writes
#: their labels. ``key`` is the banner key both readers produce.
BANNER_ROWS = (
    ("API Name", "api_name"),
    ("Use Case:", "use_case"),
    ("Service Domain:", "service_domain"),
    ("BQ", "behaviour_qualifier"),
    ("Method", "method"),
    ("Equivalent BIAN API Endpoint:", "bian_endpoint"),
    ("Proposed Reusable API Endpoint:", "business_endpoint"),
    ("SOR Name", "sor_name"),
    ("SOR API Endpoint:", "sor_endpoint"),
)

HEADER_ROW = len(BANNER_ROWS) + 1          # row 10, as in the reference

#: The nine columns of the format, in order. The SOR column's heading carries
#: the endpoint on a second line, which is where the reader looks for it.
FORMAT_COLUMNS = ("Parameter Type", "Reusable API Field Name", "Usage",
                  "Schema", "SOR API Field Name", "Example", "Description",
                  "Remarks", "Required in Swagger")

SOR_COLUMN = 4                              # zero-based, within FORMAT_COLUMNS

#: The three columns this tool adds. None of them is read back in: the reader
#: assigns column roles by heading and ignores a heading it does not know, so
#: an exported workbook re-imports as the format it already was.
DIFF_COLUMNS = ("Publication Status", "Why", "Analyst Action")

#: Written into the Publication Status column when a whole sheet failed.
NOT_GENERATED = "Not generated"

USAGE_WORD = {USAGE_REQUIRED: "Mandatory", USAGE_CONDITIONAL: "Conditional "
                                                              "Mandatory"}

# --------------------------------------------------------------------------- #
# The standard header block
# --------------------------------------------------------------------------- #

#: ``(name, usage, schema, sor field, example, description, remarks,
#: required in swagger)``, from the reference document, most complete
#: declaration of each.
REQUEST_HEADERS = (
    ("Authorization", "Mandatory", "string(2048)", "Authorization",
     "Bearer<Token>",
     "Standard header. Types can be basic, bearer, apikey etc. Consumer will "
     "set the authorization type and credentials. Apigee will handle "
     "validation against an identity service for basic authorization, OAuthV2 "
     "for bearer token authorization and apikeys policy validation.",
     "Apigee Standard authentication", "Yes"),
    ("x-BDO-Client-Request-Id", "Conditional Mandatory", "string(36)", "N/A", "",
     "The unique id (preferable uuidv4) provided by the client for a specific "
     "HTTP request.\nChannel Reference Number. This should be unique value "
     "within the channel. Required if other Otel (open telemetry) field are "
     "not provided.",
     "Apigee Standard authentication", "Yes"),
    ("x-BDO-Client-Request-Trace-Id", "Conditional Mandatory", "string(32)",
     "N/A", "",
     "The unique identifier for an entire trace in OpenTelemetry (OTEL).\nA "
     "trace represents a single request or transaction as it flows through "
     "multiple services in a distributed system.\nOnly Mandatory if Request Id "
     "is not provided",
     "Conditional Mandatory, Only Mandatory if Request Id is not provided",
     "Yes"),
    ("x-BDO-Client-Request-Span-Id", "Conditional Mandatory", "string(16)",
     "N/A", "",
     "Uniquely identifies a single span within a trace. A span represents one "
     "unit of work (e.g. HTTP Request, DB Query)\nOnly Mandatory if Request Id "
     "is not provided",
     "Conditional Mandatory, Only Mandatory if Request Id is not provided",
     "Yes"),
    ("x-BDO-API-Version", "Mandatory", "string(1)", "x-BDO-API-Version", "2.0",
     "Denotes the Major API Version.\nConsuming application will set the "
     "version to invoke for their request.",
     "Apigee Standard authentication", "Yes"),
    ("x-BDO-End-Consumer-IP", "Optional", "string(45)", "N/A", "",
     "Consuming application to pass the end-consumer (like bank operations "
     "terminal) IP Address.",
     "Apigee Standard authentication", "Yes"),
    ("x-BDO-Application-Id", "Mandatory", "string(10)", "N/A", "",
     "Unique identifier for a Consumer application that represents a business "
     "capability or product, regardless of how users access it.\nConsuming "
     "application will be provided with the unique value as part of their "
     "Onboarding Process",
     "To be populated in apigee", "No"),
)

RESPONSE_HEADERS = (
    ("x-BDO-Client-Request-Id", "", "string", "N/A", "",
     "The unique id (preferable uuidv4) provided by the client for a specific "
     "HTTP request.", "Apigee Standard authentication", ""),
    ("x-BDO-Client-Request-Trace-Id", "", "string", "N/A", "",
     "The unique identifier for an entire trace in OpenTelemetry (OTEL).\nA "
     "trace represents a single request or transaction as it flows through "
     "multiple services in a distributed system.", "", ""),
    ("x-BDO-Client-Request-Span-Id", "", "string", "N/A", "",
     "Uniquely identifies a single span within a trace. A span represents one "
     "unit of work \n(e.g.,HTTP Request, a database query).", "", ""),
)

# --------------------------------------------------------------------------- #
# Presentation
# --------------------------------------------------------------------------- #

FONT = "Calibri"
NAVY = "FF1F3864"
THIN = Side(style="thin", color="FFB9C6D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FILL_HEADER = PatternFill("solid", fgColor=NAVY)
FILL_BANNER = PatternFill("solid", fgColor="FFE7EEF7")
FILL_SECTION = PatternFill("solid", fgColor="FFD6E4F0")
FILL_DIFF = PatternFill("solid", fgColor="FFF3F0FA")
FILL_ACTION = PatternFill("solid", fgColor="FFFFF6D6")
FILL_DROPPED = PatternFill("solid", fgColor="FFFDEAEA")

COLUMN_WIDTH = {
    "Parameter Type": 18, "Reusable API Field Name": 56, "Usage": 20,
    "Schema": 16, "SOR API Field Name": 30, "Example": 20,
    "Description": 52, "Remarks": 34, "Required in Swagger": 18,
    "Publication Status": 18, "Why": 58, "Analyst Action": 58,
}

#: Publication Status values that mean the element did not reach the interface.
DROPPED = {show.DISPOSITION_LABEL[show.NO_FIELD],
           show.DISPOSITION_LABEL[show.NOT_IN_SOR],
           show.DISPOSITION_LABEL[show.LOST],
           NOT_GENERATED}

# --------------------------------------------------------------------------- #
# What to say about each element
# --------------------------------------------------------------------------- #

ACTION = {
    show.PUBLISHED: "",
    show.CONTAINER: "",
    show.VARIANT_ONLY: "Nothing, unless every variant should carry it. If so, "
                       "name the SOR field on the variants that leave it "
                       "blank.",
    show.NO_FIELD: "Name the SOR field that supplies this element, or leave "
                   "it as N/A to confirm the System of Record does not have "
                   "it.",
    show.NOT_IN_SOR: "Correct the SOR field name, or correct the SOR API "
                     "Endpoint in the banner. The name given was not found in "
                     "the endpoint the banner states.",
    show.LOST: "Nothing you can do in the workbook. Send this sheet to the "
               "maintainer: the element was mapped and verified and should "
               "have been published.",
}


def _why(row, verified):
    """The sentence for the Why column, specific where it can be."""
    note = show.DISPOSITION_NOTE[row.disposition]
    if row.disposition == show.NOT_IN_SOR:
        declared = next((d for d in row.declared.values() if d), "")
        if declared:
            return ("%s was not found in the SOR endpoint this sheet names."
                    % declared)
        return note
    if row.disposition == show.NO_FIELD and not verified:
        return ("The SOR API Field Name cell is empty or N/A, so the workbook "
                "states the System of Record does not supply this element.")
    if row.disposition == show.VARIANT_ONLY:
        yes = [e for e, (state, _d) in row.per_variant.items() if state == "yes"]
        return ("Supplied by %s and not by the other SOR endpoint(s) of this "
                "operation, so it appears only there."
                % (", ".join(e or "the unnamed endpoint" for e in yes[:3])))
    return note


# --------------------------------------------------------------------------- #
# Composing a row
# --------------------------------------------------------------------------- #


def _schema_word(node):
    """The Schema cell for a node that has no source row to copy."""
    if node.json_type in ("object", "array"):
        return node.json_type
    if node.json_type == "string" and node.fmt == "date":
        return "date"
    if node.json_type == "string" and node.fmt == "date-time":
        return "datetime"
    base = node.json_type or "string"
    return "%s(%d)" % (base, node.max_length) if node.max_length else base


def _usage_word(node):
    return USAGE_WORD.get(node.usage, "Optional")


def _dotted(path):
    return ".".join(path)


class SheetPlan:
    """One output sheet: an operation as served by one SOR endpoint."""

    def __init__(self, title, banner, endpoint, sor_header):
        self.title = title
        self.banner = banner
        self.endpoint = endpoint
        self.sor_header = sor_header
        self.extra_headings = []      # columns beyond the nine, carried
        self.rows = []                # list[dict]
        self.counts = {}
        self.note = ""

    def add(self, cells, *, kind="body", status="", why="", action="",
            extra=()):
        self.rows.append({"cells": list(cells), "kind": kind, "status": status,
                          "why": why, "action": action, "extra": list(extra)})


def _blank_row():
    return [""] * len(FORMAT_COLUMNS)


def _section_row(label):
    row = _blank_row()
    row[0] = label
    return row


def _header_rows(block):
    out = []
    for name, usage, schema, sor, example, desc, remarks, req in block:
        out.append(["Header", name, usage, schema, sor, example, desc,
                    remarks, req])
    return out


# --------------------------------------------------------------------------- #
# Planning a sheet from a generated operation
# --------------------------------------------------------------------------- #


def _carry_index(result):
    """``{source row number: original row}`` for a field mapping sheet."""
    rows = getattr(result, "source_rows", None) or []
    return {i + 1: r for i, r in enumerate(rows)}


def _extra_columns(result):
    """Headings beyond the nine the format defines, with their indices."""
    layout = getattr(result, "mapping_layout", None)
    rows = getattr(result, "source_rows", None) or []
    if layout is None or not rows:
        return []
    known = set(layout.cols.values())
    header = rows[layout.header_row - 1] if layout.header_row <= len(rows) else []
    # The columns this tool added on a previous round are recognised and
    # dropped rather than carried. Carrying them would append three more
    # beside them every time the workbook went round, so a document on its
    # fifth pass would be fifteen columns wider than the format allows.
    mine = {_norm(h) for h in DIFF_COLUMNS}
    out = []
    for i, cell in enumerate(header):
        if i in known:
            continue
        label = text(cell)
        if label and _norm(label) not in mine:
            out.append((i, label))
    return out


def _norm(value):
    return "".join(ch for ch in text(value).lower() if ch.isalnum())


def _cell(row, index):
    if row is None or index is None or index < 0 or index >= len(row):
        return ""
    value = row[index]
    return "" if value is None else value


def _body_cells(node, endpoint, path, source, layout):
    """The nine cells for one element.

    Verbatim from the source row where there is one, so an analyst's own
    wording survives the round trip untouched. Composed from the node only
    where the sheet came from the Level format and there is nothing to copy.
    """
    if source is not None and layout is not None:
        cols = layout.cols
        return [
            _cell(source, cols.get("ptype")) or "Body",
            _cell(source, cols.get("field")) or _dotted(path),
            _cell(source, cols.get("usage")),
            _cell(source, cols.get("dtype")),
            _cell(source, layout.sor_col),
            _cell(source, cols.get("example")),
            _cell(source, cols.get("desc")),
            _cell(source, cols.get("remarks")),
            _cell(source, cols.get("inswagger")),
        ]
    # The workbook's word, never the effective mapping: verification may have
    # cleared the latter, and writing the cleared value back would erase the
    # analyst's request on the next round.
    declared = node.sor_declared.get(endpoint, "") or node.sor.get(endpoint, "")
    return ["Body", _dotted(path), _usage_word(node), _schema_word(node),
            declared or "N/A", node.example or "", node.description or "",
            node.remarks or "", ""]


def _walk(node, endpoint, verification, published, prefix, out,
          source_by_row, layout):
    """Depth first, in the order the workbook declares, root row included."""
    for child in node.children:
        path = prefix + (child.name,)
        row = show.disposition_of(child, [endpoint], verification, published,
                                  path)
        source = source_by_row.get(child.row)
        out.append((child, path, row, source))
        if child.children:
            _walk(child, endpoint, verification, published, path, out,
                  source_by_row, layout)


def plan_sheet(view, endpoint, *, title, verified, ordinal=0, total=1):
    """One :class:`SheetPlan` for one operation and one SOR endpoint."""
    result = view.result
    layout = getattr(result, "mapping_layout", None)
    source_by_row = _carry_index(result)
    published = show.published_leaf_paths(view.spec)
    banner = dict(result.banner or {})
    banner.setdefault("method", view.spec.method.upper())
    banner["sor_endpoint"] = _endpoint_label(result, banner, endpoint,
                                             ordinal, total)
    if not text(banner.get("api_name")):
        # The Level format has no API Name row. Derived from the sheet name
        # rather than left blank, and recorded as derived on the provenance
        # sheet so nobody mistakes it for the analyst's own wording.
        banner["api_name"] = _derive_api_name(result.sheet)
    plan = SheetPlan(title, banner, endpoint,
                     _sor_heading(result, layout, banner, endpoint))
    plan.extra_headings = _extra_columns(result)
    extra_idx = [i for i, _label in plan.extra_headings]

    carried = _carried_non_body(result, layout)
    counts = {}

    for label, section, headers in (
            ("Request", result.request, REQUEST_HEADERS),
            ("Response", result.response, RESPONSE_HEADERS)):
        if plan.rows:
            plan.add(_blank_row(), kind="blank")
        # The header block: carried where the source had one, from the
        # constant where it did not.
        block = carried.get(label)
        if block:
            for source_row, cells in block:
                plan.add(cells, kind="carried",
                         extra=[_cell(source_row, i) for i in extra_idx])
        else:
            plan.add(_section_row("Request Parameter" if label == "Request"
                                  else "Response Header"), kind="section")
            for cells in _header_rows(headers):
                plan.add(cells, kind="header")

        plan.add(_blank_row(), kind="blank")
        plan.add(_section_row("%s Body" % label), kind="section")
        if section is None or section.root is None:
            continue
        rows = []
        _walk(section.root, endpoint, view.verification, published, (),
              rows, source_by_row, layout)
        # The root row itself, which the reference document carries.
        root_source = source_by_row.get(section.root.row)
        plan.add(_body_cells(section.root, endpoint, (section.root.name,),
                             root_source, layout),
                 kind="body", status=show.DISPOSITION_LABEL[show.CONTAINER],
                 why="The message itself.",
                 extra=[_cell(root_source, i) for i in extra_idx])
        for node, path, row, source in rows:
            status = show.DISPOSITION_LABEL[row.disposition]
            counts[status] = counts.get(status, 0) + 1
            plan.add(_body_cells(node, endpoint,
                                 (section.root.name,) + path, source, layout),
                     kind="body", status=status,
                     why=_why(row, verified),
                     action=ACTION.get(row.disposition, ""),
                     extra=[_cell(source, i) for i in extra_idx])

    plan.counts = counts
    return plan


def _endpoint_label(result, banner, endpoint, ordinal=0, total=1):
    """The SOR endpoint the banner should state.

    For a sheet that came from a field mapping document the banner is
    reproduced exactly as written. That matters more than it looks: the banner
    is the definitive statement of the SOR endpoint and the endpoint appended
    to the SOR column header is only a label, so rewriting the banner from the
    column would change which SOR the next round verifies against. On the
    reference document the two differ on one sheet, which is how this was
    found.

    A Level sheet naming several SOR endpoints becomes one sheet per endpoint,
    so there the banner has to state the endpoint of this sheet's column.
    """
    stated = text(banner.get("sor_endpoint", ""))
    if getattr(result, "mapping_layout", None) is not None:
        return stated or endpoint or ""
    # A Level banner naming several endpoints pairs them with the SOR columns
    # in order, so the sheet for column n takes fragment n, keeping the
    # analyst's own method prefix and spelling.
    parts = [p.strip() for p in stated.split(",") if p.strip()]
    if total > 1 and len(parts) == total:
        return parts[ordinal]
    if total > 1 and endpoint:
        return endpoint if endpoint.startswith("/") else "/" + endpoint
    return stated or endpoint or ""


def _derive_api_name(sheet):
    """``Card Transaction_Retrieve`` -> ``Card Transaction Retrieve API``."""
    words = re.sub(r"[_\-]+", " ", text(sheet)).strip()
    words = re.sub(r"\s{2,}", " ", words)
    if not words:
        return ""
    return words if words.lower().endswith("api") else "%s API" % words


def _sor_heading(result, layout, banner, endpoint):
    """The SOR column heading, with the endpoint on its second line."""
    rows = getattr(result, "source_rows", None) or []
    if layout is not None and layout.header_row <= len(rows):
        raw = rows[layout.header_row - 1]
        cell = _cell(raw, layout.sor_col)
        if text(cell):
            return cell                       # verbatim, newline included
    tail = endpoint or text(banner.get("sor_endpoint", ""))
    return "SOR API Field Name\n%s" % tail if tail else "SOR API Field Name"


def _carried_non_body(result, layout):
    """The header and parameter rows of a field mapping sheet, verbatim.

    Returns ``{"Request": [(source row, cells)], "Response": [...]}``. Empty
    where the sheet did not come from that format, which is the signal to use
    the constant block instead.
    """
    out = {}
    rows = getattr(result, "source_rows", None) or []
    if layout is None or not rows:
        return out
    import polymorphize_mapping as mapfmt
    from polymorphize_workbook import key as _key

    where = "Request"
    for r in range(layout.header_row, len(rows)):
        row = rows[r]
        if not any(text(c) for c in row):
            continue
        ptype = layout.cell(row, "ptype")
        pk = _key(ptype)
        field_path = layout.cell(row, "field")
        if pk in ("responseheader", "responseparameter"):
            where = "Response"
        elif pk in mapfmt.SECTION_BY_TYPE:
            where = "Response" if pk.startswith("response") else "Request"
            continue                      # the body section label is rewritten
        if not ptype:
            continue
        if pk in mapfmt.TYPE_BODY and field_path:
            continue                      # a body row, composed rather than carried
        cells = [_cell(row, layout.cols.get(role)) for role in
                 ("ptype", "field", "usage", "dtype")]
        cells.append(_cell(row, layout.sor_col))
        cells += [_cell(row, layout.cols.get(role)) for role in
                  ("example", "desc", "remarks", "inswagger")]
        out.setdefault(where, []).append((row, cells))
    return out


# --------------------------------------------------------------------------- #
# Sheet names
# --------------------------------------------------------------------------- #

_ILLEGAL = re.compile(r"[\[\]:*?/\\]")


def sheet_name(base, suffix, taken):
    """An Excel-legal, unique sheet name of at most 31 characters."""
    base = _ILLEGAL.sub(" ", text(base)).strip() or "Operation"
    suffix = _ILLEGAL.sub(" ", text(suffix)).strip()
    name = ("%s (%s)" % (base, suffix)) if suffix else base
    if len(name) > 31:
        if suffix:
            room = 31 - len(suffix) - 3
            name = "%s (%s)" % (base[:max(room, 1)].strip(), suffix)
        name = name[:31]
    stem, n = name, 2
    while name.lower() in taken:
        tail = " %d" % n
        name = stem[:31 - len(tail)] + tail
        n += 1
    taken.add(name.lower())
    return name


def _endpoint_suffix(endpoint, index, total):
    """A short, recognisable label for one SOR endpoint of an operation."""
    if total <= 1:
        return ""
    stub = text(endpoint).rstrip("/").rsplit("/", 1)[-1]
    stub = re.sub(r"[{}]", "", stub)
    return stub or ("SOR %d" % (index + 1))


# --------------------------------------------------------------------------- #
# Writing
# --------------------------------------------------------------------------- #


def _style(cell, *, bold=False, colour=None, fill=None, wrap=True, size=10):
    cell.font = Font(name=FONT, size=size, bold=bold, color=colour)
    cell.alignment = Alignment(vertical="top", wrap_text=wrap)
    cell.border = BORDER
    if fill is not None:
        cell.fill = fill


def _write_sheet(wb, plan):
    ws = wb.create_sheet(title=plan.title)
    headings = (list(FORMAT_COLUMNS) + [h for _i, h in plan.extra_headings]
                + list(DIFF_COLUMNS))
    headings[SOR_COLUMN] = plan.sor_header

    for r, (label, key_) in enumerate(BANNER_ROWS, start=1):
        c = ws.cell(row=r, column=1, value=label)
        _style(c, bold=True, fill=FILL_BANNER, wrap=False)
        v = ws.cell(row=r, column=2, value=text(plan.banner.get(key_, "")))
        _style(v, fill=FILL_BANNER)

    for i, heading in enumerate(headings, start=1):
        c = ws.cell(row=HEADER_ROW, column=i, value=heading)
        _style(c, bold=True, colour="FFFFFFFF", fill=FILL_HEADER)
    ws.row_dimensions[HEADER_ROW].height = 34

    n_format = len(FORMAT_COLUMNS)
    n_extra = len(plan.extra_headings)
    r = HEADER_ROW + 1
    for entry in plan.rows:
        cells = entry["cells"] + entry["extra"]
        fill = FILL_SECTION if entry["kind"] == "section" else None
        if entry["status"] in DROPPED:
            fill = FILL_DROPPED
        for i, value in enumerate(cells, start=1):
            c = ws.cell(row=r, column=i, value=value if value != "" else None)
            _style(c, bold=(entry["kind"] == "section"), fill=fill)
        for j, value in enumerate((entry["status"], entry["why"],
                                   entry["action"])):
            col = n_format + n_extra + j + 1
            c = ws.cell(row=r, column=col, value=value or None)
            _style(c, fill=(FILL_ACTION if j == 2 and value else FILL_DIFF))
        r += 1

    for i, heading in enumerate(headings, start=1):
        base = (FORMAT_COLUMNS[i - 1] if i <= n_format
                else (DIFF_COLUMNS[i - n_format - n_extra - 1]
                      if i > n_format + n_extra else heading))
        ws.column_dimensions[get_column_letter(i)].width = \
            COLUMN_WIDTH.get(base, 26)
    ws.freeze_panes = ws.cell(row=HEADER_ROW + 1, column=3)
    return ws


PROVENANCE_SHEET = "_Round Trip"


def _fingerprint(path):
    try:
        with open(path, "rb") as fh:
            return hashlib.sha256(fh.read()).hexdigest()[:16]
    except OSError:
        return ""


def previous_round(path):
    """The round number recorded in a workbook this tool wrote before, or 0."""
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception:                                          # noqa: BLE001
        return 0
    try:
        if PROVENANCE_SHEET not in wb.sheetnames:
            return 0
        for row in wb[PROVENANCE_SHEET].iter_rows(max_row=30, max_col=2,
                                                  values_only=True):
            if text(row[0]).lower().startswith("round"):
                digits = re.findall(r"\d+", text(row[1]))
                return int(digits[0]) if digits else 0
    except Exception:                                          # noqa: BLE001
        return 0
    finally:
        wb.close()
    return 0


def _write_provenance(wb, out, plans, skipped, spec_path, round_no):
    ws = wb.create_sheet(title=PROVENANCE_SHEET, index=0)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 74
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 16

    r = 1

    def line(label, value, *, bold=False):
        nonlocal r
        a = ws.cell(row=r, column=1, value=label)
        _style(a, bold=True, fill=FILL_BANNER, wrap=False)
        b = ws.cell(row=r, column=2, value=value)
        _style(b, bold=bold)
        r += 1

    def blank():
        nonlocal r
        r += 1

    def heading(txt):
        nonlocal r
        c = ws.cell(row=r, column=1, value=txt)
        _style(c, bold=True, colour="FFFFFFFF", fill=FILL_HEADER, wrap=False)
        for i in range(2, 7):
            _style(ws.cell(row=r, column=i), fill=FILL_HEADER)
        r += 1

    heading("Where this workbook came from")
    line("Round", str(round_no), bold=True)
    line("Source workbook", os.path.basename(out.workbook))
    line("Input format", {"level": "Level format",
                          "mapping": "field mapping document"}
         .get(out.input_format, out.input_format))
    line("Specification", os.path.basename(spec_path) if spec_path else
         "not written")
    line("Specification fingerprint", _fingerprint(spec_path) if spec_path
         else "")
    line("Written by", "SOR Polymorphizer %s" % __version__)
    line("Written at", _dt.datetime.now(_dt.timezone.utc)
         .replace(microsecond=0).isoformat())
    line("API lifecycle status", "Design")
    line("Verified against the SOR",
         "yes" if out.sor_verified else
         "no, so every SOR field name was taken on trust")
    blank()

    heading("How to use it")
    for txt in (
        "This workbook is the field mapping document for the specification "
        "named above. It was produced by the same run, so the two agree.",
        "The three columns to the right of the format's own nine say what "
        "happened to each element. A row shaded pink did not reach the "
        "interface, and the Analyst Action column says what to change.",
        "Correct those rows here, then feed this workbook back in. It is a "
        "valid input, so the cycle repeats until nothing is flagged.",
        "The specification fingerprint identifies the exact file this "
        "workbook describes. If the specification has been regenerated since, "
        "the fingerprints will differ and this workbook is out of date.",
    ):
        c = ws.cell(row=r, column=1, value=txt)
        _style(c)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 30
        r += 1
    blank()

    heading("What was published, sheet by sheet")
    labels = [show.DISPOSITION_LABEL[k] for k in
              (show.PUBLISHED, show.VARIANT_ONLY, show.NOT_IN_SOR,
               show.NO_FIELD, show.LOST)]
    for i, title in enumerate(["Sheet"] + labels, start=1):
        c = ws.cell(row=r, column=i, value=title)
        _style(c, bold=True, fill=FILL_SECTION)
    r += 1
    for plan in plans:
        _style(ws.cell(row=r, column=1, value=plan.title))
        for i, label in enumerate(labels, start=2):
            _style(ws.cell(row=r, column=i,
                           value=plan.counts.get(label, 0) or None))
        r += 1
    blank()

    if skipped:
        heading("Not exported")
        for sheet, why in skipped:
            _style(ws.cell(row=r, column=1, value=sheet))
            c = ws.cell(row=r, column=2, value=why)
            _style(c)
            ws.merge_cells(start_row=r, start_column=2, end_row=r,
                           end_column=6)
            r += 1
        blank()

    if out.input_format != "mapping":
        heading("Known losses in this conversion")
        for txt in (
            "The source was a Level format workbook, which has no Remarks "
            "column and no Required in Swagger column, so both are empty "
            "here.",
            "The Level format carries no header rows, so the standard header "
            "block above each Request Body was supplied by the tool rather "
            "than read from the workbook.",
            "The Schema and Usage wording is the tool's spelling, not the "
            "analyst's, because there was no original cell to copy.",
        ):
            c = ws.cell(row=r, column=1, value=txt)
            _style(c)
            ws.merge_cells(start_row=r, start_column=1, end_row=r,
                           end_column=6)
            ws.row_dimensions[r].height = 30
            r += 1
    return ws


def default_filename(out, merged_name="openapi"):
    return "%s_field_mapping.xlsx" % (merged_name or "openapi")


def build(out, *, spec_path=""):
    """The workbook, in memory. ``out`` is a :class:`RunResult`."""
    wb = Workbook()
    wb.remove(wb.active)
    views = show.build_views(out)
    verified = bool(out.sor_verified)
    taken, plans = set(), []

    for view in views:
        endpoints = view.endpoints or [""]
        for i, endpoint in enumerate(endpoints):
            suffix = _endpoint_suffix(endpoint, i, len(endpoints))
            title = sheet_name(view.sheet, suffix, taken)
            plan = plan_sheet(view, endpoint, title=title, verified=verified,
                              ordinal=i, total=len(endpoints))
            plans.append(plan)
            _write_sheet(wb, plan)

    skipped = []
    for spec in out.specs:
        if spec.status != "failed":
            continue
        result = next((r for r in show._results_of(out)                # noqa: SLF001
                       if r.sheet == spec.sheet), None)
        carried = _failed_sheet(wb, result, spec, taken)
        if carried is None:
            skipped.append((spec.sheet,
                            "this sheet failed to read and the Level format "
                            "leaves nothing to carry, so it could not be "
                            "written. Correct it in the source workbook."))
        else:
            plans.append(carried)

    _write_provenance(wb, out, plans, skipped, spec_path,
                      previous_round(out.workbook) + 1)
    return wb


def _failed_sheet(wb, result, spec, taken):
    """A sheet that produced no specification, carried so nothing is lost.

    A field mapping sheet is reproduced row for row, because losing an
    operation on the round trip would be worse than any finding it carries. A
    Level sheet has nothing to reproduce and returns ``None``.
    """
    if result is None or not getattr(result, "source_rows", None):
        return None
    layout = getattr(result, "mapping_layout", None)
    if layout is None:
        return None
    banner = dict(result.banner or {})
    title = sheet_name(spec.sheet, "", taken)
    plan = SheetPlan(title, banner, "",
                     "SOR API Field Name\n%s"
                     % text(banner.get("sor_endpoint", "")))
    plan.extra_headings = _extra_columns(result)
    extra_idx = [i for i, _l in plan.extra_headings]
    finding = next((f for f in spec.findings if f.severity == "error"), None)
    rows = result.source_rows
    for r in range(layout.header_row, len(rows)):
        row = rows[r]
        if not any(text(c) for c in row):
            continue
        cells = [_cell(row, layout.cols.get(role)) for role in
                 ("ptype", "field", "usage", "dtype")]
        cells.append(_cell(row, layout.sor_col))
        cells += [_cell(row, layout.cols.get(role)) for role in
                  ("example", "desc", "remarks", "inswagger")]
        plan.add(cells, kind="carried",
                 extra=[_cell(row, i) for i in extra_idx])
    if plan.rows:
        first = plan.rows[0]
        first["status"] = NOT_GENERATED
        first["why"] = (finding.what if finding else
                        "this sheet produced no specification")
        first["action"] = (finding.fix if finding else
                           "see the generation report")
    plan.counts = {NOT_GENERATED: 1}
    plan.note = "not generated"
    _write_sheet(wb, plan)
    return plan


def write(out, path, *, spec_path=""):
    """Write the field mapping document and return its path."""
    wb = build(out, spec_path=spec_path)
    wb.save(path)
    return path


__all__ = ["BANNER_ROWS", "FORMAT_COLUMNS", "DIFF_COLUMNS", "HEADER_ROW",
           "REQUEST_HEADERS", "RESPONSE_HEADERS", "PROVENANCE_SHEET",
           "SheetPlan", "plan_sheet", "sheet_name", "previous_round",
           "default_filename", "build", "write"]
