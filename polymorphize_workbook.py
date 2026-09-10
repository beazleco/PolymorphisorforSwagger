"""
polymorphize_workbook — reader for the Level-indented SOR mapping contract.

Version 6.6.

This module replaces the dotted-path reader used up to v5.3. The authoritative
workbook format expresses schema nesting through a run of ``Level 1`` ..
``Level N`` columns rather than through a dotted path in a single column, and
one worksheet describes one endpoint rather than one workbook describing one
specification.

Contract, as read from the reference credit card workbook
=========================================================

Sheet classification
    A worksheet is an *operation sheet* when its header row carries at least
    two columns matching ``Level <n>``. Everything else is a support sheet and
    is skipped without comment (the reference workbook carries a ``SOR Payload``
    sheet of raw sample payloads).

Header row
    The row, within the first few, holding the most ``Level <n>`` headers.
    Observed at row 1 in every operation sheet.

Banner
    Label/value pairs in the columns to the left of ``Level 1``, sitting *below*
    the header row. Recognised labels are Use Case, Service Domain, Equivalent
    BIAN API Endpoint, Proposed Business API Endpoint and SOR API Endpoint.

Sections
    Also in the label column. Only **Request Body** and **Response Body** are
    consumed, with bare ``Request`` and ``Response`` accepted as aliases for
    the older sheets that use them. Header and parameter sections are ignored
    by design: they carry no schema the polymorphizer is responsible for.

    A sheet that offers no Request Body at all *fails*. It does not take the
    rest of the workbook down with it.

Structure break
    Analysts paste raw SOR payload samples below the mapping grid. Those rows
    land in Level columns because JSON braces occupy the same cells. Parsing
    therefore stops at the first run of ``STRUCTURE_BREAK`` rows that carry
    neither a Level cell nor a label, and anything past that point is ignored.
    The reference workbook separates its pasted payloads with four blank rows;
    the largest blank run *inside* a genuine section is two.

Rows
    Depth is the position of the first populated Level column. A row attaches
    to the nearest preceding row one level shallower. The shallowest row in a
    section names the schema.

SOR columns
    Any header beginning ``SOR``. The remainder of the header names the SOR
    endpoint, so ``SOR /v1/account/contact`` is the endpoint
    ``/v1/account/contact`` and a bare ``SOR Field`` is the single endpoint
    declared in the banner. A non-empty cell means that endpoint supplies the
    attribute; an empty cell means it does not.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import polymorphize_core as core

__version__ = "6.6"

# --------------------------------------------------------------------------- #
# Contract constants
# --------------------------------------------------------------------------- #

#: Consecutive rows with neither a Level cell nor a label that end the grid.
STRUCTURE_BREAK = 3

#: How far down to look for the header row.
HEADER_SEARCH_ROWS = 6

LEVEL_RE = re.compile(r"^level\s*(\d+)$")

#: Section labels the reader consumes, normalised.
SECTION_REQUEST = "request"
SECTION_RESPONSE = "response"

SECTION_ALIASES = {
    "requestbody": SECTION_REQUEST,
    "request": SECTION_REQUEST,
    "responsebody": SECTION_RESPONSE,
    "response": SECTION_RESPONSE,
}

#: Section labels deliberately skipped. Recognised so that they terminate the
#: preceding section rather than being mistaken for stray content.
SECTION_IGNORED = {
    "requestheader",
    "requestheaders",
    "responseheader",
    "responseheaders",
    "requestparameter",
    "requestparameters",
    "queryparameter",
    "queryparameters",
    "pathparameter",
    "pathparameters",
    "header",
    "headers",
    "errorresponse",
    "errors",
}

#: Column role synonyms, normalised.
ROLE_SYNONYMS = {
    "dummyvalue": "example",
    "example": "example",
    "examples": "example",
    "samplevalue": "example",
    "datatype": "type",
    "type": "type",
    "required": "usage",
    "usage": "usage",
    "mandatory": "usage",
    "description": "description",
    "descriptions": "description",
    "remarks": "remarks",
    "notes": "remarks",
    "comment": "remarks",
    "comments": "remarks",
    "valuetype": "valuetype",
    "format": "format",
    "ref": "ref",
}

BANNER_KEYS = {
    "usecase": "use_case",
    "servicedomain": "service_domain",
    "equivalentbianapiendpoint": "bian_endpoint",
    "bianapiendpoint": "bian_endpoint",
    "proposedbusinessapiendpoint": "business_endpoint",
    "proposedreusableapiendpoint": "business_endpoint",
    "proposedreuseableapiendpoint": "business_endpoint",
    "sorapiendpoint": "sor_endpoint",
    "method": "method",
}

#: Usage vocabulary. ``True`` means the attribute lands in ``required``.
USAGE_REQUIRED = "mandatory"
USAGE_CONDITIONAL = "conditionalmandatory"
USAGE_OPTIONAL = "optional"

USAGE_VOCAB = {
    "mandatory": USAGE_REQUIRED,
    "m": USAGE_REQUIRED,
    "required": USAGE_REQUIRED,
    "conditionalmandatory": USAGE_CONDITIONAL,
    "conditionallymandatory": USAGE_CONDITIONAL,
    "conditional": USAGE_CONDITIONAL,
    "cm": USAGE_CONDITIONAL,
    "optional": USAGE_OPTIONAL,
    "o": USAGE_OPTIONAL,
    "": USAGE_OPTIONAL,
}

#: Cell values in an SOR column that mean "this endpoint does not supply it".
NOT_USED_VALUES = {
    "",
    "-",
    "--",
    "n/a",
    "na",
    "none",
    "not used",
    "notused",
    "not applicable",
    "no",
    "x",
    "tbc",
    "tbd",
}

#: Names that mark a container rather than an attribute.
CONTAINER_TYPES = {"object", "array"}

#: Data Type parsing. Base name to (JSON type, format).
TYPE_MAP = {
    "string": ("string", None),
    "str": ("string", None),
    "text": ("string", None),
    "char": ("string", None),
    "alphanumeric": ("string", None),
    "number": ("number", None),
    "numeric": ("number", None),
    "decimal": ("number", None),
    "amount": ("number", None),
    "integer": ("integer", None),
    "int": ("integer", None),
    "long": ("integer", None),
    "boolean": ("boolean", None),
    "bool": ("boolean", None),
    "date": ("string", "date"),
    "datetime": ("string", "date-time"),
    "date-time": ("string", "date-time"),
    "timestamp": ("string", "date-time"),
    "time": ("string", "time"),
    "object": ("object", None),
    "array": ("array", None),
    "list": ("array", None),
    "binary": ("string", "byte"),
    "base64": ("string", "byte"),
}

#: A plausible attribute or schema name. Guards against pasted JSON.
NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_\-]*$")

#: The request property whose value selects the SOR endpoint.
VARIANT_PROPERTY = "requestVariant"

# --------------------------------------------------------------------------- #
# Findings
# --------------------------------------------------------------------------- #


@dataclass
class Finding:
    """One problem, located precisely enough for an analyst to act on it."""

    code: str
    severity: str  # "error" | "warning"
    sheet: str
    cell: str
    what: str
    expected: str
    fix: str

    @property
    def where(self):
        return "%s!%s" % (self.sheet, self.cell) if self.cell else self.sheet

    def __str__(self):
        return "[%s] %s %s: %s" % (self.code, self.severity.upper(), self.where, self.what)


def col_letter(index_zero_based):
    """0 -> A, 25 -> Z, 26 -> AA."""
    n = index_zero_based + 1
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def cell_ref(row_one_based, col_zero_based):
    return "%s%d" % (col_letter(col_zero_based), row_one_based)


# --------------------------------------------------------------------------- #
# Normalisation
# --------------------------------------------------------------------------- #


def text(value):
    """A cell as collapsed single-line text."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return " ".join(str(value).split())


def text_lines(value):
    """A cell with its line breaks kept, each line collapsed.

    Description cells carry the ``requestVariant`` enumeration one value per
    line, so the line structure is meaning, not formatting.
    """
    if value is None:
        return ""
    lines = [" ".join(ln.split()) for ln in str(value).splitlines()]
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines)


def key(value):
    """Lower case, alphanumerics only. Used for every vocabulary lookup."""
    return re.sub(r"[^a-z0-9]+", "", text(value).lower())


def clean_example(value):
    """Analyst example cells carry stray quotes: ``'PHP''`` and ``\"0014''``."""
    s = text(value)
    s = s.strip().strip("\"'").rstrip("'\"").strip()
    return s


# --------------------------------------------------------------------------- #
# Layout
# --------------------------------------------------------------------------- #


@dataclass
class SorColumn:
    index: int
    endpoint: str
    header: str


@dataclass
class Layout:
    header_row: int
    level_cols: list          # column indices, shallow to deep
    label_cols: list          # indices left of Level 1
    roles: dict               # role name -> column index
    sor_cols: list            # list[SorColumn]

    @property
    def first_level(self):
        return self.level_cols[0]

    def depth_of(self, row):
        """Index into ``level_cols`` of the first populated Level cell, or -1."""
        for d, c in enumerate(self.level_cols):
            if c < len(row) and text(row[c]):
                return d
        return -1

    def name_at(self, row, depth):
        c = self.level_cols[depth]
        return text(row[c]) if c < len(row) else ""

    def label_of(self, row):
        for c in self.label_cols:
            if c < len(row) and text(row[c]):
                return text(row[c]), c
        return "", -1

    def cell(self, row, role, keep_lines=False):
        c = self.roles.get(role)
        if c is None or c >= len(row):
            return ""
        return text_lines(row[c]) if keep_lines else text(row[c])


def detect_layout(rows, sheet, findings):
    """Locate the header row and classify its columns, or return ``None``."""
    best, best_count = None, 0
    for i, row in enumerate(rows[:HEADER_SEARCH_ROWS]):
        count = sum(1 for c in row if LEVEL_RE.match(key(c).replace("level", "level ")) or
                    LEVEL_RE.match(text(c).strip().lower()))
        if count > best_count:
            best, best_count = i, count
    if best is None or best_count < 2:
        return None

    header = rows[best]
    level_pairs = []
    roles, sor_cols = {}, []
    for idx, raw in enumerate(header):
        label = text(raw)
        low = label.strip().lower()
        m = LEVEL_RE.match(low)
        if m:
            level_pairs.append((int(m.group(1)), idx))
            continue
        k = key(label)
        if not k:
            continue
        if k.startswith("sor"):
            endpoint = label.strip()[3:].strip().lstrip(":").strip()
            for noise in ("field", "fields", "column"):
                if endpoint.lower() == noise:
                    endpoint = ""
            sor_cols.append(SorColumn(idx, endpoint, label))
            continue
        role = ROLE_SYNONYMS.get(k)
        if role and role not in roles:
            roles[role] = idx

    level_pairs.sort()
    level_cols = [idx for _n, idx in level_pairs]
    expected = list(range(1, len(level_pairs) + 1))
    if [n for n, _ in level_pairs] != expected:
        findings.append(Finding(
            "L001", "error", sheet, cell_ref(best + 1, level_cols[0]),
            "the Level columns are numbered %s" % ", ".join(str(n) for n, _ in level_pairs),
            "Level 1 to Level %d with no number missing or repeated" % len(level_pairs),
            "renumber the Level headers so they run 1, 2, 3 ... with no gaps",
        ))
    label_cols = [i for i in range(level_cols[0])]
    if not label_cols:
        findings.append(Finding(
            "L002", "error", sheet, cell_ref(best + 1, 0),
            "Level 1 is the first column, so there is nowhere for the section labels",
            "at least one column to the left of Level 1",
            "insert a column before Level 1 to hold Use Case, Service Domain and the "
            "Request Body and Response Body labels",
        ))
    return Layout(best, level_cols, label_cols, roles, sor_cols)


# --------------------------------------------------------------------------- #
# Attribute rows
# --------------------------------------------------------------------------- #


@dataclass
class Node:
    name: str
    depth: int
    row: int
    json_type: str = None
    fmt: str = None
    max_length: int = None
    usage: str = USAGE_OPTIONAL
    description: str = ""
    example: str = ""
    remarks: str = ""
    sor: dict = field(default_factory=dict)   # endpoint -> SOR field name
    #: The workbook's word, kept unchanged. ``sor`` is the effective mapping
    #: and verification against the SOR specification may clear entries in it;
    #: this is what the analyst actually wrote, which the reports need in order
    #: to explain why an element was removed.
    sor_declared: dict = field(default_factory=dict)
    children: list = field(default_factory=list)
    parent: object = None

    @property
    def is_container(self):
        return self.json_type in CONTAINER_TYPES

    def supported_by(self, endpoint):
        return bool(self.sor.get(endpoint))

    @property
    def endpoints(self):
        return {e for e, v in self.sor.items() if v}


def parse_type(raw, sheet, row, col, findings):
    """``String (19)`` -> ("string", None, 19).

    A blank cell returns ``None`` for the type. The caller resolves it once the
    tree is built, because a blank type on a row that has children means
    "object" and a blank type on a leaf means an unstated scalar.
    """
    s = text(raw)
    if not s:
        return None, None, None
    m = re.match(r"^\s*([A-Za-z][A-Za-z0-9_\- ]*?)\s*(?:\(\s*(\d+)\s*(?:[,.]\s*\d+\s*)?\)|\s(\d+))?\s*$", s)
    base = m.group(1) if m else s
    length = None
    if m:
        length = m.group(2) or m.group(3)
    k = key(base)
    if k not in TYPE_MAP:
        findings.append(Finding(
            "T001", "error", sheet, cell_ref(row, col),
            "Data Type reads %r, which is not a recognised type" % s,
            "one of " + ", ".join(sorted({t for t in TYPE_MAP})),
            "replace it with a recognised type, optionally with a length in "
            "brackets, for example String (19)",
        ))
        return None, None, None
    json_type, fmt = TYPE_MAP[k]
    try:
        length = int(length) if length else None
    except (TypeError, ValueError):
        length = None
    return json_type, fmt, length


def parse_usage(raw, sheet, row, col, findings):
    k = key(raw)
    if k in USAGE_VOCAB:
        return USAGE_VOCAB[k]
    findings.append(Finding(
        "U001", "error", sheet, cell_ref(row, col),
        "Usage reads %r" % text(raw),
        "Mandatory, ConditionalMandatory or blank for optional",
        "replace the value with one of the three permitted words, or clear the "
        "cell if the attribute is optional",
    ))
    return USAGE_OPTIONAL


# --------------------------------------------------------------------------- #
# Sections
# --------------------------------------------------------------------------- #


@dataclass
class Section:
    kind: str          # SECTION_REQUEST | SECTION_RESPONSE
    label: str         # as written
    row: int           # 1-based row of the label
    root: object = None
    nodes: list = field(default_factory=list)


@dataclass
class SheetResult:
    sheet: str
    status: str = "ok"            # "ok" | "failed" | "skipped"
    layout: object = None
    banner: dict = field(default_factory=dict)
    request: object = None
    response: object = None
    endpoints: list = field(default_factory=list)   # SOR endpoints, in column order
    variants: list = field(default_factory=list)    # requestVariant enum labels
    findings: list = field(default_factory=list)
    ignored_sections: list = field(default_factory=list)
    ignored_rows: int = 0
    #: The sheet exactly as it was read, kept only by the field mapping reader
    #: and only so that the exporter can write it back without losing the rows
    #: the tool does not interpret. See :mod:`polymorphize_export`. It is the
    #: same list the reader already built, so it costs no extra memory.
    source_rows: list = field(default_factory=list)
    #: The field mapping layout, when the sheet came from that format.
    mapping_layout: object = None

    @property
    def errors(self):
        return [f for f in self.findings if f.severity == "error"]

    @property
    def warnings(self):
        return [f for f in self.findings if f.severity == "warning"]

    def fail(self, finding):
        self.findings.append(finding)
        self.status = "failed"


def _read_banner(rows, layout, result):
    """Label/value pairs between the header row and the first section."""
    banner = {}
    for i in range(layout.header_row + 1, len(rows)):
        row = rows[i]
        label, col = layout.label_of(row)
        if not label:
            continue
        k = key(label)
        if k in SECTION_ALIASES or k in SECTION_IGNORED:
            break
        mapped = BANNER_KEYS.get(k)
        values = [text(row[c]) for c in range(col + 1, min(len(row), layout.first_level + 6))
                  if c < len(row) and text(row[c])]
        if mapped:
            banner[mapped] = values[0] if values else ""
            if mapped == "service_domain":
                for v in values[1:]:
                    m = re.match(r"^\s*BQ\s*:?\s*(.+)$", v, re.I)
                    if m:
                        banner["behaviour_qualifier"] = m.group(1).strip()
                        break
        elif values:
            banner.setdefault("_extra", []).append((label, values[0]))
    result.banner = banner
    return banner


def _sections(rows, layout, result):
    """Walk the label column, honouring the structure break."""
    found, blanks = [], 0
    current = None
    for i in range(layout.header_row + 1, len(rows)):
        row = rows[i]
        label, _col = layout.label_of(row)
        depth = layout.depth_of(row)

        if not label and depth < 0:
            blanks += 1
            if blanks >= STRUCTURE_BREAK:
                result.ignored_rows = max(0, len(rows) - i)
                break
            continue
        blanks = 0

        if label:
            k = key(label)
            if k in SECTION_ALIASES:
                current = Section(SECTION_ALIASES[k], label, i + 1)
                found.append(current)
                if depth >= 0:
                    current.nodes.append((i + 1, row, depth))
                continue
            if k in SECTION_IGNORED:
                result.ignored_sections.append((label, i + 1))
                current = None
                continue
            if k in BANNER_KEYS:
                continue
            if depth < 0:
                # Not a section, not a banner key, no structure. Stray content.
                continue

        if current is not None and depth >= 0:
            current.nodes.append((i + 1, row, depth))
    return found


def _build_tree(section, layout, result):
    """Attach each row to the nearest preceding row one level shallower."""
    sheet = result.sheet
    stack = []          # list of Node, index == depth
    root = None
    endpoints = [c.endpoint for c in layout.sor_cols]

    for row_no, row, depth in section.nodes:
        name = layout.name_at(row, depth)
        if not NAME_RE.match(name):
            result.findings.append(Finding(
                "A001", "warning", sheet, cell_ref(row_no, layout.level_cols[depth]),
                "%r is not a usable attribute name, so the row was ignored" % name[:60],
                "a name starting with a letter or underscore, then letters, digits, "
                "hyphens or underscores",
                "correct the name, or move the content out of the Level columns if it "
                "is a pasted sample rather than an attribute",
            ))
            continue

        if root is None:
            if depth != 0:
                result.findings.append(Finding(
                    "A002", "warning", sheet, cell_ref(row_no, layout.level_cols[depth]),
                    "the first row of %s sits at Level %d" % (section.label, depth + 1),
                    "the schema name at Level 1",
                    "move the schema name into the Level 1 column",
                ))
        elif depth > len(stack):
            result.findings.append(Finding(
                "A003", "error", sheet, cell_ref(row_no, layout.level_cols[depth]),
                "%r is at Level %d but the row above it is at Level %d, so %d "
                "Level column%s skipped"
                % (name, depth + 1, len(stack), depth - len(stack),
                   " is" if depth - len(stack) == 1 else "s are"),
                "each row is at most one Level deeper than the row above it",
                "move %r left to Level %d, or insert its missing parent object at "
                "Level %d" % (name, len(stack) + 1, len(stack)),
            ))
            depth = len(stack)

        t_col = layout.roles.get("type", layout.level_cols[-1])
        u_col = layout.roles.get("usage", layout.level_cols[-1])
        json_type, fmt, length = parse_type(
            layout.cell(row, "type"), sheet, row_no, t_col, result.findings)
        usage = parse_usage(layout.cell(row, "usage"), sheet, row_no, u_col, result.findings) \
            if layout.cell(row, "usage") else USAGE_OPTIONAL

        node = Node(
            name=name, depth=depth, row=row_no,
            json_type=json_type, fmt=fmt, max_length=length, usage=usage,
            description=layout.cell(row, "description", keep_lines=True),
            example=clean_example(layout.cell(row, "example")),
            remarks=text(layout.cell(row, "remarks")),
        )
        for col in layout.sor_cols:
            raw = text(row[col.index]) if col.index < len(row) else ""
            value = "" if key(raw) in NOT_USED_VALUES else raw
            node.sor[col.endpoint] = value
            node.sor_declared[col.endpoint] = value

        if root is None:
            root = node
            stack = [node]
            continue

        if depth == 0:
            result.findings.append(Finding(
                "A004", "error", sheet, cell_ref(row_no, layout.level_cols[0]),
                "%r is a second row at Level 1 in %s" % (name, section.label),
                "exactly one Level 1 row per section, naming the schema",
                "move %r to Level 2 if it belongs to %r, or split it into its own "
                "section" % (name, root.name),
            ))
            depth = 1
            node.depth = 1

        parent = stack[depth - 1] if depth - 1 < len(stack) else stack[-1]
        node.parent = parent
        parent.children.append(node)
        stack = stack[:depth] + [node]

    if root is None:
        return None

    # Resolve blank Data Type cells now that the shape of the tree is known.
    def resolve(n):
        if n.json_type is None:
            if n.children:
                n.json_type = "object"
            else:
                n.json_type = "string"
                result.findings.append(Finding(
                    "T002", "warning", sheet,
                    cell_ref(n.row, layout.roles.get("type", layout.level_cols[-1])),
                    "%r has no Data Type and no nested rows, so it was treated as a "
                    "string" % n.name,
                    "a Data Type for every attribute row",
                    "set the Data Type of %r, for example String (35) or Number(12)"
                    % n.name,
                ))
        for c in n.children:
            resolve(c)
    resolve(root)

    # Containers declared as scalars, and scalars declared as containers.
    def walk(n):
        if n.children and not n.is_container:
            result.findings.append(Finding(
                "A005", "warning", sheet, cell_ref(n.row, layout.roles.get("type", 0)),
                "%r has %d nested row%s but its Data Type is %r"
                % (n.name, len(n.children), "" if len(n.children) == 1 else "s", n.json_type),
                "object for a nested group, or Array for a repeating group",
                "set the Data Type of %r to object or Array" % n.name,
            ))
            n.json_type = "object"
        if n.is_container and not n.children:
            result.findings.append(Finding(
                "A006", "warning", sheet, cell_ref(n.row, layout.roles.get("type", 0)),
                "%r is declared %s but has no nested rows beneath it, so "
                "there is no shape to publish and the element was removed "
                "from the interface" % (n.name, n.json_type),
                "at least one row one Level deeper, or a scalar Data Type",
                "add the attributes of %r one Level deeper, or change its Data Type "
                "to a scalar" % n.name,
            ))
        for c in n.children:
            walk(c)
    walk(root)

    # Duplicate siblings.
    def dupes(n, path):
        seen = {}
        for c in n.children:
            k = c.name.lower()
            if k in seen:
                result.findings.append(Finding(
                    "A007", "error", sheet, cell_ref(c.row, layout.level_cols[c.depth]),
                    "%r appears twice under %s" % (c.name, path or root.name),
                    "each name once within its parent",
                    "rename one of the two rows, or delete the duplicate at row %d"
                    % seen[k],
                ))
            else:
                seen[k] = c.row
            dupes(c, "%s.%s" % (path, c.name) if path else c.name)
    dupes(root, "")

    section.root = root
    result.endpoints = endpoints
    return root


def _variant_labels(request_root):
    """Enumeration labels declared on the requestVariant description."""
    if request_root is None:
        return []
    target = None

    def find(n):
        nonlocal target
        if n.name.lower() == VARIANT_PROPERTY.lower():
            target = n
        for c in n.children:
            find(c)
    find(request_root)
    if target is None:
        return []
    out = []
    for line in (target.description or "").splitlines():
        m = re.match(r"^\s*(\d+)\s*[.)-]\s*(.+?)\s*$", line)
        if m:
            out.append((m.group(1), m.group(2)))
    if not out:
        # Single-line form: "01. A 02. B 03. C"
        parts = re.findall(r"(\d{1,2})\s*[.)-]\s*([^0-9]+?)(?=\s+\d{1,2}\s*[.)-]|$)",
                           target.description or "")
        out = [(a, b.strip()) for a, b in parts]
    return out


def read_sheet(grid, *, strict=False):
    """One worksheet to a :class:`SheetResult`. Never raises on content."""
    rows = [list(r) for r in grid._rows]          # noqa: SLF001 - same package
    result = SheetResult(sheet=grid.title)

    probe = []
    layout = detect_layout(rows, grid.title, probe)
    if layout is None:
        result.status = "skipped"
        return result
    result.findings.extend(probe)
    result.layout = layout

    _read_banner(rows, layout, result)
    sections = _sections(rows, layout, result)

    by_kind = {}
    for s in sections:
        if s.kind in by_kind:
            result.findings.append(Finding(
                "S002", "warning", grid.title, cell_ref(s.row, layout.label_cols[0]),
                "a second %s section labelled %r" % (s.kind, s.label),
                "one Request Body and one Response Body per sheet",
                "merge it into the first %s section, or move it to its own sheet" % s.kind,
            ))
            continue
        by_kind[s.kind] = s

    # The failure the analyst has to see: no Request Body at all.
    if SECTION_REQUEST not in by_kind:
        mislabelled = [(lab, r) for lab, r in result.ignored_sections
                       if key(lab).startswith("request")]
        if mislabelled:
            lab, r = mislabelled[0]
            result.fail(Finding(
                "S001", "error", grid.title, cell_ref(r, layout.label_cols[0]),
                "this sheet has no Request Body section. The only request section is "
                "%r at row %d, and the rows beneath it carry the request schema, so "
                "the label is wrong" % (lab, r),
                "a section labelled Request Body holding the request schema",
                "change %r in cell %s to Request Body. If the sheet genuinely "
                "describes request headers, add a separate Request Body section for "
                "the schema." % (lab, cell_ref(r, layout.label_cols[0])),
            ))
        else:
            result.fail(Finding(
                "S001", "error", grid.title, cell_ref(layout.header_row + 2, layout.label_cols[0]),
                "this sheet has no Request Body section",
                "a section labelled Request Body holding the request schema",
                "add a Request Body label in column %s above the request schema rows"
                % col_letter(layout.label_cols[0]),
            ))

    if SECTION_RESPONSE not in by_kind:
        result.fail(Finding(
            "S003", "error", grid.title,
            cell_ref(layout.header_row + 2, layout.label_cols[0]),
            "this sheet has no Response Body section",
            "a section labelled Response Body holding the response schema",
            "add a Response Body label in column %s above the response schema rows"
            % col_letter(layout.label_cols[0]),
        ))

    if not layout.sor_cols:
        result.fail(Finding(
            "M001", "error", grid.title, cell_ref(layout.header_row + 1, layout.level_cols[-1]),
            "no SOR column was found on the header row",
            "at least one header beginning SOR, for example SOR /v1/card/details",
            "add an SOR column for each SOR endpoint this operation calls",
        ))

    for kind, section in by_kind.items():
        root = _build_tree(section, layout, result)
        if root is None:
            # A declared-but-empty section is permitted. The label is the
            # analyst's statement that the message was considered, so an empty
            # one is a decision rather than an omission. A *missing* label is
            # still a failure: see S001 and S003 above.
            result.findings.append(Finding(
                "S004", "warning", grid.title,
                cell_ref(section.row, layout.label_cols[0]),
                "the %s section has no rows beneath it, so the message carries "
                "no payload" % section.label,
                "either a schema at Level 1 with its attributes below, or an "
                "empty section if the message genuinely has no payload",
                "add the schema rows beneath %r if it should carry a payload. "
                "If it should not, leave the section empty and no payload will "
                "be published for it." % section.label,
            ))
        if kind == SECTION_REQUEST:
            result.request = section
        else:
            result.response = section

    result.variants = _variant_labels(result.request.root if result.request else None)

    if len(layout.sor_cols) > 1 and not result.variants:
        result.findings.append(Finding(
            "V001", "warning", grid.title, cell_ref(layout.header_row + 1, layout.sor_cols[0].index),
            "this sheet has %d SOR columns but the request declares no %s values"
            % (len(layout.sor_cols), VARIANT_PROPERTY),
            "a %s attribute whose Description lists the variants, one per line, "
            "as 01. Name" % VARIANT_PROPERTY,
            "add %s to the request and list one numbered variant per SOR column, or "
            "confirm that the variant is chosen by which identifier the caller "
            "supplies" % VARIANT_PROPERTY,
        ))
    elif result.variants and len(result.variants) != len(layout.sor_cols):
        result.findings.append(Finding(
            "V002", "error", grid.title,
            cell_ref(layout.header_row + 1, layout.sor_cols[0].index),
            "%d %s values are declared but there are %d SOR columns"
            % (len(result.variants), VARIANT_PROPERTY, len(layout.sor_cols)),
            "one SOR column per declared variant",
            "add the missing SOR column, or remove the variant that has no SOR "
            "endpoint behind it",
        ))

    if strict:
        for f in result.findings:
            if f.severity == "warning":
                f.severity = "error"
        if result.errors:
            result.status = "failed"
    return result


def read_workbook(path, *, strict=False, max_rows=20000):
    """Every operation sheet of a workbook, each isolated from the others."""
    grids = core.load_grids(path, max_rows=max_rows)
    out = []
    for g in grids:
        try:
            out.append(read_sheet(g, strict=strict))
        except Exception as exc:                       # noqa: BLE001
            r = SheetResult(sheet=g.title)
            r.fail(Finding(
                "X001", "error", g.title, "",
                "the reader stopped on this sheet: %s: %s" % (type(exc).__name__, exc),
                "a sheet the reader can parse end to end",
                "send this sheet to the maintainer. The rest of the workbook was "
                "processed.",
            ))
            out.append(r)
    return out


__all__ = [
    "Finding", "Layout", "Node", "Section", "SheetResult", "SorColumn",
    "cell_ref", "col_letter", "detect_layout", "key", "parse_type",
    "parse_usage", "read_sheet", "read_workbook", "text",
    "SECTION_REQUEST", "SECTION_RESPONSE", "STRUCTURE_BREAK", "VARIANT_PROPERTY",
]
