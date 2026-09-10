"""
polymorphize_mapping — the dotted-path field mapping document.

Version 6.6.

The second supported input format, evidenced by
``apicoeissueddeviceadministrationfieldmappingv1.0.7.xlsx``. It is the document
the API COE already maintains and already feeds to its own swagger tool, which
is why it is worth reading directly rather than asking analysts to re-cut every
data model into the Level format.

The format
==========

One sheet per operation. Banner in rows 1 to 9, header on row 10::

    API Name                        | Card Create API
    Use Case:                       | The API will be used by ...
    Service Domain:                 | Issued Device Administration
    BQ                              | Card Device Assignment
    Method                          | POST
    Equivalent BIAN API Endpoint:   | N/A
    Proposed Reusable API Endpoint: | /issued-device-administration/...
    SOR Name                        | Thales
    SOR API Endpoint:               | POST: /v2/issuers/{issuerId}/cards

    Parameter Type | Reusable API Field Name | Usage | Schema |
    SOR API Field Name | Example | Description | Remarks | Required in Swagger

Hierarchy is a dotted path in the field name column, ``CardCreateRequest``,
``CardCreateRequest.PartyIdentifier``, ``CardCreateRequest.PartyIdentifier
.partyId``, up to eight segments deep. A row whose Parameter Type cell is a
section banner and whose field name is empty opens a section.

This reader produces the same :class:`polymorphize_workbook.SheetResult` and
the same ``Node`` tree as the Level reader, so verification against the System
of Record, the merge, the showcase and every finding downstream work unchanged.

Decisions taken here, each recorded in DOCUMENTATION_NOTES.md
============================================================

**Parameter Type is authoritative for which section a row belongs to**, and the
section banner corroborates it. The row-level type is the more robust of the
two and would have caught the mislabelled section that is the single defect in
the credit card workbook. A disagreement is reported as ``F001`` rather than
resolved silently, which is the same treatment the banner and the SOR column
header already get for the endpoint.

**Only body rows are read.** ``Header``, ``Request Parameter`` and any other
Parameter Type are skipped, which is the dotted-path equivalent of the Level
format's rule that only ``Request Body`` and ``Response Body`` are read.

**``N/A`` and an empty cell both mean unmapped**, because both appear in the
wild, 216 and 289 rows respectively in the reference document. They are not
merged, though: ``N/A`` is a decision recorded and an empty cell is a decision
not yet stated, so an empty cell is reported at strict level as ``F004`` while
``N/A`` passes silently.

**``Required in Swagger`` is advisory and never overrides the mapping.** It is
blank on 607 of 783 rows in the reference document, so it cannot carry the
weight of an include or exclude instruction. Where present it is recorded as
``x-required-in-swagger`` and nothing else.

**One SOR field column, so the variant axis cannot be expressed.** Everything
generated from this format is a single pruned schema per message. Where a sheet
looks as though it wants variants the reader says so, once, as ``F005``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

import polymorphize_core as core
import polymorphize_workbook as wbk
from polymorphize_workbook import (
    Finding, SECTION_REQUEST, SECTION_RESPONSE, USAGE_CONDITIONAL,
    USAGE_OPTIONAL, USAGE_REQUIRED, Layout, Node, Section, SheetResult,
    SorColumn, cell_ref, col_letter, key, text, text_lines,
)

__version__ = "6.6"

#: Rows searched for the header. The reference document puts it on row 10; the
#: allowance covers a sheet that carries a title row above its banner.
HEADER_SEARCH_ROWS = 16

#: Banner labels, keyed on a normalised form, mapped to the same banner keys
#: the Level reader produces so everything downstream is unchanged.
BANNER_KEYS = {
    "apiname": "api_name",
    "usecase": "use_case",
    "servicedomain": "service_domain",
    "bq": "behaviour_qualifier",
    "behaviourqualifier": "behaviour_qualifier",
    "method": "method",
    "equivalentbianapiendpoint": "bian_endpoint",
    "proposedreusableapiendpoint": "business_endpoint",
    "proposedbusinessapiendpoint": "business_endpoint",
    "sorname": "sor_name",
    "sorapiendpoint": "sor_endpoint",
}

#: Column roles, in assignment priority order. The first unassigned column
#: whose normalised header contains one of the hints takes the role.
HEADER_ROLES = [
    ("ptype", ("parametertype", "paramtype")),
    ("sor", ("sorapifieldname", "sorfieldname", "sorfield", "sorattribute",
             "sorapifield", "sormapping", "sorname", "sor")),
    ("field", ("reusableapifieldname", "apifieldname", "fieldname",
               "attributename", "elementname", "attribute", "element",
               "field", "name")),
    ("usage", ("usage", "mandatoryoptional", "cardinality", "required")),
    ("dtype", ("schema", "datatype", "type", "format")),
    ("example", ("examples", "example", "samplevalue", "sample")),
    ("desc", ("description", "definition")),
    ("remarks", ("remarks", "comments", "comment", "notes")),
    ("inswagger", ("requiredinswagger", "inswagger", "includeinswagger")),
]

#: Parameter Type values that open a section rather than describe a row.
SECTION_BY_TYPE = {
    "requestbody": SECTION_REQUEST,
    "request": SECTION_REQUEST,
    "body": SECTION_REQUEST,          # two sheets label the request "Body"
    "responsebody": SECTION_RESPONSE,
    "response": SECTION_RESPONSE,
}

#: Parameter Type values recognised and deliberately skipped.
TYPE_IGNORED = {
    "header", "requestheader", "responseheader", "requestparameter",
    "responseparameter", "queryparameter", "pathparameter", "query", "path",
    "responsecode", "requestqueryparameter",
}

#: Parameter Type values that mark a row as payload.
TYPE_BODY = {"body", "requestbody", "responsebody"}

#: SOR cell contents that mean the System of Record does not supply the
#: element. ``N/A`` is a decision recorded; an empty cell is silence.
NOT_USED = {"", "-", "--", "n/a", "na", "none", "null", "nil", "tbc", "tbd",
            "tba", "notused", "notapplicable", "nomapping"}

#: Schema column vocabulary. Mirrors the Level reader's table and adds the
#: spellings this format uses.
TYPE_MAP = {
    "string": ("string", None), "str": ("string", None),
    "text": ("string", None), "char": ("string", None),
    "number": ("number", None), "decimal": ("number", None),
    "amount": ("number", None), "float": ("number", None),
    "double": ("number", None),
    "integer": ("integer", None), "int": ("integer", None),
    "long": ("integer", None),
    "boolean": ("boolean", None), "bool": ("boolean", None),
    "date": ("string", "date"),
    "datetime": ("string", "date-time"),
    "timestamp": ("string", "date-time"),
    "time": ("string", None),
    "object": ("object", None), "obj": ("object", None),
    "array": ("array", None), "list": ("array", None),
}

USAGE_VOCAB = {
    "mandatory": USAGE_REQUIRED, "required": USAGE_REQUIRED,
    "m": USAGE_REQUIRED, "yes": USAGE_REQUIRED,
    "conditionalmandatory": USAGE_CONDITIONAL,
    "conditional": USAGE_CONDITIONAL, "cm": USAGE_CONDITIONAL,
    "optional": USAGE_OPTIONAL, "o": USAGE_OPTIONAL, "no": USAGE_OPTIONAL,
    "": USAGE_OPTIONAL,
}

LEN_RE = re.compile(r"\(\s*(\d+)")


# --------------------------------------------------------------------------- #
# Layout
# --------------------------------------------------------------------------- #


@dataclass
class MappingLayout:
    """Where things are on a field mapping sheet."""

    header_row: int
    cols: dict                     # role -> column index
    sor_col: int
    sor_endpoint: str              # endpoint appended to the SOR column header
    sor_header: str

    def cell(self, row, role, keep_lines=False):
        c = self.cols.get(role)
        if c is None or c >= len(row):
            return ""
        return text_lines(row[c]) if keep_lines else text(row[c])


def detect_layout(rows, sheet, findings):
    """Find the header row and assign the column roles, or return ``None``."""
    for r in range(min(HEADER_SEARCH_ROWS, len(rows))):
        keys = [key(c) for c in rows[r]]
        if "parametertype" not in keys:
            continue
        taken, cols = set(), {}
        for role, hints in HEADER_ROLES:
            for i, k in enumerate(keys):
                if i in taken or not k:
                    continue
                if any(h in k for h in hints):
                    cols[role] = i
                    taken.add(i)
                    break
        if "field" not in cols:
            findings.append(Finding(
                "F002", "error", sheet, cell_ref(r + 1, 0),
                "the header row carries Parameter Type but no field name "
                "column, so there is nothing to read the hierarchy from",
                "a column headed Reusable API Field Name, or another "
                "recognised field name heading",
                "add or correct the field name column heading on row %d" % (r + 1),
            ))
            return None
        sor_col = cols.get("sor", -1)
        header = text(rows[r][sor_col]) if 0 <= sor_col < len(rows[r]) else ""
        raw = str(rows[r][sor_col]) if 0 <= sor_col < len(rows[r]) else ""
        endpoint = raw.split("\n", 1)[1].strip() if "\n" in raw else ""
        return MappingLayout(header_row=r + 1, cols=cols, sor_col=sor_col,
                             sor_endpoint=_normalise(endpoint), sor_header=header)
    return None


def _normalise(raw):
    """Strip a method prefix and normalise a path, leaving the parameter names."""
    s = re.sub(r"^\s*(GET|POST|PUT|PATCH|DELETE|HEAD|OPTIONS)\s*:?\s*", "",
               text(raw), flags=re.I).strip()
    return ("/" + s.strip("/")) if s else ""


# --------------------------------------------------------------------------- #
# Banner
# --------------------------------------------------------------------------- #


def read_banner(rows, layout, result):
    """Rows above the header, label in the first populated cell of the row."""
    banner = {}
    for r in range(0, layout.header_row - 1):
        row = rows[r]
        label = value = ""
        for i, cell in enumerate(row):
            if text(cell):
                label = text(cell)
                value = next((text(c) for c in row[i + 1:] if text(c)), "")
                break
        if not label:
            continue
        k = BANNER_KEYS.get(key(label))
        if k and not banner.get(k):
            banner[k] = value
    result.banner = banner
    return banner


# --------------------------------------------------------------------------- #
# Types and usage
# --------------------------------------------------------------------------- #


def parse_type(raw, sheet, row, col, findings):
    """``(json_type, format, max_length)``. ``None`` type when the cell is blank."""
    s = text(raw)
    if not s:
        return None, None, None
    m = LEN_RE.search(s)
    length = int(m.group(1)) if m else None
    word = key(re.sub(r"\(.*", "", s))
    hit = TYPE_MAP.get(word)
    if hit is None:
        findings.append(Finding(
            "T001", "error", sheet, cell_ref(row, col),
            "the Schema column reads %r, which is not a recognised type" % s[:40],
            "one of " + ", ".join(sorted(set(TYPE_MAP))),
            "replace it with a recognised type, optionally with a length in "
            "brackets, for example string(35)",
        ))
        return None, None, length
    return hit[0], hit[1], length


def parse_usage(raw, sheet, row, col, findings):
    s = key(raw)
    if s in USAGE_VOCAB:
        return USAGE_VOCAB[s]
    findings.append(Finding(
        "U001", "error", sheet, cell_ref(row, col),
        "Usage reads %r" % text(raw)[:40],
        "Mandatory, Conditional Mandatory, Optional, or blank for optional",
        "replace the value with one of the permitted words, or clear the cell "
        "if the attribute is optional",
    ))
    return USAGE_OPTIONAL


# --------------------------------------------------------------------------- #
# The tree, from dotted paths
# --------------------------------------------------------------------------- #


def _segments(path):
    return [p for p in text(path).split(".") if p.strip()]


def build_tree(entries, section, layout, result):
    """Turn a list of ``(row_no, segments, row)`` into a ``Node`` tree.

    A dotted path states its own parentage, so a row whose parent has not been
    declared is the dotted-path equivalent of the Level format's skipped level.
    It is reported as ``F003`` and the missing parents are created as objects,
    because refusing the row would lose a leaf the analyst clearly intended.
    """
    if not entries:
        return None
    by_path = {}
    root = None
    endpoint = layout.sor_endpoint or result.banner.get("sor_endpoint", "") or ""

    for row_no, segs, row in entries:
        node = _node_for(segs, row_no, row, layout, result, by_path, endpoint)
        if node is None:
            continue
        if node.depth == 0 and root is None:
            root = node
        elif node.depth == 0 and root is not None and node is not root:
            result.findings.append(Finding(
                "A004", "error", result.sheet,
                cell_ref(row_no, layout.cols["field"]),
                "%r is a second top-level schema in %s"
                % (node.name, section.label),
                "exactly one top-level path segment per section, naming the "
                "schema",
                "prefix %r with %r so it becomes part of that schema, or move "
                "it to its own sheet" % (node.name, root.name),
            ))
    if root is None:
        return None
    _settle_types(root, layout, result)
    return root


def _node_for(segs, row_no, row, layout, result, by_path, endpoint):
    """Create or update the node for one dotted path, making parents as needed."""
    parent = None
    for depth, seg in enumerate(segs):
        path = ".".join(segs[: depth + 1])
        node = by_path.get(path)
        if node is None:
            if not wbk.NAME_RE.match(seg):
                result.findings.append(Finding(
                    "A001", "warning", result.sheet,
                    cell_ref(row_no, layout.cols["field"]),
                    "%r is not a usable attribute name, so the row was ignored"
                    % seg[:60],
                    "a name starting with a letter or underscore, then "
                    "letters, digits, hyphens or underscores",
                    "correct the segment, or move the content out of the field "
                    "name column if it is a note rather than an attribute",
                ))
                return None
            node = Node(name=seg, depth=depth, row=row_no)
            by_path[path] = node
            if parent is not None:
                node.parent = parent
                parent.children.append(node)
            elif depth > 0:
                # Unreachable: a parent is always created before its child.
                pass
            if depth > 0 and parent is not None and parent.row != row_no \
                    and parent.json_type is None:
                parent.json_type = "object"
        parent = node

    leaf = parent
    if leaf.row != row_no and leaf.children:
        # The row restates a group already created by one of its children.
        # Fill in whatever the group's own row carries and keep the tree.
        pass
    leaf.row = leaf.row if leaf.children else row_no

    jt, fmt, length = parse_type(layout.cell(row, "dtype"), result.sheet,
                                 row_no, layout.cols.get("dtype", 0),
                                 result.findings)
    if jt is not None:
        leaf.json_type = jt
        leaf.fmt = fmt
    if length is not None:
        leaf.max_length = length
    usage_col = layout.cols.get("usage")
    if usage_col is not None:
        leaf.usage = parse_usage(layout.cell(row, "usage"), result.sheet,
                                 row_no, usage_col, result.findings)
    leaf.description = layout.cell(row, "desc", keep_lines=True) or leaf.description
    leaf.example = wbk.clean_example(layout.cell(row, "example")) or leaf.example
    leaf.remarks = layout.cell(row, "remarks") or leaf.remarks

    sor = layout.cell(row, "sor")
    if key(sor) not in NOT_USED:
        leaf.sor[endpoint] = sor
        leaf.sor_declared[endpoint] = sor
    elif key(sor) == "":
        result.blank_sor_cells.append((row_no, ".".join(segs)))
    return leaf


def _settle_types(node, layout, result):
    """Give every node a type once the tree is known, as the Level reader does."""
    for child in node.children:
        _settle_types(child, layout, result)
    if node.children:
        if node.json_type not in ("object", "array"):
            node.json_type = "object"
    elif node.json_type in ("object", "array"):
        result.findings.append(Finding(
            "A006", "warning", result.sheet,
            cell_ref(node.row, layout.cols.get("dtype", 0)),
            "%r is declared %s but no dotted path beneath it declares a "
            "member, so there is no shape to publish and the element was "
            "removed from the interface" % (node.name, node.json_type),
            "at least one row whose path extends this one, or a scalar Schema "
            "value",
            "add the members of %r as dotted paths beneath it, for example "
            "%s.<member>, or change its Schema to a scalar type"
            % (node.name, node.name),
        ))
    elif node.json_type is None:
        node.json_type = "string"
        result.findings.append(Finding(
            "T002", "warning", result.sheet,
            cell_ref(node.row, layout.cols.get("dtype", 0)),
            "%r has no Schema value and no nested rows, so it was treated as a "
            "string" % node.name,
            "a Schema value for every attribute row",
            "set the Schema of %r, for example string(35)" % node.name,
        ))


# --------------------------------------------------------------------------- #
# One sheet
# --------------------------------------------------------------------------- #


def read_sheet(grid, *, strict=False):
    """One field mapping worksheet to a :class:`SheetResult`."""
    rows = [list(r) for r in grid._rows]      # noqa: SLF001 - same package
    result = SheetResult(sheet=grid.title)
    result.blank_sor_cells = []
    # Kept for the exporter, so a row the tool does not interpret survives a
    # round trip instead of being silently dropped. Same list, not a copy.
    result.source_rows = rows
    layout = detect_layout(rows, grid.title, result.findings)
    if layout is None:
        result.status = "skipped"
        return result

    read_banner(rows, layout, result)
    endpoint = layout.sor_endpoint or _normalise(result.banner.get("sor_endpoint", ""))
    result.layout = Layout(
        header_row=layout.header_row,
        level_cols=[layout.cols["field"]],
        label_cols=[layout.cols.get("ptype", 0)],
        roles={"example": layout.cols.get("example", -1),
               "type": layout.cols.get("dtype", -1),
               "usage": layout.cols.get("usage", -1),
               "description": layout.cols.get("desc", -1)},
        sor_cols=[SorColumn(index=layout.sor_col, endpoint=endpoint,
                            header=layout.sor_header)])
    result.endpoints = [endpoint]
    result.mapping_layout = layout

    if not result.banner.get("sor_endpoint"):
        result.fail(Finding(
            "R003", "error", grid.title, cell_ref(1, 0),
            "no SOR API Endpoint is stated in the banner, and it is the only "
            "place the tool reads the SOR endpoint from",
            "the SOR endpoint in the SOR API Endpoint banner row",
            "fill in the SOR API Endpoint row. The endpoint appended to the "
            "SOR API Field Name column header is a label and is not read for "
            "this purpose.",
        ))

    # Walk the grid, splitting rows into sections by the Parameter Type column.
    sections, current, seen = {}, None, []
    entries = {SECTION_REQUEST: [], SECTION_RESPONSE: []}
    ignored = 0
    blank_run = 0
    for r in range(layout.header_row, len(rows)):
        row = rows[r]
        row_no = r + 1
        if not any(text(c) for c in row):
            blank_run += 1
            if blank_run >= wbk.STRUCTURE_BREAK:
                break
            continue
        blank_run = 0
        ptype = layout.cell(row, "ptype")
        field_path = layout.cell(row, "field")
        pk = key(ptype)

        if ptype and not field_path:
            kind = SECTION_BY_TYPE.get(pk)
            if kind:
                if kind in sections:
                    result.findings.append(Finding(
                        "S002", "warning", grid.title,
                        cell_ref(row_no, layout.cols.get("ptype", 0)),
                        "a second %s section labelled %r"
                        % ("request" if kind == SECTION_REQUEST else "response",
                           ptype),
                        "one request section and one response section per sheet",
                        "merge it into the first one, or move it to its own "
                        "sheet",
                    ))
                else:
                    sections[kind] = Section(kind=kind, label=ptype, row=row_no)
                current = kind
            else:
                current = None
                seen.append((ptype, row_no))
            continue

        if not field_path:
            continue
        if pk in TYPE_IGNORED:
            ignored += 1
            continue
        if pk and pk not in TYPE_BODY:
            ignored += 1
            continue

        # Parameter Type is authoritative; the section banner corroborates.
        if current is None:
            result.findings.append(Finding(
                "F001", "warning", grid.title,
                cell_ref(row_no, layout.cols.get("ptype", 0)),
                "%r sits under no recognised section, so it was read into the "
                "request" % field_path[:50],
                "every body row to follow a Request Body or Response Body "
                "section row",
                "add a Request Body or Response Body row in the Parameter "
                "Type column above this row",
            ))
            current = SECTION_REQUEST
            sections.setdefault(current,
                                Section(kind=current, label="Request Body",
                                        row=row_no))
        entries[current].append((row_no, _segments(field_path), row))

    result.ignored_rows = ignored
    result.ignored_sections = seen

    for kind, code, label in ((SECTION_REQUEST, "S001", "Request Body"),
                              (SECTION_RESPONSE, "S003", "Response Body")):
        if kind not in sections:
            result.fail(Finding(
                code, "error", grid.title, cell_ref(layout.header_row + 1, 0),
                "this sheet has no %s section" % label,
                "a %s row in the Parameter Type column above the schema rows"
                % label,
                "add a %s row in column %s" % (label,
                                               col_letter(layout.cols.get("ptype", 0))),
            ))

    if result.status == "failed":
        return result

    for kind, attr in ((SECTION_REQUEST, "request"), (SECTION_RESPONSE, "response")):
        section = sections.get(kind)
        if section is None:
            continue
        section.root = build_tree(entries[kind], section, layout, result)
        if section.root is None and entries[kind]:
            result.findings.append(Finding(
                "S004", "warning", grid.title, cell_ref(section.row, 0),
                "the %s section has rows beneath it but no usable schema was "
                "read from them" % section.label,
                "a top-level path segment naming the schema, with its "
                "attributes as dotted paths beneath it",
                "check the field name column for this section",
            ))
        setattr(result, attr, section)

    if strict:
        for row_no, path in result.blank_sor_cells:
            result.findings.append(Finding(
                "F004", "warning", grid.title,
                cell_ref(row_no, layout.sor_col),
                "%r leaves the SOR API Field Name cell empty rather than "
                "stating N/A" % path[:60],
                "either an SOR field name, or N/A to record that the System of "
                "Record does not supply it",
                "write N/A where the element is genuinely unsupported, so an "
                "unfinished row can be told from a finished one",
            ))
        for f in result.findings:
            if f.severity == "warning":
                f.severity = "error"
        if result.errors:
            result.status = "failed"
    return result


def read_workbook(path, *, strict=False, max_rows=20000):
    """Every sheet of a field mapping workbook, in order."""
    grids = core.load_grids(path, max_rows=max_rows)
    out = []
    for g in grids:
        try:
            out.append(read_sheet(g, strict=strict))
        except Exception as exc:                               # noqa: BLE001
            r = SheetResult(sheet=g.title)
            r.fail(Finding(
                "X001", "error", g.title, "",
                "the reader stopped on this sheet: %s: %s"
                % (type(exc).__name__, exc),
                "a sheet the reader can parse end to end",
                "send this sheet to the maintainer. The rest of the workbook "
                "was processed.",
            ))
            out.append(r)
    return out


__all__ = ["BANNER_KEYS", "HEADER_ROLES", "SECTION_BY_TYPE", "TYPE_IGNORED",
           "TYPE_BODY", "NOT_USED", "TYPE_MAP", "USAGE_VOCAB",
           "MappingLayout", "detect_layout", "read_banner", "parse_type",
           "parse_usage", "build_tree", "read_sheet", "read_workbook"]
