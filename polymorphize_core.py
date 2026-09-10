#!/usr/bin/env python3
"""
polymorphize_core.py  —  SOR Polymorphizer engine
=================================================

Rebuilt engine (v5.0) that eliminates OpenAPI/Swagger attributes with no
System-of-Record (SOR) connection, using polymorphism (``allOf`` sub-typing,
optionally with a ``discriminator``). Shared by the command-line wrapper
(``polymorphize_cli.py``), the Windows GUI (``polymorphize_gui.py``) and the
batch runner (``polymorphize_batch.py``).

Why this was rebuilt (defects corrected in v5.0)
------------------------------------------------
D1  MAPPING LAYOUT WAS ASSUMED, NOT DETECTED.
    v4 read fixed columns (levels C..G, data type I, SOR J). Real mapping
    workbooks put the header row part-way down the sheet and carry the field
    name as ONE dotted path column. On the CASA v1.4.0 mapping, v4 therefore
    read 'Usage'..'Description' as the name and an empty column as the SOR
    field, so 945 of 986 attributes were logged 'absent-from-mapping' and the
    payload was emptied. v5 locates the header row and maps columns by label,
    supporting both the dotted-path and the legacy Level 1..n layouts.

D2  MATCHING WAS A FLAT, GLOBAL LEAF NAMESPACE.
    v4 keyed the mapping on the bare leaf name only, so the same leaf in two
    operations collided and qualified mapping rows could never be used. v5
    builds the set of usage paths for every schema (following $ref) and
    resolves most-specific-first: exact path, then path suffixes, then leaf.

D3  'NOT FOUND' WAS TREATED AS 'DELETE'.
    v4 defaulted to eliminating anything it could not find. A lookup failure
    was therefore indistinguishable from a genuine 'not used in SOR'. v5
    defaults to KEEP AND FLAG; elimination requires a positive not-in-SOR
    signal. ``on_unmatched='drop'`` restores the old behaviour explicitly.

D4  EMPTIED OBJECTS WERE LEFT AS SHELLS.
    v4 removed leaf attributes but left the parent object with
    ``properties: {}`` and every $ref still pointing at it, so payloads
    rendered as {}. v5 cascade-prunes: an object that loses all of its
    properties is removed, along with the references to it, the array
    properties whose item type died, and the matching ``required`` entries.

D5  DE-DUPLICATION RAN AFTER ELIMINATION AND COMPARED KEY SETS ONLY.
    Two schemas that differed in the input became 'identical' once both were
    gutted, and were merged. On CASA this collapsed
    AccountLimitResponse.PaymentTransaction into PaymentTransaction and
    silently lost paymentTransactionStatus. v5 computes equivalence from the
    ORIGINAL document, compares shape rather than key sets, and re-verifies
    equivalence at apply time.

D7  TEXT ENCODING WAS LEFT TO THE PLATFORM.
    The output document was written with the platform default encoding and the
    HTML change report then read it back as hard-coded UTF-8. On Windows that
    combination raised UnicodeDecodeError on the em dash in the CASA
    description (byte 0x97), and it also meant a specification saved as ANSI
    could not be read at all. v5.1 sniffs the input encoding (BOM, UTF-8, then
    the Windows single-byte encodings), always writes UTF-8, and treats report
    generation as non-fatal so a cosmetic failure never discards a completed
    transform.

D6  NOTHING CHECKED THE OUTPUT.
    v5 adds circuit breakers (minimum mapping match rate, maximum total
    attribute loss) and post-run assertions (no empty object schema, no
    dangling $ref, no required-without-property). A failing run raises rather
    than writing a plausible-looking but empty specification.

Dependencies: ruamel.yaml, openpyxl
"""
import copy
import difflib
import html as _html
import os
import re

import openpyxl
from ruamel.yaml import YAML
from ruamel.yaml.comments import CommentedMap, CommentedSeq

__version__ = "6.7"

CONTAINER_TYPES = {"object", "array"}

# Legacy fixed-column layout. Still honoured when 'auto' detection is turned
# off, or when a sheet has no recognisable header row.
DEFAULT_SHEET_CFG = {
    "level_first_col": 3,   # column C  (Level 1)
    "level_last_col": 7,    # column G  (Level 5)
    "dtype_col": 9,         # column I  (Data Type)
    "sor_col": 10,          # column J  (SOR Field)
    "not_used_text": "Not Used In SOR",
    "auto_layout": True,    # detect the header row and columns by label
    "header_scan_rows": 40,
}

# Cell values that mean 'this attribute has no SOR field'.
NOT_USED_TOKENS = {
    "", "na", "n/a", "none", "null", "nil", "-", "--", "tbd", "tba",
    "notused", "notusedinsor", "notusedinthesor", "notapplicable",
    "nomapping", "nosormapping", "toboconfirmed", "tobeconfirmed",
}

# Rows whose 'Parameter Type' cell is only a section banner.
SECTION_MARKERS = {
    "requestparameter", "requestbody", "requestheader", "responsebody",
    "responseheader", "responseparameter", "queryparameter", "pathparameter",
    "requestqueryparameter", "responsecode", "header", "query", "path",
}

# Header labels, in assignment priority order. The first unassigned column
# whose normalised header contains one of the hints takes the role.
HEADER_ROLES = [
    ("ptype", ("parametertype", "paramtype", "type of parameter")),
    ("sor", ("sorapifieldname", "sorfieldname", "sorfield", "sorattribute",
             "sorapifield", "sormapping", "sorname", "sor")),
    ("field", ("reusableapifieldname", "apifieldname", "fieldname", "attributename",
               "elementname", "attribute", "element", "field", "name")),
    ("usage", ("usage", "mandatoryoptional", "cardinality", "required")),
    ("dtype", ("schema", "datatype", "type", "format")),
    ("example", ("examples", "example", "samplevalue", "sample")),
    ("desc", ("description", "definition")),
    ("remarks", ("remarks", "comments", "comment", "notes")),
]


class MappingCoverageError(RuntimeError):
    """Raised when the mapping matched so little of the spec that the run is
    almost certainly a join failure rather than a filtering decision."""


class ExcessiveLossError(RuntimeError):
    """Raised when a run would strip more of the specification than allowed."""


class OutputAssertionError(RuntimeError):
    """Raised when the transformed document fails a post-run assertion."""


# --------------------------------------------------------------------------- #
# Normalisation helpers
# --------------------------------------------------------------------------- #
def norm(name):
    """Normalise a name for matching (case / space / punctuation insensitive)."""
    if name is None:
        return ""
    return re.sub(r"[^a-z0-9]", "", str(name).replace("\xa0", " ").lower())


def _clean_cell(value):
    """Excel cells in these workbooks are padded with non-breaking spaces."""
    if value is None:
        return ""
    return str(value).replace("\xa0", " ").strip()


def split_path(path):
    """'A.B[].c' -> ['A', 'B', 'c'] (array markers dropped)."""
    raw = str(path or "").replace("\xa0", " ")
    raw = re.sub(r"\[[^\]]*\]", "", raw)
    return [seg.strip() for seg in raw.split(".") if seg.strip()]


def norm_path(segments):
    """Normalise a list of path segments into a single comparable key."""
    if isinstance(segments, str):
        segments = split_path(segments)
    return ".".join(norm(s) for s in segments if norm(s))


def is_not_used(sor_value, extra_tokens=()):
    """True when the SOR cell carries no usable field reference."""
    text = _clean_cell(sor_value)
    key = norm(text)
    if key in NOT_USED_TOKENS:
        return True
    for tok in extra_tokens:
        t = norm(tok)
        if t and (key == t or key.startswith(t)):
            return True
    return False


# --------------------------------------------------------------------------- #
# Workbook loading
# --------------------------------------------------------------------------- #
# A heavily formatted mapping workbook takes about a minute to open in
# openpyxl's normal mode and a tenth of a second in read-only mode, because
# normal mode materialises every style. Everything here only reads, so each
# sheet is pulled once into a dense grid and indexed from memory. The shim
# keeps the familiar ``ws.cell(r, c).value`` call shape.


class _CellValue:
    __slots__ = ("value",)

    def __init__(self, value):
        self.value = value


class SheetGrid:
    """An in-memory worksheet: 1-based indexing, ``.cell(r, c).value``."""

    __slots__ = ("title", "_rows", "max_row", "max_column")

    def __init__(self, title, rows):
        self.title = title
        self._rows = rows
        self.max_row = len(rows)
        self.max_column = max((len(r) for r in rows), default=0)

    def value(self, row, col):
        if 1 <= row <= self.max_row:
            r = self._rows[row - 1]
            if 1 <= col <= len(r):
                return r[col - 1]
        return None

    def cell(self, row, col):
        return _CellValue(self.value(row, col))


def _row_is_blank(row):
    return all(v is None or (isinstance(v, str) and not v.strip()) for v in row)


def load_grids(path, max_rows=20000, blank_run=300):
    """Every worksheet of a workbook as a SheetGrid, read once, read-only.

    Real mapping workbooks declare enormous sheet dimensions because stray
    formatting sits far below the data: the reference CASA workbook reports
    3.1 million rows across 26 sheets. Reading to the declared dimension takes
    over half a minute, so the read stops after `blank_run` consecutive empty
    rows and trailing blanks are trimmed.
    """
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    grids = []
    try:
        for ws in wb.worksheets:
            rows, blanks = [], 0
            for raw in ws.iter_rows(values_only=True):
                row = list(raw)
                if _row_is_blank(row):
                    blanks += 1
                    if blanks >= blank_run:
                        break
                else:
                    blanks = 0
                rows.append(row)
                if len(rows) >= max_rows:
                    break
            while rows and _row_is_blank(rows[-1]):
                rows.pop()
            grids.append(SheetGrid(ws.title, rows))
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return grids


# --------------------------------------------------------------------------- #
# D1 — mapping layout detection
# --------------------------------------------------------------------------- #
def detect_layout(ws, cfg):
    """Locate the header row and resolve columns by label.

    Returns a layout dict:
        {'header_row': int, 'cols': {role: col}, 'levels': [col, ...],
         'mode': 'path' | 'levels' | 'fallback', 'sheet': str}
    """
    max_scan = min(int(cfg.get("header_scan_rows", 40)), ws.max_row or 1)
    best = None
    for r in range(1, max_scan + 1):
        headers = {}
        for c in range(1, (ws.max_column or 1) + 1):
            text = norm(ws.cell(r, c).value)
            if text:
                headers[c] = text
        if not headers:
            continue
        levels = [c for c, t in headers.items() if re.fullmatch(r"level\d+", t)]
        cols, used = {}, set()
        for role, hints in HEADER_ROLES:
            for c in sorted(headers):
                if c in used or c in levels:
                    continue
                if any(h in headers[c] for h in hints):
                    cols[role] = c
                    used.add(c)
                    break
        score = 0
        if "sor" in cols:
            score += 3
        if levels:
            score += 2 + min(len(levels), 5) * 0.1
        elif "field" in cols:
            score += 2
        if "dtype" in cols:
            score += 1
        if "usage" in cols:
            score += 0.5
        if score >= 4 and (best is None or score > best[0]):
            best = (score, {"header_row": r, "cols": cols, "levels": sorted(levels),
                            "mode": "levels" if levels else "path", "sheet": ws.title})
    if best:
        return best[1]
    # Nothing recognisable — fall back to the configured fixed columns.
    return {
        "header_row": 1,
        "cols": {"dtype": int(cfg["dtype_col"]), "sor": int(cfg["sor_col"])},
        "levels": list(range(int(cfg["level_first_col"]), int(cfg["level_last_col"]) + 1)),
        "mode": "fallback",
        "sheet": ws.title,
    }


def _row_path(ws, r, layout):
    """Extract the attribute's dotted path from one row."""
    cols = layout["cols"]
    if layout["levels"]:
        segs = [_clean_cell(ws.cell(r, c).value) for c in layout["levels"]]
        segs = [s for s in segs if s and not s.lower().startswith("level ")]
        return segs
    col = cols.get("field")
    if not col:
        return []
    return split_path(_clean_cell(ws.cell(r, col).value))


# --------------------------------------------------------------------------- #
# Mapping index
# --------------------------------------------------------------------------- #
class SorMapping:
    """Path-aware index of a SOR field-mapping workbook (D1 + D2)."""

    def __init__(self):
        self.rows = []           # every attribute row, in sheet order
        self.by_path = {}        # normalised full path  -> entry
        self.by_suffix = {}      # normalised suffix     -> entry
        self.by_leaf = {}        # normalised leaf name  -> entry
        self.layouts = []        # one detected layout per sheet
        self.stats = {"sheets": 0, "rows_read": 0, "attribute_rows": 0,
                      "section_rows": 0, "mapped": 0, "unmapped": 0}

    # -- construction ----------------------------------------------------- #
    def _better(self, existing, candidate):
        if existing is None:
            return True
        if candidate["status"] == "mapped" and existing["status"] != "mapped":
            return True
        return False

    def _register(self, entry):
        segs = entry["segments"]
        key = norm_path(segs)
        if not key:
            return
        if self._better(self.by_path.get(key), entry):
            self.by_path[key] = entry
        for length in range(2, len(segs) + 1):
            skey = norm_path(segs[-length:])
            if self._better(self.by_suffix.get(skey), entry):
                self.by_suffix[skey] = entry
        leaf = norm(segs[-1])
        if leaf and self._better(self.by_leaf.get(leaf), entry):
            self.by_leaf[leaf] = entry

    def add_workbook(self, path, cfg):
        extra_not_used = [cfg.get("not_used_text", "")]
        for ws in load_grids(path):
            layout = (detect_layout(ws, cfg) if cfg.get("auto_layout", True)
                      else detect_layout_fixed(ws, cfg))
            self.layouts.append(layout)
            self.stats["sheets"] += 1
            cols = layout["cols"]
            for r in range(layout["header_row"] + 1, (ws.max_row or 0) + 1):
                self.stats["rows_read"] += 1
                segs = _row_path(ws, r, layout)
                if not segs:
                    self.stats["section_rows"] += 1
                    continue
                if len(segs) == 1 and norm(segs[0]) in SECTION_MARKERS:
                    self.stats["section_rows"] += 1
                    continue
                dtype = _clean_cell(ws.cell(r, cols["dtype"]).value) if cols.get("dtype") else ""
                sor_raw = _clean_cell(ws.cell(r, cols["sor"]).value) if cols.get("sor") else ""
                mapped = not is_not_used(sor_raw, extra_not_used)
                entry = {
                    "segments": segs,
                    "path": ".".join(segs),
                    "leaf": segs[-1],
                    "dtype": dtype,
                    "container": norm(dtype) in {norm(t) for t in CONTAINER_TYPES},
                    "sor": sor_raw if mapped else "",
                    "status": "mapped" if mapped else "unmapped",
                    "sheet": ws.title,
                    "row": r,
                }
                self.rows.append(entry)
                self.stats["attribute_rows"] += 1
                self.stats["mapped" if mapped else "unmapped"] += 1
                self._register(entry)
        return self

    # -- resolution (D2) --------------------------------------------------- #
    def resolve(self, candidate_paths, leaf):
        """Resolve an attribute most-specific-first.

        `candidate_paths` is a list of segment-lists (the attribute's full
        path under each place its owning schema is used). Returns
        (entry_or_None, strategy, keys_tried).
        """
        tried = []

        # tier 1 — exact qualified path
        hits = []
        for segs in candidate_paths:
            key = norm_path(segs)
            if not key:
                continue
            tried.append(key)
            e = self.by_path.get(key)
            if e:
                hits.append(e)
        if hits:
            return self._best(hits), "exact-path", tried

        # tier 2 — longest path suffix, longest first
        longest = max((len(s) for s in candidate_paths), default=0)
        for length in range(min(longest, 6), 1, -1):
            hits = []
            for segs in candidate_paths:
                if len(segs) < length:
                    continue
                key = norm_path(segs[-length:])
                tried.append(key)
                e = self.by_suffix.get(key)
                if e:
                    hits.append(e)
            if hits:
                return self._best(hits), f"suffix-{length}", tried

        # tier 3 — bare leaf name
        lkey = norm(leaf)
        tried.append(lkey)
        e = self.by_leaf.get(lkey)
        if e:
            return e, "leaf", tried
        return None, "none", tried

    @staticmethod
    def _best(hits):
        for h in hits:
            if h["status"] == "mapped":
                return h
        return hits[0]

    def nearest_keys(self, key, n=3):
        return difflib.get_close_matches(key, list(self.by_path.keys()), n=n, cutoff=0.6)

    # -- legacy view ------------------------------------------------------- #
    def as_flat_index(self):
        """v4-compatible {normalised-leaf: {...}} view, for old callers."""
        return {k: {"status": v["status"], "sor": v["sor"], "leaf": v["leaf"]}
                for k, v in self.by_leaf.items()}


def detect_layout_fixed(ws, cfg):
    return {
        "header_row": 1,
        "cols": {"dtype": int(cfg["dtype_col"]), "sor": int(cfg["sor_col"])},
        "levels": list(range(int(cfg["level_first_col"]), int(cfg["level_last_col"]) + 1)),
        "mode": "fixed",
        "sheet": ws.title,
    }


def load_mapping(paths, sheet_cfg=None):
    """Load one or more mapping workbooks into a single SorMapping."""
    cfg = dict(DEFAULT_SHEET_CFG, **(sheet_cfg or {}))
    if isinstance(paths, str):
        paths = [paths]
    m = SorMapping()
    for p in paths:
        m.add_workbook(p, cfg)
    return m


def build_sor_index(paths, sheet_cfg=None):
    """Backwards-compatible flat index (kept so v4 callers still work)."""
    return load_mapping(paths, sheet_cfg).as_flat_index()


# --------------------------------------------------------------------------- #
# SOR API YAML files -> the field names the System-of-Record exposes
# --------------------------------------------------------------------------- #
def collect_sor_fields(paths):
    """Normalised property names exposed by one or more SOR OpenAPI files."""
    if isinstance(paths, str):
        paths = [paths]
    yaml = YAML(typ="safe")
    fields = set()

    def walk(node):
        if isinstance(node, dict):
            props = node.get("properties")
            if isinstance(props, dict):
                for k in props:
                    fields.add(norm(k))
            for v in node.values():
                walk(v)
        elif isinstance(node, list):
            for v in node:
                walk(v)

    for p in paths:
        text, _enc = read_text(p)
        walk(yaml.load(text))
    fields.discard("")
    return fields


def _sor_tokens(field_value):
    """Split a mapping's SOR-field cell into candidate field tokens.

    SOR fields are written as dotted paths ('depAcctInfo.custRelation.custPermId'),
    so every path segment is a candidate as well as the whole string.
    """
    out = []
    for chunk in re.split(r"[,/;|]+|\s{2,}", str(field_value)):
        if norm(chunk):
            out.append(norm(chunk))
        for seg in str(chunk).split("."):
            if norm(seg):
                out.append(norm(seg))
    return out


def _field_in_sor(field_value, sor_fields):
    toks = _sor_tokens(field_value)
    return any(t in sor_fields for t in toks) if toks else False


def validate_mapping(mapping, sor_fields):
    """Cross-check every mapped attribute's SOR field against the SOR YAML(s)."""
    confirmed, missing = [], []
    seen = set()
    rows = mapping.rows if isinstance(mapping, SorMapping) else []
    if not rows and isinstance(mapping, dict):        # legacy flat index
        rows = [{"leaf": v.get("leaf", k), "sor": v.get("sor", ""),
                 "status": v.get("status")} for k, v in mapping.items()]
    for entry in rows:
        if entry.get("status") != "mapped":
            continue
        pair = (entry.get("path", entry.get("leaf", "")), entry.get("sor", ""))
        if pair in seen:
            continue
        seen.add(pair)
        (confirmed if _field_in_sor(pair[1], sor_fields) else missing).append(pair)
    confirmed.sort()
    missing.sort()
    return {"confirmed": confirmed, "missing": missing, "field_count": len(sor_fields)}


# --------------------------------------------------------------------------- #
# Schema helpers
# --------------------------------------------------------------------------- #
def is_container(node):
    if not isinstance(node, dict):
        return False
    if "$ref" in node or "properties" in node or "items" in node:
        return True
    if "allOf" in node or "oneOf" in node or "anyOf" in node:
        return True
    return node.get("type") in CONTAINER_TYPES


def plain(node):
    if isinstance(node, dict):
        return {k: plain(v) for k, v in node.items()}
    if isinstance(node, list):
        return [plain(v) for v in node]
    return node


def drop_from_required(schema, name):
    req = schema.get("required")
    if isinstance(req, list) and name in req:
        req.remove(name)
        if not req:
            del schema["required"]


def _set_comment(node, text, indent=0):
    try:
        node.yaml_set_start_comment(text, indent=indent)
    except Exception:
        pass


def ref_name(node):
    if isinstance(node, dict):
        ref = node.get("$ref")
        if isinstance(ref, str) and ref.startswith("#/components/schemas/"):
            return ref.split("/")[-1]
    return None


def _resolve_local(doc, ref):
    if not isinstance(ref, str) or not ref.startswith("#/"):
        return None
    node = doc
    for part in ref[2:].split("/"):
        if not isinstance(node, dict) or part not in node:
            return None
        node = node[part]
    return node


# --------------------------------------------------------------------------- #
# D2 — usage paths for every schema
# --------------------------------------------------------------------------- #
def _content_schema_ref(doc, holder, depth=0):
    """Schema name behind a requestBody / response holder, following $ref."""
    if not isinstance(holder, dict) or depth > 4:
        return None
    if "$ref" in holder and not str(holder["$ref"]).startswith("#/components/schemas/"):
        return _content_schema_ref(doc, _resolve_local(doc, holder["$ref"]), depth + 1)
    content = holder.get("content")
    if isinstance(content, dict):
        for _mt, mv in content.items():
            sch = mv.get("schema") if isinstance(mv, dict) else None
            name = ref_name(sch)
            if name:
                return name
            if isinstance(sch, dict) and ref_name(sch.get("items")):
                return ref_name(sch["items"])
    return None


def operation_body_schemas(doc):
    """(request_schema_names, success_response_schema_names)."""
    requests, responses = [], []
    for _p, ops in (doc.get("paths") or {}).items():
        if not isinstance(ops, dict):
            continue
        for _m, op in ops.items():
            if not isinstance(op, dict):
                continue
            r = _content_schema_ref(doc, op.get("requestBody", {}))
            if r and r not in requests:
                requests.append(r)
            for code, resp in (op.get("responses") or {}).items():
                if str(code).startswith("2"):
                    r = _content_schema_ref(doc, resp)
                    if r and r not in responses:
                        responses.append(r)
    return requests, responses


def build_usage_paths(doc, max_depth=10, max_paths_per_schema=24, max_visits=250000):
    """Map every named schema to the dotted paths at which it is used.

    Array levels do not add a segment, matching how mapping workbooks write
    'Parent.ArrayProp.child' rather than 'Parent.ArrayProp[].child'. The walk
    is breadth-first with a global visited set, a depth bound and a per-schema
    cap, so a widely shared or self-referencing schema cannot blow up.
    """
    from collections import deque

    schemas = (doc.get("components") or {}).get("schemas") or {}
    requests, responses = operation_body_schemas(doc)
    roots = list(dict.fromkeys(list(requests) + list(responses)))
    if not roots:
        roots = list(schemas.keys())

    # child references are a property of the schema, not of the path it is
    # reached by, so resolve them once per schema and compose.
    rel = {name: _child_refs(sc, []) for name, sc in schemas.items()}

    usages = {name: [] for name in schemas}
    seen = set()
    dq = deque((r, (r,)) for r in roots if r in schemas)
    visits = 0
    while dq and visits < max_visits:
        name, path = dq.popleft()
        if (name, path) in seen:
            continue
        seen.add((name, path))
        visits += 1
        bucket = usages.setdefault(name, [])
        if len(bucket) >= max_paths_per_schema:
            continue
        bucket.append(list(path))
        if len(path) >= max_depth:
            continue
        for child_name, child_rel in rel.get(name, ()):
            if child_name in schemas:
                dq.append((child_name, path + tuple(child_rel)))

    for name in schemas:                     # unreachable schemas stand alone
        if not usages.get(name):
            usages[name] = [[name]]
    return usages


def _child_refs(node, prefix, key=None):
    """Yield (schema_name, path) for every $ref reachable inside `node`."""
    out = []
    if not isinstance(node, dict):
        return out
    props = node.get("properties")
    if isinstance(props, dict):
        for pname, pval in props.items():
            child = prefix + [pname]
            name = ref_name(pval)
            if name:
                out.append((name, child))
                continue
            if isinstance(pval, dict):
                iname = ref_name(pval.get("items"))
                if iname:
                    out.append((iname, child))
                    continue
                out.extend(_child_refs(pval, child))
    items = node.get("items")
    if isinstance(items, dict):
        name = ref_name(items)
        if name:
            out.append((name, prefix))
        else:
            out.extend(_child_refs(items, prefix))
    for comb in ("allOf", "oneOf", "anyOf"):
        seq = node.get(comb)
        if isinstance(seq, list):
            for sub in seq:
                name = ref_name(sub)
                if name:
                    out.append((name, prefix))
                else:
                    out.extend(_child_refs(sub, prefix))
    return out


# --------------------------------------------------------------------------- #
# D3 — pruning, keep-and-flag by default
# --------------------------------------------------------------------------- #
def prune_document(doc, mapping, *, on_unmatched="keep", protected=("ErrorResponse",),
                   sor_fields=None, require_sor_field=False, log=lambda m: None):
    """Eliminate leaf attributes with no SOR connection.

    Returns a diagnostics dict. Elimination requires a positive not-in-SOR
    signal unless `on_unmatched='drop'` is passed explicitly (D3).
    """
    schemas = (doc.get("components") or {}).get("schemas") or {}
    usages = build_usage_paths(doc)
    protected = set(protected or ())

    dropped, kept, unmatched, per_schema = [], [], [], {}

    def decide(candidates, leaf):
        entry, strategy, tried = mapping.resolve(candidates, leaf)
        if entry is None:
            return (on_unmatched == "drop"), "absent-from-mapping", strategy, tried, None
        if entry["status"] == "unmapped":
            return True, "not-used-in-SOR", strategy, tried, entry
        if require_sor_field and sor_fields is not None and \
                not _field_in_sor(entry.get("sor", ""), sor_fields):
            return True, f"SOR field '{entry.get('sor', '')}' not in SOR YAML", strategy, tried, entry
        return False, None, strategy, tried, entry

    def walk(schema, prefixes, display, owner):
        if not isinstance(schema, dict):
            return
        props = schema.get("properties")
        if isinstance(props, dict):
            for pname in list(props.keys()):
                pval = props[pname]
                child_prefixes = [p + [pname] for p in prefixes]
                if is_container(pval) and not is_scalar_array(pval):
                    if ref_name(pval) is None and ref_name((pval or {}).get("items")) is None:
                        walk(pval, child_prefixes, f"{display}.{pname}", owner)
                    continue
                remove, reason, strategy, tried, entry = decide(child_prefixes, pname)
                rec = {"path": f"{display}.{pname}", "attribute": pname,
                       "strategy": strategy, "keys_tried": tried[:4],
                       "sor": (entry or {}).get("sor", ""),
                       "mapping_row": f"{(entry or {}).get('sheet','')}!{(entry or {}).get('row','')}"
                                      if entry else "", "owner": owner}
                st = per_schema.setdefault(owner, {"total": 0, "removed": 0})
                st["total"] += 1
                # 'not found in the mapping' is recorded independently of what
                # is then done with it, so the match rate stays a measure of
                # the lookup and never of the elimination policy.
                if entry is None:
                    unmatched.append(rec)
                if remove:
                    del props[pname]
                    drop_from_required(schema, pname)
                    st["removed"] += 1
                    rec["reason"] = reason
                    dropped.append(rec)
                else:
                    rec["reason"] = ("absent-from-mapping (kept and flagged)"
                                     if entry is None else "mapped to a SOR field")
                    kept.append(rec)
        if isinstance(schema.get("items"), dict) and ref_name(schema["items"]) is None:
            walk(schema["items"], prefixes, display, owner)
        for comb in ("allOf", "oneOf", "anyOf"):
            seq = schema.get(comb)
            if isinstance(seq, list):
                for sub in seq:
                    if isinstance(sub, dict) and ref_name(sub) is None:
                        walk(sub, prefixes, display, owner)

    for name in list(schemas.keys()):
        if name in protected:
            continue
        prefixes = usages.get(name) or [[name]]
        walk(schemas[name], prefixes, name, name)

    total = len(dropped) + len(kept)
    matched = total - len(unmatched)
    diagnostics = {
        "attributes_total": total,
        "attributes_matched": matched,
        "attributes_unmatched": len(unmatched),
        "match_rate": (matched / total) if total else 1.0,
        "dropped": dropped,
        "kept": kept,
        "unmatched": unmatched,
        "per_schema": per_schema,
        "strategies": _count_by(kept + dropped, "strategy"),
    }
    log(f"  Resolved {matched}/{total} attribute(s) against the mapping "
        f"({diagnostics['match_rate'] * 100:.1f}%).")
    return diagnostics


def _count_by(records, key):
    out = {}
    for r in records:
        out[r.get(key, "?")] = out.get(r.get(key, "?"), 0) + 1
    return dict(sorted(out.items(), key=lambda kv: -kv[1]))


# --------------------------------------------------------------------------- #
# D4 — cascade pruning
# --------------------------------------------------------------------------- #
def _is_empty_object(node):
    """True only for an object that CARRIES a properties map which is now empty.

    An object that never declared properties in the source is free-form by
    design and is left exactly as it was found — the tool only removes shells
    that its own pruning created.
    """
    if not isinstance(node, dict):
        return False
    if any(k in node for k in ("allOf", "oneOf", "anyOf", "$ref", "enum",
                               "additionalProperties", "items")):
        return False
    if node.get("type") not in (None, "object"):
        return False
    props = node.get("properties")
    return isinstance(props, dict) and len(props) == 0


def _body_schema_names(doc):
    """Schema names referenced directly by an operation body or a shared
    requestBody / response. These must never be deleted outright."""
    names = set()
    requests, responses = operation_body_schemas(doc)
    names.update(requests)
    names.update(responses)
    comps = doc.get("components") or {}
    for section in ("requestBodies", "responses"):
        for _k, holder in (comps.get(section) or {}).items():
            n = _content_schema_ref(doc, holder)
            if n:
                names.add(n)
    return names


def cascade_prune(doc, *, protected=("ErrorResponse",), max_rounds=25, log=lambda m: None):
    """Remove objects emptied by pruning, and every reference to them.

    An array property whose item type died is removed with it. `required`
    entries are cleaned. Schemas that form an operation body are never
    deleted; they are reported instead, because an empty request or response
    body means the mapping, not the specification, needs attention.
    """
    schemas = (doc.get("components") or {}).get("schemas") or {}
    keep_names = set(protected or ()) | _body_schema_names(doc)
    removed_schemas, removed_props, orphan_bodies = [], [], []

    for _round in range(max_rounds):
        dead = {n for n, sc in schemas.items()
                if n not in keep_names and _is_empty_object(sc)}
        stranded = {n for n, sc in schemas.items()
                    if n in keep_names and _is_empty_object(sc)}
        for n in sorted(stranded):
            if n not in orphan_bodies:
                orphan_bodies.append(n)
        inline_removed = _strip_empty_inline(schemas, removed_props)
        if not dead and not inline_removed:
            break
        for n in sorted(dead):
            del schemas[n]
            removed_schemas.append(n)
        if dead:
            _strip_refs(doc, dead, removed_props)

    if removed_schemas:
        log(f"  Cascade: removed {len(removed_schemas)} object(s) left with no attributes.")
    if removed_props:
        log(f"  Cascade: removed {len(removed_props)} reference(s) to removed objects.")
    if orphan_bodies:
        log(f"  WARNING: {len(orphan_bodies)} operation body schema(s) would be empty: "
            + ", ".join(orphan_bodies))
    return {"removed_schemas": removed_schemas, "removed_properties": removed_props,
            "empty_body_schemas": orphan_bodies}


def _strip_empty_inline(schemas, removed_props):
    """Delete inline object properties that were emptied by pruning."""
    n = 0
    for owner, sc in schemas.items():
        n += _strip_empty_inline_node(sc, owner, removed_props)
    return n


def _strip_empty_inline_node(node, display, removed_props):
    if not isinstance(node, dict):
        return 0
    n = 0
    props = node.get("properties")
    if isinstance(props, dict):
        for pname in list(props.keys()):
            pval = props[pname]
            if not isinstance(pval, dict):
                continue
            n += _strip_empty_inline_node(pval, f"{display}.{pname}", removed_props)
            target = pval
            if isinstance(pval.get("items"), dict) and "$ref" not in pval["items"]:
                target = pval["items"]
            if ref_name(pval) is None and _is_empty_object(target):
                del props[pname]
                drop_from_required(node, pname)
                removed_props.append({"path": f"{display}.{pname}",
                                      "reason": "inline object left with no attributes"})
                n += 1
    for comb in ("allOf", "oneOf", "anyOf"):
        seq = node.get(comb)
        if isinstance(seq, list):
            for sub in seq:
                n += _strip_empty_inline_node(sub, display, removed_props)
    if isinstance(node.get("items"), dict):
        n += _strip_empty_inline_node(node["items"], display + "[]", removed_props)
    return n


def _strip_refs(doc, dead_names, removed_props):
    """Remove every property, array or allOf member that points at a dead schema."""

    def visit(node, display):
        if isinstance(node, dict):
            props = node.get("properties")
            if isinstance(props, dict):
                for pname in list(props.keys()):
                    pval = props[pname]
                    if _points_at(pval, dead_names):
                        del props[pname]
                        drop_from_required(node, pname)
                        removed_props.append({"path": f"{display}.{pname}",
                                              "reason": "referenced an object that was removed"})
                        continue
                    visit(pval, f"{display}.{pname}")
            for comb in ("allOf", "oneOf", "anyOf"):
                seq = node.get(comb)
                if isinstance(seq, list):
                    keep = [s for s in seq if not _points_at(s, dead_names)]
                    if len(keep) != len(seq):
                        removed_props.append({"path": f"{display}.{comb}",
                                              "reason": "member referenced an object that was removed"})
                        seq[:] = keep
                    for s in seq:
                        visit(s, display)
            if isinstance(node.get("items"), dict):
                visit(node["items"], display + "[]")
        elif isinstance(node, list):
            for i, v in enumerate(node):
                visit(v, f"{display}[{i}]")

    schemas = (doc.get("components") or {}).get("schemas") or {}
    for name, sc in schemas.items():
        visit(sc, name)


def _points_at(node, dead_names):
    if not isinstance(node, dict):
        return False
    if ref_name(node) in dead_names:
        return True
    items = node.get("items")
    if isinstance(items, dict) and ref_name(items) in dead_names:
        return True
    return False


# --------------------------------------------------------------------------- #
# D5 — de-duplication computed from the ORIGINAL document
# --------------------------------------------------------------------------- #
def _shape(node, depth=0):
    """Comparable shape of a schema node, ignoring prose and examples."""
    if not isinstance(node, dict) or depth > 8:
        return node if not isinstance(node, dict) else "?"
    out = {}
    for k in ("type", "format", "maxLength", "minLength", "pattern", "enum", "$ref"):
        if k in node:
            out[k] = plain(node[k])
    props = node.get("properties")
    if isinstance(props, dict):
        out["properties"] = {k: _shape(v, depth + 1) for k, v in sorted(props.items())}
    if isinstance(node.get("items"), dict):
        out["items"] = _shape(node["items"], depth + 1)
    req = node.get("required")
    if isinstance(req, list):
        out["required"] = sorted(req)
    for comb in ("allOf", "oneOf", "anyOf"):
        if isinstance(node.get(comb), list):
            out[comb] = [_shape(s, depth + 1) for s in node[comb]]
    return out


def detect_dedupe_candidates(schemas, strictness="deep"):
    """Inline objects equivalent to a named schema, computed non-mutatively.

    `strictness='deep'` requires full deep equality; `'shape'` ignores
    descriptions and examples but still requires the same attributes with the
    same types. Key-set comparison (the v4 behaviour) is deliberately gone.
    """
    def canon(node):
        return plain(node) if strictness == "deep" else _shape(node)

    named = {name: canon(sc) for name, sc in schemas.items()
             if isinstance(sc, dict) and isinstance(sc.get("properties"), dict)
             and sc["properties"]}
    out = []
    for owner, sc in schemas.items():
        props = sc.get("properties") if isinstance(sc, dict) else None
        if not isinstance(props, dict):
            continue
        for prop, node in props.items():
            if not (isinstance(node, dict) and isinstance(node.get("properties"), dict)
                    and node["properties"] and "$ref" not in node):
                continue
            c = canon(node)
            for target, tc in named.items():
                if target != owner and c == tc:
                    out.append((owner, prop, target))
                    break
    return out


def dedupe_inline(schemas, owner, prop, target, log, strictness="deep"):
    """Apply one de-duplication, re-verifying equivalence at apply time."""
    if owner not in schemas or target not in schemas:
        log.append(f"SKIPPED {owner}.{prop} -> {target} (schema no longer present)")
        return False
    op = schemas[owner].get("properties")
    if not isinstance(op, dict):
        return False
    node = op.get(prop)
    if not isinstance(node, dict) or "properties" not in node:
        log.append(f"SKIPPED {owner}.{prop} -> {target} (no longer an inline object)")
        return False
    canon = (lambda n: plain(n)) if strictness == "deep" else _shape
    if canon(node) != canon(schemas[target]):
        log.append(f"SKIPPED {owner}.{prop} -> {target} "
                   f"(shapes diverged after pruning, not merged)")
        return False
    op[prop] = CommentedMap([("$ref", f"#/components/schemas/{target}")])
    log.append(f"{owner}.{prop} -> $ref {target} (equivalent in the source document)")
    return True


def auto_dedupe_all(schemas, log, strictness="deep", candidates=None):
    for owner, prop, target in (candidates if candidates is not None
                                else detect_dedupe_candidates(schemas, strictness)):
        dedupe_inline(schemas, owner, prop, target, log, strictness)


# --------------------------------------------------------------------------- #
# Polymorphism refactor
# --------------------------------------------------------------------------- #
def refactor_polymorphism(schemas, members, base_name, disc_prop, add_disc, log):
    """Factor common properties of `members` into `base_name`; rewrite each
    member as allOf:[base, {extras}]. `members` is [(schema_name, disc_value)]."""
    present = [(n, v) for n, v in members
               if n in schemas and isinstance(schemas[n].get("properties"), dict)]
    if len(present) < 2:
        return False
    propmaps = {n: schemas[n]["properties"] for n, _ in present}
    common = set.intersection(*[set(p.keys()) for p in propmaps.values()])
    first = present[0][0]
    common = {k for k in common
              if all(plain(propmaps[n][k]) == plain(propmaps[first][k]) for n, _ in present)}
    if not common:
        return False

    base = CommentedMap()
    base["type"] = "object"
    base_props = CommentedMap()
    for k in propmaps[first]:
        if k in common:
            base_props[k] = propmaps[first][k]
    reqs = [set(schemas[n].get("required", []) or []) for n, _ in present]
    common_req = [k for k in base_props if all(k in r for r in reqs)]
    if add_disc:
        base_props[disc_prop] = CommentedMap([("type", "string")])
        common_req = [disc_prop] + common_req
    base["properties"] = base_props
    if common_req:
        base["required"] = common_req
    if add_disc:
        base["discriminator"] = CommentedMap([
            ("propertyName", disc_prop),
            ("mapping", CommentedMap([(v, f"#/components/schemas/{n}") for n, v in present])),
        ])
    schemas[base_name] = base
    _set_comment(base, f">>> CHANGE (added): base schema generated by SOR Polymorphizer - "
                       f"shared core of {', '.join(n for n, _ in present)}.", indent=4)

    for n, _v in present:
        extras = CommentedMap()
        for k in list(schemas[n]["properties"].keys()):
            if k not in common:
                extras[k] = schemas[n]["properties"][k]
        allof = CommentedSeq()
        allof.append(CommentedMap([("$ref", f"#/components/schemas/{base_name}")]))
        if extras:
            ext = CommentedMap()
            ext["type"] = "object"
            ext["properties"] = extras
            ext_req = [k for k in (schemas[n].get("required", []) or []) if k in extras]
            if ext_req:
                ext["required"] = ext_req
            allof.append(ext)
        new = CommentedMap()
        new["allOf"] = allof
        schemas[n] = new
        _set_comment(new, f">>> CHANGE (restructured): now allOf[{base_name}]"
                          + (f" + {list(extras.keys())}" if extras else "")
                          + " - attributes with no SOR connection omitted.", indent=4)
        log.append(f"{n} -> allOf[{base_name}" + (f", +{list(extras.keys())}]" if extras else "]"))
    log.insert(0, f"Base {base_name} holds common {sorted(common)}"
                  + (f" + discriminator '{disc_prop}'." if add_disc else "."))
    return True


def family_base_and_values(names):
    pre = os.path.commonprefix(names)
    suf = os.path.commonprefix([n[::-1] for n in names])[::-1]
    values = {}
    for n in names:
        mid = n[len(pre):len(n) - len(suf)] if (len(pre) + len(suf)) < len(n) else n
        values[n] = norm(mid) or norm(n)
    base = (pre + suf + "Base") if (pre or suf) else "PolymorphicBase"
    return base, values


def auto_detect_groups(doc, disc_prop):
    requests, responses = operation_body_schemas(doc)
    groups = []
    for family in (responses, requests):
        names = sorted(set(family))
        if len(names) >= 2:
            base, values = family_base_and_values(names)
            groups.append({"base": base, "disc_prop": disc_prop,
                           "members": [(n, values[n]) for n in names]})
    return groups


def _group_common(schemas, members):
    present = [(n, v) for n, v in members
               if n in schemas and isinstance(schemas[n].get("properties"), dict)]
    if len(present) < 2:
        return None
    propmaps = {n: schemas[n]["properties"] for n, _ in present}
    common = set.intersection(*[set(p.keys()) for p in propmaps.values()])
    first = present[0][0]
    common = {k for k in common
              if all(plain(propmaps[n][k]) == plain(propmaps[first][k]) for n, _ in present)}
    return common or None


def group_to_spec(g):
    members = ", ".join(f"{n}:{v}" for n, v in g["members"])
    return f'{g["base"]} = {members}'


def dedupe_to_spec(t):
    owner, prop, target = t
    return f"{owner}.{prop} = {target}"


def propose_discriminator(groups):
    if not groups:
        return ""
    cleaned = [re.sub(r"(Request|Response|Base)", "", g["base"]) for g in groups]
    pre = re.sub(r"[^A-Za-z0-9]+$", "", os.path.commonprefix([c for c in cleaned if c]))
    if not pre:
        alt = [re.sub(r"Base$", "", g["base"]) for g in groups]
        pre = re.sub(r"[^A-Za-z0-9]+$", "", os.path.commonprefix(alt))
    if not pre:
        return "context"
    return pre[0].lower() + pre[1:] + "Context"


# --------------------------------------------------------------------------- #
# D8 — per-operation allOf profiling (the tool's original purpose)
# --------------------------------------------------------------------------- #
#
# A schema shared by several operations must be used in full at every one of
# them. Where the mapping gives those operations DIFFERENT subsets, plain
# pruning can only keep the union, so every operation carries attributes it
# does not map. That is the problem this tool exists to solve, and elimination
# cannot solve it.
#
# The fix is specialisation by inheritance:
#
#   * compute the mapped subset of the schema at EACH position it is used;
#   * the attributes mapped at every position become the BASE (the schema keeps
#     its own name, so unaffected references do not move);
#   * each distinct remaining subset becomes a derived profile,
#     allOf: [ $ref base, { the attributes that position adds } ];
#   * a position whose profile equals the base references the base directly;
#   * a position that maps NOTHING is removed altogether.
#
# No discriminator is added. The profile is selected by the position in the
# document, which is fixed at design time by the method and path, so there is
# nothing to resolve at runtime. Adding a selector would introduce a field the
# SOR does not supply, and would have to be marked required.


def _rel_ref_sites(node, prefix, out, depth=0):
    """Every $ref site inside one schema, as (rel_segments, container, key, kind, target)."""
    if not isinstance(node, dict) or depth > 12:
        return
    props = node.get("properties")
    if isinstance(props, dict):
        for pname, pval in props.items():
            rel = prefix + [pname]
            if not isinstance(pval, dict):
                continue
            if ref_name(pval):
                out.append((tuple(rel), props, pname, "prop", ref_name(pval)))
                continue
            items = pval.get("items")
            if isinstance(items, dict) and ref_name(items):
                out.append((tuple(rel), pval, "items", "items", ref_name(items)))
                continue
            _rel_ref_sites(pval, rel, out, depth + 1)
    items = node.get("items")
    if isinstance(items, dict) and not ref_name(items):
        _rel_ref_sites(items, prefix, out, depth + 1)
    for comb in ("allOf", "oneOf", "anyOf"):
        seq = node.get(comb)
        if isinstance(seq, list):
            for sub in seq:
                if isinstance(sub, dict) and not ref_name(sub):
                    _rel_ref_sites(sub, prefix, out, depth + 1)


def collect_positions(doc, max_positions=20000):
    """Every place a named schema is referenced, with the concrete container.

    Returns {schema_name: [(abs_path_tuple, container, key, kind), ...]} where
    `kind` is 'prop' (container[key] is the property value) or 'items'
    (container[key] is an array's items node). Array levels add no path
    segment, matching the mapping workbook's convention.
    """
    from collections import deque

    schemas = (doc.get("components") or {}).get("schemas") or {}

    rel_cache = {}
    for name, sc in schemas.items():
        out = []
        _rel_ref_sites(sc, [], out)
        rel_cache[name] = out

    requests, responses = operation_body_schemas(doc)
    roots = list(dict.fromkeys(list(requests) + list(responses))) or list(schemas)
    positions = {name: [] for name in schemas}
    seen, count = set(), 0
    dq = deque((r, (r,)) for r in roots if r in schemas)
    while dq and count < max_positions:
        name, path = dq.popleft()
        if (name, path) in seen or len(path) > 16:
            continue
        seen.add((name, path))
        count += 1
        for rel, container, key, kind, target in rel_cache.get(name, ()):
            abs_path = path + rel
            positions.setdefault(target, []).append((abs_path, container, key, kind))
            dq.append((target, abs_path))
    return positions


def is_scalar_array(node):
    """An array of scalars is a data attribute, not structure."""
    if not isinstance(node, dict) or node.get("type") != "array":
        return False
    items = node.get("items")
    if not isinstance(items, dict) or ref_name(items):
        return False
    return not items.get("properties") and items.get("type") not in ("object", None)


def _leaf_names(schema):
    """Attributes that carry data: scalars and arrays of scalars."""
    props = (schema or {}).get("properties")
    if not isinstance(props, dict):
        return []
    return [p for p, v in props.items() if not is_container(v) or is_scalar_array(v)]


def _safe_name(base, suffix, taken):
    name = f"{base}{suffix}"
    n = 2
    while name in taken:
        name = f"{base}{suffix}{n}"
        n += 1
    return name


def profile_shared_schemas(doc, mapping, *, on_unmatched="keep", protected=("ErrorResponse",),
                           sor_fields=None, require_sor_field=False, log=lambda m: None):
    """Split shared schemas into a base plus per-operation allOf profiles.

    Runs AFTER prune_document, which leaves each shared schema holding the
    union of what its positions map. This distributes that union back to the
    positions that actually map each part.

    The algorithm is recursive over the POSITION graph, not flat over schemas.
    A schema's positions are equivalent only if their whole sub-trees are
    equivalent, because a nested container physically lives in one shared
    dictionary: two routes cannot point it at different children unless their
    common parent is specialised first. Signatures are therefore computed
    depth-first and schemas are split from the leaves upward.

    The base never re-declares a container property whose target varies by
    route. `allOf` conjoins rather than overrides, so a base and an extension
    that both declare the same property would be intersected, not replaced.
    Anything that varies is declared only on the profiles that need it.
    """
    schemas = (doc.get("components") or {}).get("schemas") or {}
    protected = set(protected or ())
    positions = collect_positions(doc)

    # ---- child sites of every schema, resolved once ---------------------- #
    child_sites = {}
    for name, sc in schemas.items():
        out = []
        _rel_ref_sites(sc, [], out)
        child_sites[name] = out                     # (rel, container, key, kind, target)

    def mapped_at(path, attr):
        entry, _s, _t = mapping.resolve([list(path) + [attr]], attr)
        if entry is None:
            return on_unmatched != "drop"
        if entry["status"] == "unmapped":
            return False
        if require_sor_field and sor_fields is not None and \
                not _field_in_sor(entry.get("sor", ""), sor_fields):
            return False
        return True

    # ---- signature of a position, computed depth-first ------------------- #
    sig_memo = {}

    def signature(name, path, depth=0):
        key = (name, path)
        if key in sig_memo:
            return sig_memo[key]
        if depth > 14 or name not in schemas:
            sig_memo[key] = (frozenset(), ())
            return sig_memo[key]
        sig_memo[key] = (frozenset(), ())            # cycle guard
        leaves = frozenset(a for a in _leaf_names(schemas[name]) if mapped_at(path, a))
        kids = []
        for rel, _c, _k, _kind, target in child_sites.get(name, ()):
            kids.append((rel, signature(target, path + rel, depth + 1)))
        sig = (leaves, tuple(sorted(kids, key=lambda x: x[0])))
        sig_memo[key] = sig
        return sig

    for name in list(positions):
        for path, _c, _k, _kind in positions.get(name) or []:
            signature(name, path)

    def is_dead(sig):
        """A position contributes nothing: no mapped leaf and every child dead."""
        leaves, kids = sig
        return not leaves and all(is_dead(cs) for _rel, cs in kids)

    # ---- decide a target per (schema, signature), leaves upward ---------- #
    depth_of = {}
    for name, sites in positions.items():
        depth_of[name] = max((len(p) for p, _c, _k, _kind in sites), default=0)
    order = sorted(positions, key=lambda n: -depth_of.get(n, 0))

    target_of = {}          # (schema, signature) -> new schema name, or None if dead
    actions, created, removed_positions = [], [], []

    for name in order:
        if name in protected or name not in schemas:
            continue
        sites = positions.get(name) or []
        if len(sites) < 2:
            continue
        sigs = {}
        for path, _c, _k, _kind in sites:
            sigs.setdefault(signature(name, path), []).append(path)
        if len(sigs) < 2:
            continue

        live = {sg: ps for sg, ps in sigs.items() if not is_dead(sg)}
        for sg in sigs:
            if sg not in live:
                target_of[(name, sg)] = None
        if not live:
            continue

        schema = schemas[name]
        all_leaves = _leaf_names(schema)
        orig_required = list(schema.get("required") or [])

        # what every live route agrees on
        common_leaves = set.intersection(*[set(sg[0]) for sg in live]) if live else set()
        kid_target = {}                              # (rel) -> {sig-target per route}
        for sg in live:
            for rel, cs in sg[1]:
                kid_target.setdefault(rel, set()).add(_kid_target(rel, cs, child_sites[name], target_of))
        common_kids = {rel for rel, t in kid_target.items() if len(t) == 1 and next(iter(t)) is not None}

        if len(live) == 1:
            # every live route agrees; only dead routes differ, so no split
            target_of[(name, next(iter(live)))] = name
            _reduce_base(schema, all_leaves, common_leaves, child_sites[name],
                         common_kids, kid_target, target_of, live, name, actions)
            continue

        base_possible = bool(common_leaves or common_kids)
        for sg in sorted(live, key=lambda g: (-len(g[0]), sorted(g[0]))):
            route_leaves = set(sg[0])
            route_kids = {rel: _kid_target(rel, cs, child_sites[name], target_of)
                          for rel, cs in sg[1]}
            same_as_base = (base_possible
                            and route_leaves == common_leaves
                            and all(route_kids.get(rel) == next(iter(kid_target[rel]))
                                    for rel in common_kids)
                            and not (set(route_kids) - common_kids
                                     and any(route_kids[r] is not None
                                             for r in set(route_kids) - common_kids)))
            if same_as_base:
                target_of[(name, sg)] = name
                continue
            roots = sorted({p[0] for p in live[sg]})
            suffix = f"For{roots[0]}" if len(roots) == 1 else None
            new = _safe_name(name, suffix, set(schemas)) if suffix else \
                _safe_name(name, f"Profile{1 + sum(1 for k in target_of if k[0] == name and isinstance(target_of[k], str) and target_of[k].startswith(name + 'Profile'))}", set(schemas))
            target_of[(name, sg)] = new

            ext = CommentedMap()
            ext["type"] = "object"
            ext["description"] = (f"The part of {name} that the mapping supports for "
                                  f"{', '.join(roots)}.")
            ep = CommentedMap()
            for a in all_leaves:
                if a in route_leaves and (not base_possible or a not in common_leaves):
                    ep[a] = _copy_prop(schema, a)
            for rel, _c, _k, _kind, _t in child_sites[name]:
                if len(rel) != 1:
                    continue
                tgt = route_kids.get(rel)
                if tgt is None or rel in common_kids:
                    continue
                ep[rel[0]] = CommentedMap([("$ref", f"#/components/schemas/{tgt}")])
            ext["properties"] = ep
            req = [r for r in orig_required if r in ep]
            if req:
                ext["required"] = CommentedSeq(req)

            derived = CommentedMap()
            if base_possible:
                seq = CommentedSeq()
                seq.append(CommentedMap([("$ref", f"#/components/schemas/{name}")]))
                seq.append(ext)
                derived["allOf"] = seq
                actions.append(f"{new} = allOf[{name}] + {sorted(ep.keys())}  "
                               f"(used by {', '.join(roots)})")
            else:
                derived = ext
                actions.append(f"{new} = standalone {sorted(ep.keys())}  "
                               f"(used by {', '.join(roots)})")
            schemas[new] = derived
            _set_comment(derived, f">>> CHANGE (added): per-operation profile of {name} "
                                  f"for {', '.join(roots)}.", indent=4)
            created.append(new)

        _reduce_base(schema, all_leaves, common_leaves if base_possible else set(),
                     child_sites[name], common_kids if base_possible else set(),
                     kid_target, target_of, live, name, actions)

    # ---- apply: rewrite each position to its route's target -------------- #
    for name, sites in positions.items():
        for path, container, key, kind in sites:
            sg = signature(name, path)
            if (name, sg) not in target_of:
                continue
            target = target_of[(name, sg)]
            if not isinstance(container, dict):
                continue
            if target is None:
                if kind == "items":
                    container["items"] = CommentedMap([("x-sor-removed", True)])
                elif key in container:
                    del container[key]
                else:
                    continue
                removed_positions.append(".".join(path))
                continue
            if target == name or key not in container:
                continue
            node = container[key]
            if isinstance(node, dict):
                node.clear()
                node["$ref"] = f"#/components/schemas/{target}"

    _scrub_dead_arrays(schemas, removed_positions)

    if actions or removed_positions:
        log(f"  Profiling: {len(created)} derived profile(s) created, "
            f"{len(removed_positions)} position(s) removed that mapped nothing.")
    return {"actions": actions, "created": created,
            "removed_positions": sorted(set(removed_positions))}


def _copy_prop(schema, attr):
    import copy as _c
    return _c.deepcopy(schema["properties"][attr])


def _kid_target(rel, child_sig, sites, target_of):
    """The schema name a route's child position should point at, or None if dead."""
    for r, _c, _k, _kind, target in sites:
        if r == rel:
            if (target, child_sig) in target_of:
                return target_of[(target, child_sig)]
            return target
    return None


def _reduce_base(schema, all_leaves, common_leaves, sites, common_kids, kid_target,
                 target_of, live, name, actions):
    """Strip from the base everything that varies by route."""
    props = schema.get("properties")
    if not isinstance(props, dict):
        return
    moved = []
    for a in all_leaves:
        if a not in common_leaves and a in props:
            del props[a]
            drop_from_required(schema, a)
            moved.append(a)
    for rel, _c, _k, _kind, _t in sites:
        if len(rel) != 1:
            continue
        pname = rel[0]
        if rel in common_kids or pname not in props:
            continue
        if rel in kid_target:
            del props[pname]
            drop_from_required(schema, pname)
            moved.append(pname)
    for rel in common_kids:
        tgt = next(iter(kid_target[rel]))
        pname = rel[0] if len(rel) == 1 else None
        if pname and tgt and pname in props and isinstance(props[pname], dict):
            node = props[pname]
            if ref_name(node) and ref_name(node) != tgt:
                node.clear()
                node["$ref"] = f"#/components/schemas/{tgt}"
    if moved:
        actions.insert(0, f"{name} reduced to the shared core "
                          f"{sorted(common_leaves | {r[0] for r in common_kids if len(r) == 1})} "
                          f"(moved to profiles: {sorted(moved)})")
        _set_comment(schema, f">>> CHANGE (reduced): now the shared core of {name}; "
                             f"per-operation attributes moved to profiles.", indent=4)


def _scrub_dead_arrays(schemas, removed_positions):
    def visit(node, display):
        if isinstance(node, dict):
            props = node.get("properties")
            if isinstance(props, dict):
                for k in list(props.keys()):
                    v = props[k]
                    if isinstance(v, dict) and isinstance(v.get("items"), dict) \
                            and v["items"].get("x-sor-removed"):
                        del props[k]
                        drop_from_required(node, k)
                        removed_positions.append(f"{display}.{k}[]")
                        continue
                    visit(v, f"{display}.{k}")
            if isinstance(node.get("items"), dict):
                visit(node["items"], display + "[]")
            for comb in ("allOf", "oneOf", "anyOf"):
                if isinstance(node.get(comb), list):
                    for sub in node[comb]:
                        visit(sub, display)
    for nm, sc in list(schemas.items()):
        visit(sc, nm)


# --------------------------------------------------------------------------- #
# D6 — post-run assertions
# --------------------------------------------------------------------------- #
def assert_output(doc, protected=("ErrorResponse",)):
    """Structural checks the transformed document must pass."""
    problems = []
    schemas = (doc.get("components") or {}).get("schemas") or {}
    protected = set(protected or ())

    for name, sc in schemas.items():
        if name in protected:
            continue
        if _is_empty_object(sc):
            problems.append(f"schema '{name}' has no attributes left")

    def walk(node, display):
        if isinstance(node, dict):
            ref = node.get("$ref")
            if isinstance(ref, str) and _resolve_local(doc, ref) is None:
                problems.append(f"dangling $ref at {display}: {ref}")
            props = node.get("properties")
            if isinstance(props, dict):
                req = node.get("required")
                if isinstance(req, list):
                    for r in req:
                        if r not in props:
                            problems.append(f"{display}: required '{r}' is not a property")
                for k, v in props.items():
                    if isinstance(v, dict) and _is_empty_object(v):
                        problems.append(f"{display}.{k} is an object with no attributes")
                    walk(v, f"{display}.{k}")
            for k in ("items", "additionalProperties"):
                if isinstance(node.get(k), dict):
                    walk(node[k], f"{display}.{k}")
            for comb in ("allOf", "oneOf", "anyOf"):
                if isinstance(node.get(comb), list):
                    for i, s in enumerate(node[comb]):
                        walk(s, f"{display}.{comb}[{i}]")
        elif isinstance(node, list):
            for i, v in enumerate(node):
                walk(v, f"{display}[{i}]")

    for name, sc in schemas.items():
        if name in protected:
            continue
        walk(sc, name)
    for section in ("paths", "components"):
        walk(doc.get(section) or {}, section)
    return sorted(set(problems))


# --------------------------------------------------------------------------- #
# Reporting
# --------------------------------------------------------------------------- #
def annotate_document(doc, result):
    """Record every change inside the output document itself."""
    dropped = result["dropped"]
    casc = result["cascade"]
    diag = result["diagnostics"]
    banner = (
        "############################################################\n"
        f"#  Transformed by SOR Polymorphizer v{__version__}\n"
        f"#   - mapping matched {diag['attributes_matched']}/{diag['attributes_total']} "
        f"attributes ({diag['match_rate'] * 100:.1f}%)\n"
        f"#   - {len(dropped)} attribute(s) eliminated (positive no-SOR signal)\n"
        f"#   - {diag['attributes_unmatched']} attribute(s) kept and flagged "
        f"(not found in the mapping)\n"
        f"#   - {len(casc['removed_schemas'])} emptied object(s) and "
        f"{len(casc['removed_properties'])} reference(s) cascade-removed\n"
        f"#   - {len((result.get('profiling') or {}).get('created', []))} per-operation "
        f"profile(s) created by allOf specialisation\n"
        f"#   - {len((result.get('profiling') or {}).get('removed_positions', []))} position(s) "
        f"removed that mapped nothing\n"
        f"#   - {result['n_bases']} explicit polymorphism group(s) applied\n"
        f"#   - {len(result['dedupe_log'])} inline object(s) de-duplicated\n"
        "#  Full change list: see info.x-sor-polymorphizer below,\n"
        "#  or the colour-coded HTML change report.\n"
        "#  Change sites are marked in-file with '>>> CHANGE' comments.\n"
        "############################################################"
    )
    _set_comment(doc, banner, indent=0)

    info = doc.get("info")
    if isinstance(info, dict):
        rec = CommentedMap()
        rec["toolVersion"] = __version__
        rec["summary"] = (
            f"{len(dropped)} attributes eliminated; "
            f"{diag['attributes_unmatched']} kept and flagged as absent from the mapping; "
            f"{len(casc['removed_schemas'])} emptied objects cascade-removed; "
            f"{result['n_bases']} groups polymorphized; "
            f"{len(result['dedupe_log'])} inline objects de-duplicated.")
        cov = CommentedMap()
        cov["attributesInSpec"] = diag["attributes_total"]
        cov["matchedToMapping"] = diag["attributes_matched"]
        cov["notFoundInMapping"] = diag["attributes_unmatched"]
        cov["matchRate"] = f"{diag['match_rate'] * 100:.1f}%"
        cov["resolutionStrategies"] = CommentedMap(list(diag["strategies"].items()))
        rec["mappingCoverage"] = cov
        rec["eliminatedAttributes"] = CommentedSeq(
            [f"{d['path']}  ({d['reason']}; matched by {d['strategy']})" for d in dropped])
        if diag["unmatched"]:
            rec["keptButNotInMapping"] = CommentedSeq(
                [u["path"] for u in diag["unmatched"]])
        if casc["removed_schemas"]:
            rec["cascadeRemovedObjects"] = CommentedSeq(list(casc["removed_schemas"]))
        if casc["removed_properties"]:
            rec["cascadeRemovedReferences"] = CommentedSeq(
                [f"{p['path']}  ({p['reason']})" for p in casc["removed_properties"]])
        if casc["empty_body_schemas"]:
            rec["emptyOperationBodies"] = CommentedSeq(list(casc["empty_body_schemas"]))
        prof = result.get("profiling") or {}
        if prof.get("actions"):
            rec["perOperationProfiles"] = CommentedSeq(list(prof["actions"]))
        if prof.get("removed_positions"):
            rec["positionsRemovedMappingNothing"] = CommentedSeq(list(prof["removed_positions"]))
        rec["polymorphism"] = CommentedSeq(list(result["poly_log"]))
        rec["deduplication"] = CommentedSeq(list(result["dedupe_log"]))
        info["x-sor-polymorphizer"] = rec

    schemas = (doc.get("components") or {}).get("schemas") or {}
    by_schema = {}
    for d in dropped:
        top = d["path"].split(".")[0].split("[")[0]
        by_schema.setdefault(top, []).append(d["path"].split(".", 1)[-1])
    for name, attrs in by_schema.items():
        node = schemas.get(name)
        if isinstance(node, dict) and "allOf" not in node:
            _set_comment(node, ">>> CHANGE (removed, no SOR connection): " + ", ".join(attrs),
                         indent=4)


def _write(path, text):
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(text)
    return path


# --------------------------------------------------------------------------- #
# Text I/O — always explicit about encoding
# --------------------------------------------------------------------------- #
# Every read and write in this module names its encoding. Relying on the
# platform default silently produced cp1252 output on Windows, which the HTML
# report then failed to read back as UTF-8. Inputs are decoded with a short
# fallback chain so a specification saved by a Windows editor still loads.
READ_ENCODINGS = ("utf-8", "cp1252", "latin-1")
_BOM_UTF8 = b"\xef\xbb\xbf"


def read_text(path):
    """Read a text file, returning (text, encoding_used).

    A UTF-8 byte-order mark is recognised and stripped. Otherwise UTF-8 is
    tried first, then the Windows single-byte encodings, so a specification
    saved as ANSI by a Windows editor still loads. Whatever comes in, this
    module always writes UTF-8 back out.
    """
    with open(path, "rb") as fh:
        raw = fh.read()
    if raw.startswith(_BOM_UTF8):
        return raw[len(_BOM_UTF8):].decode("utf-8"), "utf-8 (BOM)"
    for enc in READ_ENCODINGS:
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace"), "utf-8 (with replacements)"


def build_report(swagger, mapping_display, out, result):
    diag = result["diagnostics"]
    casc = result["cascade"]
    L = [f"# SOR-driven refactor report (SOR Polymorphizer v{__version__})\n",
         f"* Source swagger: `{swagger}`",
         f"* SOR mapping:   `{mapping_display}`",
         f"* Output:        `{out}`\n",
         "## Mapping coverage\n",
         f"* Attributes in the specification: **{diag['attributes_total']}**",
         f"* Resolved against the mapping:    **{diag['attributes_matched']}** "
         f"({diag['match_rate'] * 100:.1f}%)",
         f"* Not found in the mapping:        **{diag['attributes_unmatched']}** "
         f"(kept and flagged)",
         f"* Resolution strategies: {diag['strategies']}",
         f"* Mapping rows indexed: {result['mapping_stats']['attribute_rows']} "
         f"across {result['mapping_stats']['sheets']} sheet(s) "
         f"({result['mapping_stats']['mapped']} mapped, "
         f"{result['mapping_stats']['unmapped']} marked not-used)\n",
         "### Detected spreadsheet layout\n"]
    for lay in result["layouts"]:
        cols = ", ".join(f"{k}=col{v}" for k, v in sorted(lay["cols"].items()))
        lv = f", levels={lay['levels']}" if lay["levels"] else ""
        L.append(f"* `{lay['sheet']}` — mode **{lay['mode']}**, header row {lay['header_row']}, "
                 f"{cols}{lv}")
    L.append(f"\n## Attributes eliminated ({len(result['dropped'])})\n")
    L += ([f"* `{d['path']}` — {d['reason']} (matched by {d['strategy']})"
           for d in result["dropped"]] or ["* (none)"])
    if diag["unmatched"]:
        L.append(f"\n## Kept but not found in the mapping ({len(diag['unmatched'])})\n")
        L.append("These were retained. Each line shows the lookup keys that were tried, "
                 "so a join failure is visible as a join failure.\n")
        for u in diag["unmatched"][:400]:
            L.append(f"* `{u['path']}` — tried: {', '.join('`%s`' % k for k in u['keys_tried'])}")
        if len(diag["unmatched"]) > 400:
            L.append(f"* … {len(diag['unmatched']) - 400} more")
    L.append(f"\n## Cascade removals\n")
    L.append(f"* Emptied objects removed: {len(casc['removed_schemas'])}")
    for n in casc["removed_schemas"]:
        L.append(f"    * `{n}`")
    L.append(f"* References removed: {len(casc['removed_properties'])}")
    for p in casc["removed_properties"][:200]:
        L.append(f"    * `{p['path']}` — {p['reason']}")
    if casc["empty_body_schemas"]:
        L.append("* **Operation bodies that would be empty (left in place, review the mapping):**")
        for n in casc["empty_body_schemas"]:
            L.append(f"    * `{n}`")
    L.append("\n## Polymorphism applied\n")
    L += (result["poly_log"] or ["* (no shared-schema differences required a base/profile split)"])
    L.append("\n## De-duplication\n")
    L += ([f"* {d}" for d in result["dedupe_log"]] or ["* (none)"])
    val = result.get("validation")
    if val is not None:
        L.append(f"\n## SOR YAML validation ({len(val['confirmed'])} confirmed, "
                 f"{len(val['missing'])} not found)\n")
        L.append(f"* {val['field_count']} SOR field name(s) discovered.")
        for leaf, field in val["missing"][:200]:
            L.append(f"    * `{leaf}` → `{field}` (not in SOR YAML)")
    L.append("\n## Output assertions\n")
    L += ([f"* FAILED: {p}" for p in result["assertions"]] or
          ["* All post-run assertions passed (no empty objects, no dangling refs, "
           "no required-without-property)."])
    return "\n".join(L) + "\n"


def write_html_report(original_path, output_path, result, html_path):
    orig = read_text(original_path)[0].splitlines()
    new = read_text(output_path)[0].splitlines()

    rows = []
    for line in difflib.unified_diff(orig, new, fromfile="original", tofile="polymorphized",
                                     lineterm="", n=3):
        esc = _html.escape(line)
        cls = ("meta" if line.startswith(("+++", "---")) else
               "hunk" if line.startswith("@@") else
               "add" if line.startswith("+") else
               "del" if line.startswith("-") else "ctx")
        rows.append(f'<div class="line {cls}">{esc or "&nbsp;"}</div>')

    diag = result["diagnostics"]
    casc = result["cascade"]

    def table(headers, body_rows, empty="none"):
        if not body_rows:
            return f'<p><em>{empty}</em></p>'
        th = "".join(f"<th>{_html.escape(h)}</th>" for h in headers)
        return f"<table><tr>{th}</tr>{''.join(body_rows)}</table>"

    elim = [f'<tr><td class="mono">{_html.escape(d["path"])}</td>'
            f'<td>{_html.escape(d["reason"])}</td>'
            f'<td>{_html.escape(d["strategy"])}</td>'
            f'<td class="mono">{_html.escape(str(d.get("sor", "")))}</td></tr>'
            for d in result["dropped"]]
    unmatched = [f'<tr><td class="mono">{_html.escape(u["path"])}</td>'
                 f'<td class="mono small">{_html.escape(", ".join(u["keys_tried"]))}</td></tr>'
                 for u in diag["unmatched"][:500]]
    lay = [f'<tr><td>{_html.escape(l["sheet"])}</td><td>{_html.escape(l["mode"])}</td>'
           f'<td class="num">{l["header_row"]}</td>'
           f'<td class="mono small">{_html.escape(", ".join(f"{k}=col{v}" for k, v in sorted(l["cols"].items())))}'
           f'{_html.escape(" levels=" + str(l["levels"]) if l["levels"] else "")}</td></tr>'
           for l in result["layouts"]]
    casc_rows = [f'<tr><td class="mono">{_html.escape(p["path"])}</td>'
                 f'<td>{_html.escape(p["reason"])}</td></tr>'
                 for p in casc["removed_properties"][:500]]
    asserts = ("".join(f'<li class="bad">{_html.escape(p)}</li>' for p in result["assertions"])
               or '<li class="ok">All post-run assertions passed.</li>')
    rate = diag["match_rate"] * 100
    rate_cls = "good" if rate >= 75 else ("warn" if rate >= 40 else "bad")

    return _write(html_path, f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<title>SOR Polymorphizer - change report</title>
<style>
 body{{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f4f6f8;color:#1f2a33}}
 header{{background:#21295C;color:#fff;padding:20px 28px}}
 header h1{{margin:0;font-size:20px}} header p{{margin:6px 0 0;color:#CADCFC}}
 .wrap{{padding:22px 28px;max-width:1240px;margin:0 auto}}
 .cards{{display:flex;gap:14px;margin-bottom:20px;flex-wrap:wrap}}
 .card{{background:#fff;border-radius:10px;padding:16px 20px;box-shadow:0 1px 4px rgba(0,0,0,.1);min-width:150px}}
 .card .n{{font-size:30px;font-weight:700;color:#065A82}} .card .l{{font-size:13px;color:#5A6B78}}
 .card .n.good{{color:#1E7A34}} .card .n.warn{{color:#9A6A00}} .card .n.bad{{color:#B3261E}}
 h2{{font-size:15px;color:#21295C;margin:22px 0 8px}}
 table{{border-collapse:collapse;width:100%;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.08)}}
 td,th{{padding:7px 12px;border-bottom:1px solid #eef2f5;text-align:left;font-size:13px;vertical-align:top}}
 th{{background:#eef3f7}} .num{{text-align:right;font-variant-numeric:tabular-nums}}
 .mono{{font-family:Consolas,monospace;color:#B33A3A}} .small{{font-size:11.5px;color:#5A6B78}}
 ul{{background:#fff;border-radius:8px;padding:12px 12px 12px 32px;box-shadow:0 1px 4px rgba(0,0,0,.08);font-size:13px}}
 li.bad{{color:#B3261E;font-weight:600}} li.ok{{color:#1E7A34}}
 .diff{{background:#0d1b2a;border-radius:8px;padding:14px;overflow:auto;font-family:Consolas,monospace;font-size:12.5px;line-height:1.5;max-height:70vh}}
 .line{{white-space:pre;padding:0 6px;border-radius:2px}}
 .add{{background:#123d1e;color:#8ff0a4}} .del{{background:#4a1420;color:#ff9db0}}
 .hunk{{color:#6ea8fe;margin-top:6px}} .meta{{color:#8b949e}} .ctx{{color:#c9d1d9}}
 .legend span{{display:inline-block;padding:2px 10px;border-radius:4px;margin-right:8px;font-size:12px}}
</style></head><body>
<header><h1>SOR Polymorphizer — change report</h1>
<p>{_html.escape(os.path.basename(original_path))} → {_html.escape(os.path.basename(output_path))}
 · engine v{__version__}</p></header>
<div class="wrap">
 <div class="cards">
   <div class="card"><div class="n {rate_cls}">{rate:.0f}%</div>
     <div class="l">mapping match rate<br>{diag['attributes_matched']} of {diag['attributes_total']} attributes</div></div>
   <div class="card"><div class="n">{len(result['dropped'])}</div><div class="l">attributes eliminated<br>(positive no-SOR signal)</div></div>
   <div class="card"><div class="n">{diag['attributes_unmatched']}</div><div class="l">kept &amp; flagged<br>(absent from mapping)</div></div>
   <div class="card"><div class="n">{len(casc['removed_schemas'])}</div><div class="l">emptied objects<br>cascade-removed</div></div>
   <div class="card"><div class="n">{len(result['dedupe_log'])}</div><div class="l">inline objects<br>de-duplicated</div></div>
   <div class="card"><div class="n">{result['n_bases']}</div><div class="l">polymorphism<br>groups</div></div>
 </div>
 <h2>Post-run assertions</h2><ul>{asserts}</ul>
 <h2>Detected spreadsheet layout</h2>
 {table(["Sheet", "Mode", "Header row", "Columns"], lay)}
 <h2>Attributes eliminated</h2>
 {table(["Attribute", "Reason", "Matched by", "SOR field"], elim)}
 <h2>Kept but not found in the mapping</h2>
 {table(["Attribute", "Lookup keys tried"], unmatched,
        "Every attribute in the specification was found in the mapping.")}
 <h2>Cascade removals</h2>
 {table(["Reference", "Reason"], casc_rows, "Nothing needed cascade removal.")}
 <h2>Polymorphism applied</h2><ul>{''.join(f'<li>{_html.escape(str(x))}</li>' for x in result['poly_log']) or '<li><em>none</em></li>'}</ul>
 <h2>De-duplication</h2><ul>{''.join(f'<li>{_html.escape(str(x))}</li>' for x in result['dedupe_log']) or '<li><em>none</em></li>'}</ul>
 <h2>Full diff <span class="legend"><span class="add">added</span><span class="del">removed</span></span></h2>
 <div class="diff">{''.join(rows)}</div>
</div></body></html>
""")


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #
def _load_doc(path):
    """Load a YAML document, tolerating a non-UTF-8 source file."""
    yaml = YAML()
    yaml.preserve_quotes = True
    yaml.width = 4096
    text, enc = read_text(path)
    return yaml.load(text), yaml, enc


def diagnose(swagger, mapping, *, sheet_cfg=None, protected=("ErrorResponse",),
             on_unmatched="keep", log=lambda m: None):
    """Dry run that reports mapping coverage only. Writes nothing."""
    m = load_mapping(mapping, sheet_cfg)
    doc, _yaml, _enc = _load_doc(swagger)
    diag = prune_document(copy.deepcopy(doc), m, on_unmatched=on_unmatched,
                          protected=protected, log=log)
    return {"diagnostics": diag, "layouts": m.layouts, "mapping_stats": m.stats,
            "nearest": {u["path"]: m.nearest_keys(u["keys_tried"][0] if u["keys_tried"] else "")
                        for u in diag["unmatched"][:50]}}


def analyze(swagger, mapping, *, sor_swaggers=None, require_sor_field=False,
            keep_absent=True, on_unmatched=None, protected=("ErrorResponse",),
            profile_shared=True, sheet_cfg=None, discriminator_prop="resultContext",
            dedupe_strictness="deep"):
    """Inspect a swagger + mapping and propose the transform. Writes nothing."""
    on_unmatched = on_unmatched or ("keep" if keep_absent else "keep")
    m = load_mapping(mapping, sheet_cfg)
    sor_swaggers = list(sor_swaggers or [])
    sor_fields = collect_sor_fields(sor_swaggers) if sor_swaggers else None
    doc, _yaml, _enc = _load_doc(swagger)
    schemas = (doc.get("components") or {}).get("schemas") or {}

    # D5: dedup equivalence is decided on the ORIGINAL document.
    dedupe = detect_dedupe_candidates(schemas, dedupe_strictness)

    diag = prune_document(doc, m, on_unmatched=on_unmatched, protected=protected,
                          sor_fields=sor_fields, require_sor_field=require_sor_field)
    profiling = {"actions": [], "created": [], "removed_positions": []}
    if profile_shared:
        profiling = profile_shared_schemas(
            doc, m, on_unmatched=on_unmatched, protected=protected,
            sor_fields=sor_fields, require_sor_field=require_sor_field)
    cascade = cascade_prune(doc, protected=protected)
    dlog = []
    auto_dedupe_all(schemas, dlog, dedupe_strictness, dedupe)
    groups = [g for g in auto_detect_groups(doc, discriminator_prop)
              if _group_common(schemas, g["members"])]

    return {
        "eliminations": [(d["path"], d["reason"]) for d in diag["dropped"]],
        "unmatched": [(u["path"], ", ".join(u["keys_tried"])) for u in diag["unmatched"]],
        "diagnostics": diag,
        "profiling": profiling,
        "cascade": cascade,
        "layouts": m.layouts,
        "mapping_stats": m.stats,
        "groups": groups,
        "dedupe": dedupe,
        "group_specs": [group_to_spec(g) for g in groups],
        "dedupe_specs": [dedupe_to_spec(t) for t in dedupe],
        "discriminator": propose_discriminator(groups),
        "validation": validate_mapping(m, sor_fields) if sor_fields is not None else None,
        "assertions": assert_output(doc, protected),
    }


def run(swagger, mapping, out, report=None, *, html_report=None, sor_swaggers=None,
        require_sor_field=False, keep_absent=True, on_unmatched=None,
        add_discriminator=False, discriminator_prop="resultContext", poly_groups=None,
        dedupe_pairs=None, auto_poly=False, auto_dedupe=False, protected=("ErrorResponse",),
        profile_shared=True, sheet_cfg=None, dedupe_strictness="deep", min_match_rate=0.25,
        max_total_loss=0.75, enforce_guards=True, enforce_assertions=True,
        log=lambda m: None):
    """Transform `swagger` into `out`. Returns a results dict.

    Guards (D6): the run aborts before writing if the mapping matched less
    than `min_match_rate` of the specification's attributes, or if more than
    `max_total_loss` of them would be eliminated. Pass
    `enforce_guards=False` to downgrade both to warnings.
    """
    on_unmatched = on_unmatched or "keep"
    mapping_display = os.path.basename(mapping) if isinstance(mapping, str) else str(mapping)
    log(f"Reading SOR mapping: {mapping_display}")
    m = load_mapping(mapping, sheet_cfg)
    for lay in m.layouts:
        cols = ", ".join(f"{k}=col{v}" for k, v in sorted(lay["cols"].items()))
        log(f"  layout [{lay['sheet']}]: mode={lay['mode']} header_row={lay['header_row']} {cols}"
            + (f" levels={lay['levels']}" if lay["levels"] else ""))
    log(f"  {m.stats['attribute_rows']} attribute row(s) indexed from "
        f"{m.stats['sheets']} sheet(s): {m.stats['mapped']} mapped, "
        f"{m.stats['unmapped']} marked not-used.")
    if m.stats["attribute_rows"] == 0:
        raise MappingCoverageError(
            "No attribute rows were read from the mapping workbook. The header row or "
            "column layout was not recognised — check the 'Detected spreadsheet layout' "
            "section, or set the columns explicitly.")

    sor_swaggers = list(sor_swaggers or [])
    sor_fields = collect_sor_fields(sor_swaggers) if sor_swaggers else None
    if sor_fields is not None:
        log(f"  {len(sor_fields)} SOR field name(s) loaded from {len(sor_swaggers)} SOR YAML file(s).")

    doc, yaml, src_encoding = _load_doc(swagger)
    if not src_encoding.startswith("utf-8"):
        log(f"  note: input decoded as {src_encoding}; the output is written as UTF-8.")
    schemas = (doc.get("components") or {}).get("schemas") or {}
    validation = validate_mapping(m, sor_fields) if sor_fields is not None else None
    if validation and validation["missing"]:
        log(f"  SOR YAML check: {len(validation['missing'])} mapped field(s) not found in the SOR YAML(s).")

    # D5 — decide de-duplication from the untouched document, before pruning.
    dedupe_candidates = (list(dedupe_pairs) if dedupe_pairs
                         else (detect_dedupe_candidates(schemas, dedupe_strictness)
                               if auto_dedupe else []))
    if dedupe_candidates:
        log(f"  {len(dedupe_candidates)} de-duplication candidate(s) identified in the source document.")

    # D3 — prune, keeping anything the mapping does not positively reject.
    diag = prune_document(doc, m, on_unmatched=on_unmatched, protected=protected,
                          sor_fields=sor_fields, require_sor_field=require_sor_field, log=log)
    dropped = diag["dropped"]
    log(f"Eliminated {len(dropped)} attribute(s) with a positive no-SOR signal; "
        f"kept {diag['attributes_unmatched']} not found in the mapping.")

    # D6 — circuit breakers, before anything is written.
    total = diag["attributes_total"] or 1
    loss = len(dropped) / total
    guard_msgs = []
    if diag["match_rate"] < min_match_rate:
        guard_msgs.append(
            f"Mapping matched only {diag['match_rate'] * 100:.1f}% of the "
            f"{diag['attributes_total']} attributes in the specification (floor is "
            f"{min_match_rate * 100:.0f}%). This is a lookup failure, not a filter: check "
            f"the detected spreadsheet layout and the field-name column.")
    if loss > max_total_loss:
        guard_msgs.append(
            f"The run would eliminate {loss * 100:.1f}% of all attributes "
            f"(ceiling is {max_total_loss * 100:.0f}%).")
    for msg in guard_msgs:
        log("  GUARD: " + msg)
    if guard_msgs and enforce_guards:
        raise (MappingCoverageError if diag["match_rate"] < min_match_rate
               else ExcessiveLossError)(" ".join(guard_msgs))

    for name, st in sorted(diag["per_schema"].items()):
        if st["total"] and st["removed"] == st["total"]:
            log(f"  note: '{name}' lost all {st['total']} attribute(s); it will be cascade-removed.")

    # D8 — specialise shared schemas per operation BEFORE the cascade, so that
    # positions mapping nothing are removed and the cascade can then tidy up
    # anything they left behind.
    profiling = {"actions": [], "created": [], "removed_positions": []}
    if profile_shared:
        profiling = profile_shared_schemas(
            doc, m, on_unmatched=on_unmatched, protected=protected,
            sor_fields=sor_fields, require_sor_field=require_sor_field, log=log)

    # D4 — cascade: no hollow objects left behind.
    cascade = cascade_prune(doc, protected=protected, log=log)

    # D5 — apply de-duplication, re-verifying equivalence.
    dedupe_log = []
    for owner, prop, target in dedupe_candidates:
        dedupe_inline(schemas, owner, prop, target, dedupe_log, dedupe_strictness)

    poly_log, n_bases = [], 0
    groups = list(poly_groups or [])
    if auto_poly and not groups:
        groups = auto_detect_groups(doc, discriminator_prop)
    for g in groups:
        sub = []
        if refactor_polymorphism(schemas, g["members"], g["base"],
                                 g.get("disc_prop", discriminator_prop),
                                 add_discriminator, sub):
            poly_log += sub
            n_bases += 1
        else:
            log(f"  Group '{g['base']}' skipped (no shared identical core).")

    result = {"dropped": dropped, "diagnostics": diag, "cascade": cascade,
              "layouts": m.layouts, "mapping_stats": m.stats,
              "poly_log": poly_log, "dedupe_log": dedupe_log, "n_bases": n_bases,
              "profiling": profiling,
              "validation": validation, "out": out, "guards": guard_msgs}

    # D6 — assert before declaring success.
    problems = assert_output(doc, protected)
    result["assertions"] = problems
    if problems:
        log(f"  ASSERTION FAILURES ({len(problems)}):")
        for p in problems[:25]:
            log(f"    ! {p}")

    annotate_document(doc, result)
    with open(out, "w", encoding="utf-8", newline="\n") as fh:
        yaml.dump(doc, fh)
    log(f"Wrote {out}  (UTF-8)")

    # The transformed document is already on disk. A failure while producing a
    # report is therefore a reporting problem, not a transform problem, and
    # must not discard a good run.
    result["report_errors"] = []
    try:
        report_txt = build_report(swagger, mapping_display, out, result)
    except Exception as e:                                    # noqa: BLE001
        report_txt = f"Report generation failed: {type(e).__name__}: {e}\n"
        result["report_errors"].append(f"markdown report: {type(e).__name__}: {e}")
        log("  WARNING: could not build the Markdown report: " f"{type(e).__name__}: {e}")
    result["report"] = report_txt
    if report:
        try:
            _write(report, report_txt)
            log(f"Wrote {report}")
        except Exception as e:                                # noqa: BLE001
            result["report_errors"].append(f"markdown report: {type(e).__name__}: {e}")
            log(f"  WARNING: could not write {report}: {type(e).__name__}: {e}")
    result["html"] = None
    if html_report:
        try:
            result["html"] = write_html_report(swagger, out, result, html_report)
            log(f"Wrote {html_report}")
        except Exception as e:                                # noqa: BLE001
            result["report_errors"].append(f"HTML report: {type(e).__name__}: {e}")
            log(f"  WARNING: could not write the HTML change report "
                f"({type(e).__name__}: {e}). The transformed YAML at {out} is unaffected.")

    if problems and enforce_assertions:
        raise OutputAssertionError(
            f"{len(problems)} post-run assertion(s) failed; the output and reports were "
            f"written for inspection. First: {problems[0]}")
    return result


# --------------------------------------------------------------------------- #
# Parsers shared by CLI + GUI + batch
# --------------------------------------------------------------------------- #
def parse_group_spec(text):
    """'Base = SchemaA:valA, SchemaB:valB' -> {'base', 'members'}."""
    if "=" not in text:
        raise ValueError(f"group must be 'Base = SchemaA:val, SchemaB:val' (got: {text!r})")
    base, rest = text.split("=", 1)
    members = []
    for part in rest.split(","):
        part = part.strip()
        if not part:
            continue
        if ":" in part:
            schema, val = part.split(":", 1)
            members.append((schema.strip(), val.strip()))
        else:
            members.append((part, norm(part)))
    return {"base": base.strip(), "members": members}


def parse_dedupe_spec(text):
    """'Owner.prop = TargetSchema' -> (owner, prop, target)."""
    if "=" not in text or "." not in text.split("=", 1)[0]:
        raise ValueError(f"dedupe must be 'Owner.prop = Target' (got: {text!r})")
    left, target = text.split("=", 1)
    owner, prop = left.strip().split(".", 1)
    return (owner.strip(), prop.strip(), target.strip())
