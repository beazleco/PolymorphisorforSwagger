"""
polymorphize_sor — verification against the System of Record specification.

Version 6.8.

The question this module answers
===============================

The workbook says an interface element is backed by an SOR field by naming that
field in an SOR column. That is an assertion, not evidence. Until the SOR's own
specification is consulted, a typo, a renamed field or a stale mapping all look
exactly like a genuine mapping, and the element is published as though the SOR
supported it. Publishing an element the SOR cannot supply is the defect this
whole tool exists to remove, so the assertion has to be checked.

The check is: take the SOR endpoint named for that column, take the message the
row belongs to, and look for the named field there.

The four rules, as settled with the analyst
===========================================

1. **An unresolved field is excluded from the interface**, exactly as an empty
   SOR cell is, and every exclusion is reported with the sheet, the cell, the
   field name and the endpoint that was searched.
2. **Matching is on the leaf name alone.** ``accountInfo.currency`` and
   ``currency`` both resolve to any element named ``currency`` anywhere in the
   message. This is the most forgiving rule and was chosen deliberately; the
   cost is that a name occurring in more than one place is ambiguous, so those
   are resolved and reported rather than silently taken.
3. **The ``SOR API Endpoint`` banner cell is definitive.** It is the
   declaration of record. An endpoint written on an SOR column header is a
   label and is ignored for this purpose; where the two differ, the banner
   wins and the difference is reported so the workbook can be tidied.
4. **Request Body is checked against the SOR request**, meaning its request
   body plus its query and path parameters, and **Response Body against the
   SOR response**.

Without an SOR specification
============================

Verification is optional. With no SOR files supplied the tool behaves as it
did before, taking the workbook's word, and says so in every report. That is a
weaker guarantee and the reports label it as such.
"""

from __future__ import annotations

import os
import re
from collections import defaultdict
from dataclasses import dataclass, field

import polymorphize_core as core
import polymorphize_workbook as wbk
from polymorphize_workbook import Finding, cell_ref, col_letter, key, text

__version__ = "6.8"

#: Methods whose request is carried by parameters rather than a body.
PARAMETER_METHODS = {"get", "delete", "head", "options"}

#: How a cell naming several SOR fields is split. Both separators appear in the
#: reference workbook: a comma for alternatives, a space for several fields
#: contributing to one element.
CANDIDATE_SPLIT = re.compile(r"[,;]|\s+")

MAX_SCHEMA_DEPTH = 12


# --------------------------------------------------------------------------- #
# Paths
# --------------------------------------------------------------------------- #


def normalise_path(raw):
    """A path in the one form used for every comparison.

    The banner in the reference workbook writes the same endpoint six ways:
    with a method and a colon, with a method and no colon, with neither, with
    and without a leading slash, and with a path parameter whose name need not
    match the one in the SOR specification. All of that is folded away, and
    ``{cardNumber}`` becomes ``{}`` so a template matches whatever the SOR
    calls its parameter.
    """
    s = text(raw)
    s = re.sub(r"^\s*(GET|POST|PUT|PATCH|DELETE|HEAD|OPTIONS)\s*:?\s*", "", s,
               flags=re.I)
    s = s.strip().rstrip("/")
    if not s:
        return ""
    if not s.startswith("/"):
        s = "/" + s
    s = re.sub(r"\{[^}]*\}", "{}", s)
    return s


def split_endpoints(raw):
    """The banner cell as a list of endpoints.

    Analysts separate them by newline, by comma, or by nothing but a space, so
    all three are handled. A space is only a separator where the next token
    starts a new path or names a method.
    """
    s = text(raw).replace("\n", " ")
    if not s:
        return []
    parts = re.split(r"[,;]", s)
    out = []
    for part in parts:
        tokens = part.split()
        current = []
        for token in tokens:
            starts = (token.startswith("/") or
                      re.match(r"^(GET|POST|PUT|PATCH|DELETE|HEAD|OPTIONS):?$",
                               token, re.I) or
                      re.match(r"^v\d+/", token))
            if starts and current and any(
                    t.startswith("/") or re.match(r"^v\d+/", t) for t in current):
                out.append(" ".join(current))
                current = [token]
            else:
                current.append(token)
        if current:
            out.append(" ".join(current))
    return [p for p in (normalise_path(x) for x in out) if p]


# --------------------------------------------------------------------------- #
# The index
# --------------------------------------------------------------------------- #


@dataclass
class MessageIndex:
    """Leaf names of one message of one SOR operation."""

    #: leaf name, lower case -> the dotted paths it occurs at
    leaves: dict = field(default_factory=lambda: defaultdict(list))
    #: total leaves, for the report
    total: int = 0

    def find(self, leaf):
        return self.leaves.get(leaf.lower(), [])


@dataclass
class OperationIndex:
    path: str
    method: str
    request: MessageIndex = field(default_factory=MessageIndex)
    response: MessageIndex = field(default_factory=MessageIndex)
    source: str = ""

    def side(self, name):
        return self.request if name == "request" else self.response


class SorIndex:
    """Every operation of every supplied SOR specification."""

    def __init__(self):
        self.operations = {}        # (normalised path, method) -> OperationIndex
        self.by_path = defaultdict(list)
        self.sources = []
        self.load_errors = []

    # -- construction -----------------------------------------------------

    @classmethod
    def from_paths(cls, paths, *, log=None):
        """Load one or more specification files, or every file in a folder."""
        index = cls()
        for entry in _expand(paths):
            try:
                doc = _load(entry)
            except Exception as exc:                           # noqa: BLE001
                index.load_errors.append((entry, "%s: %s"
                                          % (type(exc).__name__, exc)))
                continue
            if not isinstance(doc, dict) or "paths" not in doc:
                index.load_errors.append(
                    (entry, "no paths section, so this is not an OpenAPI or "
                            "Swagger specification"))
                continue
            index.sources.append(entry)
            index._ingest(doc, entry)
            if log:
                log("  read   %-40s %d operation(s)"
                    % (os.path.basename(entry), len(doc.get("paths") or {})))
        return index

    def _ingest(self, doc, source):
        root_params = doc.get("paths") or {}
        for raw_path, item in root_params.items():
            if not isinstance(item, dict):
                continue
            norm = normalise_path(raw_path)
            shared = item.get("parameters") or []
            for method, operation in item.items():
                if method.lower() not in ("get", "post", "put", "patch",
                                          "delete", "head", "options"):
                    continue
                if not isinstance(operation, dict):
                    continue
                op = OperationIndex(norm, method.lower(), source=source)
                params = list(shared) + list(operation.get("parameters") or [])
                _index_parameters(params, doc, op.request)
                _index_body(operation, doc, op.request)
                _index_responses(operation, doc, op.response)
                self.operations[(norm, method.lower())] = op
                self.by_path[norm].append(op)

    # -- lookup -----------------------------------------------------------

    def has_path(self, path):
        return normalise_path(path) in self.by_path

    def operations_for(self, path):
        return self.by_path.get(normalise_path(path), [])

    def pick(self, path, method=None):
        """The operation for a path, and the method when one is stated."""
        norm = normalise_path(path)
        if method and (norm, method.lower()) in self.operations:
            return self.operations[(norm, method.lower())]
        ops = self.by_path.get(norm) or []
        return ops[0] if len(ops) == 1 else (ops[0] if ops else None)

    @property
    def empty(self):
        return not self.operations

    def summary(self):
        return {
            "files": len(self.sources),
            "operations": len(self.operations),
            "paths": len(self.by_path),
            "load_errors": len(self.load_errors),
        }


def _expand(paths):
    """A list of files, expanding any folder into the specifications inside it."""
    out = []
    for p in ([paths] if isinstance(paths, str) else list(paths or [])):
        p = str(p).strip()
        if not p:
            continue
        if os.path.isdir(p):
            for name in sorted(os.listdir(p)):
                if name.lower().endswith((".yaml", ".yml", ".json")):
                    out.append(os.path.join(p, name))
        else:
            out.append(p)
    return out


def _load(path):
    """One specification, YAML or JSON, read through the encoding sniffer."""
    import io

    from ruamel.yaml import YAML
    text_in, _encoding = core.read_text(path)
    yaml = YAML(typ="safe")
    yaml.allow_duplicate_keys = True
    return yaml.load(io.StringIO(text_in))


# --------------------------------------------------------------------------- #
# Walking a specification's schemas
# --------------------------------------------------------------------------- #


def _deref(node, doc, seen=()):
    """Follow a local ``$ref``. A missing or remote target yields nothing."""
    depth = 0
    while isinstance(node, dict) and "$ref" in node and depth < 20:
        ref = node["$ref"]
        if not isinstance(ref, str) or not ref.startswith("#/"):
            return {}
        if ref in seen:
            return {}
        seen = tuple(seen) + (ref,)
        target = doc
        for part in ref[2:].split("/"):
            part = part.replace("~1", "/").replace("~0", "~")
            if isinstance(target, dict) and part in target:
                target = target[part]
            else:
                return {}
        node = target
        depth += 1
    return node if isinstance(node, dict) else {}


def _walk_schema(schema, doc, message, prefix=(), depth=0, seen=frozenset()):
    """Record every leaf name of a schema against the paths it occurs at."""
    if depth > MAX_SCHEMA_DEPTH or not isinstance(schema, dict):
        return
    if "$ref" in schema:
        ref = schema["$ref"]
        if ref in seen:
            return
        _walk_schema(_deref(schema, doc), doc, message, prefix, depth + 1,
                     seen | {ref})
        return
    for keyword in ("allOf", "oneOf", "anyOf"):
        if isinstance(schema.get(keyword), list):
            for sub in schema[keyword]:
                _walk_schema(sub, doc, message, prefix, depth + 1, seen)
    if schema.get("type") == "array" or "items" in schema:
        _walk_schema(schema.get("items") or {}, doc, message, prefix,
                     depth + 1, seen)
        return
    props = schema.get("properties")
    if isinstance(props, dict):
        for name, sub in props.items():
            path = prefix + (name,)
            sub = sub if isinstance(sub, dict) else {}
            resolved = _deref(sub, doc) if "$ref" in sub else sub
            has_children = bool(
                (resolved.get("properties")) or
                any(isinstance(resolved.get(k), list)
                    for k in ("allOf", "oneOf", "anyOf")) or
                resolved.get("type") == "array" or "items" in resolved)
            # A container's own name is recorded too: an analyst legitimately
            # maps an array or an object, for instance an array named cards.
            message.leaves[name.lower()].append(".".join(path))
            message.total += 1
            if has_children:
                _walk_schema(sub, doc, message, path, depth + 1, seen)
    add = schema.get("additionalProperties")
    if isinstance(add, dict):
        _walk_schema(add, doc, message, prefix, depth + 1, seen)


def _index_parameters(params, doc, message):
    """Query, path and header parameters, which are a GET's request."""
    for raw in params or []:
        p = _deref(raw, doc) if isinstance(raw, dict) and "$ref" in raw else raw
        if not isinstance(p, dict):
            continue
        name = p.get("name")
        if name:
            message.leaves[str(name).lower()].append(str(name))
            message.total += 1
        schema = p.get("schema")
        if isinstance(schema, dict):
            _walk_schema(schema, doc, message, (str(name),) if name else ())


def _index_body(operation, doc, message):
    body = operation.get("requestBody")
    if isinstance(body, dict):
        body = _deref(body, doc) if "$ref" in body else body
        for _mt, holder in (body.get("content") or {}).items():
            if isinstance(holder, dict) and isinstance(holder.get("schema"), dict):
                _walk_schema(holder["schema"], doc, message)
    # Swagger 2.0 keeps the body among the parameters.
    for raw in operation.get("parameters") or []:
        p = _deref(raw, doc) if isinstance(raw, dict) and "$ref" in raw else raw
        if isinstance(p, dict) and p.get("in") == "body" and \
                isinstance(p.get("schema"), dict):
            _walk_schema(p["schema"], doc, message)


def _index_responses(operation, doc, message):
    """The success response. A 2xx is preferred, then default."""
    responses = operation.get("responses")
    if not isinstance(responses, dict):
        return
    chosen = [c for c in responses if str(c).startswith("2")]
    chosen = chosen or [c for c in responses if str(c) == "default"]
    chosen = chosen or list(responses)[:1]
    for code in chosen:
        holder = responses.get(code)
        holder = _deref(holder, doc) if isinstance(holder, dict) and \
            "$ref" in holder else holder
        if not isinstance(holder, dict):
            continue
        if isinstance(holder.get("schema"), dict):          # Swagger 2.0
            _walk_schema(holder["schema"], doc, message)
        for _mt, content in (holder.get("content") or {}).items():
            if isinstance(content, dict) and isinstance(content.get("schema"), dict):
                _walk_schema(content["schema"], doc, message)


# --------------------------------------------------------------------------- #
# Endpoint identification
# --------------------------------------------------------------------------- #


@dataclass
class EndpointDecision:
    column: object            # wbk.SorColumn
    endpoint: str             # normalised, or "" when unusable
    source: str               # "column" | "banner" | "conflict" | "none"
    banner: str = ""
    stated_method: str = ""


def stated_method(raw):
    m = re.match(r"^\s*(GET|POST|PUT|PATCH|DELETE|HEAD|OPTIONS)\b",
                 text(raw), re.I)
    return m.group(1).lower() if m else ""


def decide_endpoints(result, findings):
    """Which SOR endpoint each column refers to.

    The ``SOR API Endpoint`` banner cell is definitive. An endpoint written on
    a column header is a label and does not decide anything: it is compared
    with the banner only so a difference can be reported.

    One banner endpoint serves every column. Several are paired with the
    columns in order, which is the same positional convention the
    ``requestVariant`` values already use, and a mismatch in the counts is an
    error rather than a guess.
    """
    layout = result.layout
    banner_raw = result.banner.get("sor_endpoint", "")
    banner_list = split_endpoints(banner_raw)
    columns = layout.sor_cols
    label_cell = cell_ref(layout.header_row + 2, layout.label_cols[0]
                          if layout.label_cols else 0)
    out = []

    if not banner_list:
        findings.append(Finding(
            "R003", "error", result.sheet, label_cell,
            "the SOR API Endpoint banner cell is empty, and it is the only "
            "place the tool reads the SOR endpoint from",
            "the SOR endpoint, or one per SOR column, in the SOR API Endpoint "
            "banner cell",
            "fill in the SOR API Endpoint cell. A column header such as "
            "%r is a label and is not read for this purpose."
            % (columns[0].header if columns else "SOR /v1/..."),
        ))
        return [EndpointDecision(c, "", "none", banner_raw) for c in columns]

    if len(banner_list) != 1 and len(banner_list) != len(columns):
        findings.append(Finding(
            "R002", "error", result.sheet, label_cell,
            "the SOR API Endpoint banner names %d endpoints but the sheet has "
            "%d SOR column%s, so they cannot be paired"
            % (len(banner_list), len(columns), "" if len(columns) == 1 else "s"),
            "one endpoint in the banner, or exactly one per SOR column in the "
            "order the columns appear",
            "list one endpoint per SOR column in the SOR API Endpoint cell, in "
            "the same left-to-right order as the columns, or reduce it to the "
            "single endpoint this sheet calls",
        ))
        return [EndpointDecision(c, "", "conflict", banner_raw) for c in columns]

    for i, column in enumerate(columns):
        endpoint = banner_list[0] if len(banner_list) == 1 else banner_list[i]
        stated = banner_list[0] if len(banner_list) == 1 else banner_list[i]
        raw_for_method = (banner_raw if len(banner_list) == 1
                          else _banner_token(banner_raw, i))
        header_endpoint = normalise_path(column.endpoint)
        if header_endpoint and header_endpoint != endpoint:
            findings.append(Finding(
                "R001", "warning", result.sheet,
                cell_ref(layout.header_row + 1, column.index),
                "the column header names %r but the SOR API Endpoint banner "
                "names %r for this column, and the banner is definitive, so %r "
                "was used" % (header_endpoint, endpoint, endpoint),
                "the column header and the banner to name the same endpoint",
                "correct the column header to %r, or correct the banner. Only "
                "the banner is read, so the interface is unaffected either way, "
                "but a reader of the workbook will be misled."
                % ("SOR " + endpoint),
            ))
        out.append(EndpointDecision(column, endpoint, "banner", banner_raw,
                                    stated_method(raw_for_method)))
    return out


def _banner_token(banner_raw, index):
    """The nth endpoint of the banner as written, method prefix and all."""
    s = text(banner_raw).replace("\n", " ")
    parts = re.split(r"[,;]", s)
    tokens = []
    for part in parts:
        words = part.split()
        current = []
        for word in words:
            starts = (word.startswith("/") or
                      re.match(r"^(GET|POST|PUT|PATCH|DELETE|HEAD|OPTIONS):?$",
                               word, re.I) or
                      re.match(r"^v\d+/", word))
            if starts and current and any(
                    t.startswith("/") or re.match(r"^v\d+/", t) for t in current):
                tokens.append(" ".join(current))
                current = [word]
            else:
                current.append(word)
        if current:
            tokens.append(" ".join(current))
    return tokens[index] if index < len(tokens) else ""


# --------------------------------------------------------------------------- #
# Resolving one field name
# --------------------------------------------------------------------------- #

RESOLVED = "resolved"
AMBIGUOUS = "ambiguous"
MISSING = "missing"
UNCHECKED = "unchecked"
NO_ENDPOINT = "no-endpoint"


@dataclass
class Resolution:
    status: str
    declared: str = ""          # the cell value as written
    candidates: list = field(default_factory=list)
    matched: str = ""           # the leaf that matched
    at: list = field(default_factory=list)   # paths in the SOR message
    endpoint: str = ""
    method: str = ""
    side: str = ""

    @property
    def ok(self):
        return self.status in (RESOLVED, AMBIGUOUS, UNCHECKED)


def candidates_of(value):
    """A cell naming one or several SOR fields, as a list of leaf names."""
    out = []
    for token in CANDIDATE_SPLIT.split(text(value)):
        token = token.strip().strip(",;")
        if not token:
            continue
        leaf = token.split(".")[-1].strip()
        if leaf:
            out.append((token, leaf))
    return out


def resolve(value, op, side):
    """One SOR cell against one message of one SOR operation.

    Matching is on the leaf name alone, by decision. A cell naming several
    fields resolves when any of them does, which is right for both spellings
    the workbook uses: a comma means alternatives and a space means several
    fields feeding one element.
    """
    pairs = candidates_of(value)
    if not pairs:
        return Resolution(MISSING, declared=text(value))
    message = op.side(side)
    best = None
    for declared, leaf in pairs:
        at = message.find(leaf)
        if not at:
            continue
        status = AMBIGUOUS if len(at) > 1 else RESOLVED
        found = Resolution(status, declared=text(value),
                           candidates=[p[0] for p in pairs],
                           matched=leaf, at=at, endpoint=op.path,
                           method=op.method, side=side)
        if status == RESOLVED:
            return found
        best = best or found
    if best:
        return best
    return Resolution(MISSING, declared=text(value),
                      candidates=[p[0] for p in pairs],
                      endpoint=op.path, method=op.method, side=side)


# --------------------------------------------------------------------------- #
# Verifying a whole sheet
# --------------------------------------------------------------------------- #


@dataclass
class SheetVerification:
    sheet: str
    checked: bool = False
    decisions: list = field(default_factory=list)
    resolutions: dict = field(default_factory=dict)  # (id(node), endpoint) -> Resolution
    counts: dict = field(default_factory=dict)

    def of(self, node, endpoint):
        return self.resolutions.get((id(node), endpoint))


def verify_sheet(result, index, findings, *, request_side="request"):
    """Check every SOR cell of one sheet, and clear the ones that do not resolve.

    Clearing the mapping is what excludes the element: everything downstream
    already treats an empty SOR cell as "this endpoint does not supply it", so
    no other stage needs to know verification happened. The evidence is kept
    for the reports.
    """
    v = SheetVerification(sheet=result.sheet)
    if result.layout is None or result.status == "failed":
        return v
    if index is None or index.empty:
        return v

    v.decisions = decide_endpoints(result, findings)
    by_endpoint = {}
    for decision in v.decisions:
        if not decision.endpoint:
            v.checked = False
            return v
        ops = index.operations_for(decision.endpoint)
        if not ops:
            findings.append(Finding(
                "R004", "error", result.sheet,
                cell_ref(result.layout.header_row + 1, decision.column.index),
                "the SOR endpoint %r is not in any of the supplied SOR "
                "specifications" % decision.endpoint,
                "every SOR endpoint the workbook names to be present in the "
                "SOR specifications given to the tool",
                "supply the SOR specification that defines %r, or correct the "
                "endpoint in the workbook. Nothing can be verified for this "
                "operation until one of the two is done." % decision.endpoint,
            ))
            v.checked = False
            return v
        op = index.pick(decision.endpoint, decision.stated_method)
        if decision.stated_method and \
                (decision.endpoint, decision.stated_method) not in index.operations:
            findings.append(Finding(
                "R005", "warning", result.sheet,
                cell_ref(result.layout.header_row + 1, decision.column.index),
                "the workbook states %s for %r but the SOR specification "
                "defines only %s, so %s was used"
                % (decision.stated_method.upper(), decision.endpoint,
                   ", ".join(sorted(o.method.upper() for o in ops)),
                   op.method.upper()),
                "the method in the workbook to match the SOR specification",
                "correct the method in the SOR API Endpoint cell, or confirm "
                "that %s is the operation intended" % op.method.upper(),
            ))
        by_endpoint[decision.column.endpoint] = op

    v.checked = True
    counts = {RESOLVED: 0, AMBIGUOUS: 0, MISSING: 0}
    layout = result.layout

    for section, side in ((result.request, request_side),
                          (result.response, "response")):
        if section is None or section.root is None:
            continue

        def walk(node):
            for child in node.children:
                for column in layout.sor_cols:
                    declared = child.sor.get(column.endpoint)
                    if not declared:
                        continue
                    op = by_endpoint.get(column.endpoint)
                    if op is None:
                        continue
                    res = resolve(declared, op, side)
                    v.resolutions[(id(child), column.endpoint)] = res
                    counts[res.status] = counts.get(res.status, 0) + 1
                    if res.status == MISSING:
                        # This is the exclusion. The mapping is cleared, so the
                        # element is pruned exactly as an empty cell would be.
                        child.sor[column.endpoint] = ""
                        findings.append(Finding(
                            "R010", "warning", result.sheet,
                            cell_ref(child.row, column.index),
                            "%r names the SOR field %r, but no element called "
                            "%r exists in the %s of %s %s, so the element was "
                            "removed from the interface"
                            % (child.name, res.declared,
                               (res.candidates[0].split(".")[-1]
                                if res.candidates else res.declared),
                               side, op.method.upper(), op.path),
                            "an SOR field name that exists in that message of "
                            "that SOR endpoint",
                            "correct the field name in cell %s, or clear the "
                            "cell if the SOR genuinely does not supply %r. If "
                            "the field does exist, check that the sheet names "
                            "the right SOR endpoint."
                            % (cell_ref(child.row, column.index), child.name),
                        ))
                    elif res.status == AMBIGUOUS:
                        findings.append(Finding(
                            "R011", "warning", result.sheet,
                            cell_ref(child.row, column.index),
                            "%r matched the SOR field %r in %d different places "
                            "in the %s of %s %s: %s"
                            % (child.name, res.matched, len(res.at), side,
                               op.method.upper(), op.path,
                               ", ".join(res.at[:4])),
                            "a field name that identifies one element",
                            "write the fuller path in cell %s so it is clear "
                            "which one is meant. Matching is on the last "
                            "segment, so the element was kept, but a reader "
                            "cannot tell which SOR field feeds it."
                            % cell_ref(child.row, column.index),
                        ))
                walk(child)
        walk(section.root)

    v.counts = counts
    total = sum(counts.values())
    if total and counts[MISSING] == total:
        findings.append(Finding(
            "R020", "error", result.sheet, "",
            "not one of the %d SOR field names on this sheet exists in the SOR "
            "endpoint it names, which almost always means the sheet points at "
            "the wrong endpoint or the wrong SOR specification was supplied"
            % total,
            "at least one SOR field name to resolve",
            "check the SOR API Endpoint cell and the SOR column headers against "
            "the SOR specification. No specification was written for this "
            "operation, because verifying every element away would publish an "
            "empty interface.",
        ))
    elif counts[MISSING]:
        findings.append(Finding(
            "R021", "warning", result.sheet, "",
            "%d of %d SOR field names on this sheet do not exist in the SOR "
            "endpoint, so %d element%s removed from the interface"
            % (counts[MISSING], total, counts[MISSING],
               " was" if counts[MISSING] == 1 else "s were"),
            "every SOR field name in the workbook to exist in the SOR "
            "specification",
            "the R010 findings name each one with its cell. Correct them in the "
            "workbook, or accept the removals if the SOR really has changed.",
        ))
    return v


def verify_workbook(results, index, *, log=None):
    """Verify every sheet. Returns ``{sheet: SheetVerification}``."""
    out = {}
    for r in results:
        if r.status == "skipped":
            continue
        findings = []
        v = verify_sheet(r, index, findings)
        r.findings.extend(findings)
        if any(f.severity == "error" for f in findings):
            r.status = "failed"
        out[r.sheet] = v
        if log and v.checked:
            log("  check  %-34s %d resolved, %d ambiguous, %d removed"
                % (r.sheet, v.counts.get(RESOLVED, 0),
                   v.counts.get(AMBIGUOUS, 0), v.counts.get(MISSING, 0)))
    return out


__all__ = ["AMBIGUOUS", "EndpointDecision", "MISSING", "MessageIndex",
           "NO_ENDPOINT", "OperationIndex", "RESOLVED", "Resolution",
           "SheetVerification", "SorIndex", "UNCHECKED", "candidates_of",
           "decide_endpoints", "normalise_path", "resolve", "split_endpoints",
           "stated_method", "verify_sheet", "verify_workbook"]
