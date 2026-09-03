"""
polymorphize_generate — one OpenAPI specification per worksheet.

Version 6.3.

The workbook is authoritative. Each operation sheet describes one endpoint, so
one workbook yields one specification per sheet rather than one specification
overall. A sheet that cannot be read does not stop the others.

Polymorphism patterns
=====================

The shape of the output follows the number of SOR endpoints behind the
operation, because that is what decides whether the variance is settled at
design time or at run time.

P1, one SOR column
    One schema per message, carrying only the attributes that SOR supplies.
    No ``allOf``, no ``oneOf``: there is nothing to vary.

P2, several SOR columns with a ``requestVariant`` enumeration
    The caller names the variant, so the variance is genuine run-time
    polymorphism within one operation. A base schema holds the attributes
    common to every variant, each variant is ``allOf`` base plus its own
    delta, and the message schema is a ``oneOf`` over the variants with
    ``discriminator: requestVariant``. The discriminator property is retained
    and made required whatever the SOR columns say, because it is consumed by
    the API layer rather than by SOR, and each variant narrows its ``enum`` to
    the single code that selects it.

    The response varies by the same variant but carries no tag of its own, so
    the response is a ``oneOf`` without a discriminator. OpenAPI cannot state
    that the response subtype follows a request property; ``x-selected-by``
    records the dependency for a reader.

P3, several SOR columns with no ``requestVariant``
    The SOR endpoint is chosen by which identifier the caller supplies, so
    there is no property to discriminate on. The output is a ``oneOf`` without
    a discriminator over ``allOf`` derived schemas, which is valid OpenAPI and
    states the exclusivity without inventing a tag.
"""

from __future__ import annotations

import datetime as _dt
import io
import json
import os
import re
from dataclasses import dataclass, field

from ruamel.yaml import YAML

import polymorphize_core as core
import polymorphize_workbook as wbk
from polymorphize_workbook import (
    Finding, SECTION_REQUEST, VARIANT_PROPERTY, cell_ref, key, text,
)

__version__ = "6.3"

OPENAPI_VERSION = "3.0.3"

#: Sheet-name suffix to HTTP method.
METHOD_BY_SUFFIX = {
    "retrieve": "get",
    "retrieveall": "get",
    "list": "get",
    "initiate": "post",
    "create": "post",
    "register": "post",
    "request": "post",
    "execute": "post",
    "evaluate": "post",
    "update": "put",
    "amend": "put",
    "cancel": "put",
    "control": "put",
    "exchange": "post",
    "capture": "post",
    "notify": "post",
    "grant": "post",
}

METHODS_WITHOUT_BODY = {"get", "delete", "head"}

BQ_UNSET = {"", "n/a", "na", "none", "tbc", "tbd", "-"}


def _plain(node):
    """Plain dicts and lists only.

    ruamel serialises an ``OrderedDict`` as a ``!!omap`` tagged sequence, which
    is valid YAML and useless as an OpenAPI document. Ordinary dictionaries
    keep insertion order anyway, so the mapping type is dropped on the way out.
    """
    if isinstance(node, dict):
        return {k: _plain(v) for k, v in node.items()}
    if isinstance(node, (list, tuple)):
        return [_plain(v) for v in node]
    return node


def dump_document(document, fmt="yaml"):
    """A specification as text. Always UTF-8 clean, never a tab in sight."""
    document = _plain(document)
    if fmt == "json":
        return json.dumps(document, indent=2, ensure_ascii=False) + "\n"
    yaml = YAML()
    yaml.default_flow_style = False
    yaml.width = 100
    yaml.allow_unicode = True
    buf = io.StringIO()
    yaml.dump(document, buf)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# Naming
# --------------------------------------------------------------------------- #


def pascal(value):
    """``Account Contact Details`` -> ``AccountContactDetails``."""
    parts = re.split(r"[^A-Za-z0-9]+", text(value))
    return "".join(p[:1].upper() + p[1:] for p in parts if p)


def slug(value):
    """``Card Issued Device_Retrieve`` -> ``card-issued-device-retrieve``."""
    s = re.sub(r"[^A-Za-z0-9]+", "-", text(value)).strip("-").lower()
    return re.sub(r"-{2,}", "-", s)


def endpoint_slug(endpoint, ordinal):
    """An SOR endpoint path to a schema-name fragment."""
    parts = [p for p in re.split(r"[^A-Za-z0-9]+", text(endpoint)) if p]
    parts = [p for p in parts if not re.fullmatch(r"v\d+", p, re.I)]
    return pascal(" ".join(parts)) or "Variant%d" % ordinal


def derive_method(sheet, banner):
    """The HTTP method of the *business* API.

    The BIAN action term decides it: the last segment of the equivalent BIAN
    endpoint, else the suffix of the sheet name. The SOR API Endpoint banner
    cell is deliberately not consulted. It states the method of the downstream
    SOR call, which is a different API and frequently a different verb: the
    reference workbook fulfils an ``Initiate`` with ``PUT /v1/card/activate``
    and an ``Update`` with ``DELETE /v1/card/cancel``.
    """
    stated = re.match(r"^\s*(GET|POST|PUT|PATCH|DELETE)\b",
                      text(banner.get("method", "")), re.I)
    if stated:
        return stated.group(1).lower()
    for source in (banner.get("bian_endpoint", ""), sheet):
        tokens = [t for t in re.split(r"[^A-Za-z0-9]+", text(source)) if t]
        for token in reversed(tokens):
            hit = METHOD_BY_SUFFIX.get(key(token))
            if hit:
                return hit
    return "post"


def derive_path(sheet, banner):
    """The business path, from the banner when it states one."""
    for field_name in ("business_endpoint", "bian_endpoint"):
        raw = text(banner.get(field_name, ""))
        raw = re.sub(r"^(GET|POST|PUT|PATCH|DELETE)\s*:?\s*", "", raw, flags=re.I)
        raw = raw.split()[0] if raw.split() else ""
        if raw and key(raw) not in BQ_UNSET:
            return "/" + raw.strip("/")
    sd = slug(banner.get("service_domain", "")) or "service-domain"
    return "/%s/%s" % (sd, slug(sheet) or "operation")


# --------------------------------------------------------------------------- #
# Pruning
# --------------------------------------------------------------------------- #


def supported(node, endpoints):
    """True when the node, or anything beneath it, is supplied by ``endpoints``."""
    if any(node.sor.get(e) for e in endpoints):
        return True
    return any(supported(c, endpoints) for c in node.children)


def retained(node, endpoints, *, protect=()):
    """Children of ``node`` that survive pruning against ``endpoints``."""
    out = []
    for c in node.children:
        if c.name in protect or supported(c, endpoints):
            out.append(c)
    return out


# --------------------------------------------------------------------------- #
# Schema emission
# --------------------------------------------------------------------------- #


def _scalar(node, endpoints, *, variant_codes=None):
    """The schema fragment for one leaf."""
    out = {"type": node.json_type}
    if node.fmt:
        out["format"] = node.fmt
    if node.max_length and node.json_type == "string":
        out["maxLength"] = node.max_length
    if node.name == VARIANT_PROPERTY and variant_codes:
        out["enum"] = list(variant_codes)
    if node.description:
        out["description"] = node.description
    if node.example:
        out["example"] = node.example
    if node.usage == wbk.USAGE_CONDITIONAL:
        out["x-conditional"] = True
    mapped = {e: node.sor[e] for e in endpoints if node.sor.get(e)}
    if mapped:
        if len(mapped) == 1 and len(endpoints) == 1:
            out["x-sor-field"] = next(iter(mapped.values()))
        else:
            out["x-sor-fields"] = mapped
    return out


def emit(node, endpoints, *, protect=(), variant_codes=None):
    """One node and everything retained beneath it, as an OpenAPI schema."""
    if node.json_type == "array":
        kids = retained(node, endpoints, protect=protect)
        item = {"type": "object", "properties": {}}
        props, required = {}, []
        for c in kids:
            props[c.name] = emit(c, endpoints, protect=protect,
                                 variant_codes=variant_codes)
            if c.usage == wbk.USAGE_REQUIRED or c.name in protect:
                required.append(c.name)
        # A single object child is the array's element, not a property of it.
        if len(kids) == 1 and kids[0].is_container and kids[0].json_type == "object":
            item = props[kids[0].name]
        else:
            item["properties"] = props
            if required:
                item["required"] = required
        out = {"type": "array", "items": item}
        if node.description:
            out["description"] = node.description
        mapped = {e: node.sor[e] for e in endpoints if node.sor.get(e)}
        if mapped:
            out["x-sor-fields"] = mapped
        return out

    if node.json_type == "object":
        props, required = {}, []
        for c in retained(node, endpoints, protect=protect):
            props[c.name] = emit(c, endpoints, protect=protect,
                                 variant_codes=variant_codes)
            if c.usage == wbk.USAGE_REQUIRED or c.name in protect:
                required.append(c.name)
        out = {"type": "object", "properties": props}
        if required:
            out["required"] = required
        if node.description:
            out["description"] = node.description
        return out

    return _scalar(node, endpoints, variant_codes=variant_codes)


def _narrow_variant_description(schema, variant):
    """On a derived schema, the tag documents only the variant it selects."""
    if not isinstance(schema, dict):
        return
    props = schema.get("properties")
    if isinstance(props, dict) and VARIANT_PROPERTY in props:
        entry = props[VARIANT_PROPERTY]
        if isinstance(entry, dict):
            entry["description"] = "%s. %s" % (variant.code, variant.label)
    for value in (props or {}).values():
        _narrow_variant_description(value, variant)


def _leaf_signature(node, endpoints, protect):
    """Every retained path beneath a node, for intersection arithmetic."""
    out = set()

    def walk(n, path):
        for c in retained(n, endpoints, protect=protect):
            p = path + (c.name,)
            if c.is_container:
                walk(c, p)
            else:
                out.add(p)
    walk(node, ())
    return out


def _subtract(schema, keep_paths, path=()):
    """Strip from ``schema`` every leaf whose path is not in ``keep_paths``."""
    if not isinstance(schema, dict):
        return schema
    if schema.get("type") == "array":
        item = schema.get("items")
        if isinstance(item, dict):
            _subtract(item, keep_paths, path)
        return schema
    props = schema.get("properties")
    if not isinstance(props, dict):
        return schema
    for name in list(props):
        child = props[name]
        p = path + (name,)
        if isinstance(child, dict) and (child.get("properties") or
                                        child.get("type") == "array"):
            _subtract(child, keep_paths, p)
            inner = child.get("items", child) if child.get("type") == "array" else child
            if isinstance(inner, dict) and not inner.get("properties") and \
                    inner.get("type") in (None, "object"):
                del props[name]
                if name in schema.get("required", []):
                    schema["required"].remove(name)
        elif p not in keep_paths:
            del props[name]
            if name in schema.get("required", []):
                schema["required"].remove(name)
    if "required" in schema and not schema["required"]:
        del schema["required"]
    return schema


# --------------------------------------------------------------------------- #
# Variants
# --------------------------------------------------------------------------- #


@dataclass
class Variant:
    code: str
    label: str
    endpoint: str
    name_fragment: str


def pair_variants(result, findings):
    """Match declared ``requestVariant`` values to SOR columns.

    Correspondence is checked by name before position, so a reordered column
    is caught rather than silently mismatched.
    """
    cols = result.layout.sor_cols
    declared = result.variants
    out = []
    if declared and len(declared) == len(cols):
        for i, ((code, label), col) in enumerate(zip(declared, cols), 1):
            lk, ek = key(label), key(col.endpoint)
            overlap = sum(1 for tok in re.split(r"[^a-z0-9]+", label.lower())
                          if tok and tok[:5] in ek)
            if ek and not overlap:
                findings.append(Finding(
                    "V003", "warning", result.sheet,
                    cell_ref(result.layout.header_row + 1, col.index),
                    "variant %s %r was paired with SOR column %r by position, but "
                    "the two names have nothing in common" % (code, label, col.header),
                    "the nth variant to describe the nth SOR column",
                    "reorder the SOR columns to match the order of the "
                    "%s values, or reword one of them so the pairing is obvious"
                    % VARIANT_PROPERTY,
                ))
            out.append(Variant(code, label, col.endpoint,
                               pascal(label) or endpoint_slug(col.endpoint, i)))
        return out
    for i, col in enumerate(cols, 1):
        out.append(Variant("%02d" % i, col.endpoint or "Variant %d" % i,
                           col.endpoint, endpoint_slug(col.endpoint, i)))
    return out


# --------------------------------------------------------------------------- #
# One sheet to one specification
# --------------------------------------------------------------------------- #


@dataclass
class SheetSpec:
    sheet: str
    status: str = "ok"
    pattern: str = ""
    document: dict = None
    path: str = ""
    method: str = ""
    findings: list = field(default_factory=list)
    stats: dict = field(default_factory=dict)
    variants: list = field(default_factory=list)


def _message(result, section, variants, pattern, schemas, *, is_request):
    """Add every schema for one message to ``schemas`` and name the entry."""
    root = section.root
    base_name = pascal(root.name) or ("Request" if is_request else "Response")
    endpoints = [c.endpoint for c in result.layout.sor_cols]
    protect = (VARIANT_PROPERTY,) if is_request else ()
    codes = [v.code for v in variants] if variants else None

    if pattern == "P1":
        schemas[base_name] = emit(root, endpoints, protect=protect)
        return base_name, {base_name: len(_leaf_signature(root, endpoints, protect))}

    common = [v.endpoint for v in variants]
    inter = None
    per = {}
    for v in variants:
        sig = _leaf_signature(root, [v.endpoint], protect)
        per[v.code] = sig
        inter = sig if inter is None else (inter & sig)
    inter = inter or set()

    core_name = base_name + "Base"
    base_schema = emit(root, common, protect=protect,
                       variant_codes=codes if is_request else None)
    _subtract(base_schema, inter)
    root_note = (root.description + " ") if root.description else ""
    base_schema["description"] = (
        root_note + "Attributes common to every variant of %s." % base_name)
    schemas[core_name] = base_schema

    refs, mapping, counts = [], {}, {}
    for v in variants:
        name = base_name + "For" + v.name_fragment
        delta = per[v.code] - inter
        derived = emit(root, [v.endpoint], protect=protect,
                       variant_codes=[v.code] if is_request else None)
        _subtract(derived, delta | ({(VARIANT_PROPERTY,)} if is_request else set()))
        # The wrapper carries the variant's own description; repeating the
        # schema-level one inside the allOf adds nothing.
        derived.pop("description", None)
        if is_request:
            derived.setdefault("required", [])
            if VARIANT_PROPERTY not in derived["required"]:
                derived["required"].append(VARIANT_PROPERTY)
            _narrow_variant_description(derived, v)
        entry = {
            "allOf": [
                {"$ref": "#/components/schemas/%s" % core_name},
                derived,
            ],
            "description": "%s as served by %s." % (base_name, v.endpoint or v.label),
            "x-sor-endpoint": v.endpoint,
            "x-variant": v.code,
            "x-variant-name": v.label,
        }
        schemas[name] = entry
        refs.append({"$ref": "#/components/schemas/%s" % name})
        mapping[v.code] = "#/components/schemas/%s" % name
        counts[name] = len(per[v.code])

    wrapper = {"oneOf": refs}
    if is_request and pattern == "P2":
        wrapper["discriminator"] = {"propertyName": VARIANT_PROPERTY,
                                    "mapping": mapping}
    elif not is_request and pattern == "P2":
        wrapper["x-selected-by"] = VARIANT_PROPERTY
        wrapper["description"] = (
            "The subtype returned follows the %s sent on the request. OpenAPI "
            "cannot express a response discriminator that depends on a request "
            "property, so the dependency is recorded here."
            % VARIANT_PROPERTY)
    else:
        wrapper["description"] = (
            "Exactly one subtype applies. The SOR endpoint is selected by which "
            "identifier the caller supplies, so there is no property to "
            "discriminate on.")
    schemas[base_name] = wrapper
    counts[core_name] = len(inter)
    return base_name, counts


def generate_sheet(result, *, title=None, version="1.0.0"):
    """One :class:`polymorphize_workbook.SheetResult` to one specification."""
    spec = SheetSpec(sheet=result.sheet)
    spec.findings = list(result.findings)

    if result.status == "skipped":
        spec.status = "skipped"
        return spec
    # A missing Request Body label already failed the sheet in the reader. A
    # label that is present but empty is permitted, and reaches here with
    # ``result.request.root`` unset.
    if result.status == "failed" or result.request is None:
        spec.status = "failed"
        return spec

    endpoints = [c.endpoint for c in result.layout.sor_cols]
    variants = pair_variants(result, spec.findings) if len(endpoints) > 1 else []
    if len(endpoints) == 1:
        spec.pattern = "P1"
    elif result.variants and len(result.variants) == len(endpoints):
        spec.pattern = "P2"
    else:
        spec.pattern = "P3"
    spec.variants = variants

    banner = result.banner
    method = derive_method(result.sheet, banner)
    path = derive_path(result.sheet, banner)

    # A message carries a payload only when at least one of its attributes
    # names an SOR field. An empty message is permitted, and publishes nothing
    # rather than an empty object: no requestBody on the request side, and 204
    # No Content on the response side.
    schemas = {}

    def payload(label, section, is_request):
        if section is None or section.root is None:
            return None, {}
        if not _mapped_leaves(section.root, endpoints):
            rows = _count_leaves(section.root)
            spec.findings.append(Finding(
                "E001", "warning", result.sheet,
                cell_ref(section.root.row, result.layout.sor_cols[0].index),
                "the %s section declares %d attribute%s and not one of them names "
                "an SOR field, so no payload was published for it"
                % (label, rows, "" if rows == 1 else "s"),
                "either an SOR field against the attributes this message carries, "
                "or an empty section if it genuinely carries none",
                "fill in the SOR column for the attributes this message carries. "
                "If it carries none, delete the rows and leave the %s label in "
                "place." % label,
            ))
            return None, {}
        return _message(result, section, variants, spec.pattern, schemas,
                        is_request=is_request)

    req_name, req_counts = payload("Request Body", result.request, True)
    res_name, res_counts = payload("Response Body", result.response, False)

    if req_name is None and method in METHODS_WITHOUT_BODY:
        # No body to carry, so the natural verb stands.
        pass
    elif method in METHODS_WITHOUT_BODY:
        spec.findings.append(Finding(
            "G001", "warning", result.sheet, "",
            "the sheet name implies %s but the Request Body section declares a "
            "schema, and %s cannot carry a body, so the operation was emitted as "
            "POST" % (method.upper(), method.upper()),
            "either a method that carries a body, or the request expressed as "
            "query parameters",
            "confirm POST for this operation, or move the request attributes into "
            "a Request Parameter section so they become query parameters",
        ))
        method = "post"

    operation = {
        "operationId": (pascal(result.sheet)[:1].lower() + pascal(result.sheet)[1:]),
        "summary": text(banner.get("use_case", ""))[:120] or result.sheet,
        "tags": [text(banner.get("service_domain", "")) or "Operations"],
        "responses": {},
    }
    if req_name:
        operation["requestBody"] = {
            "required": True,
            "content": {"application/json": {
                "schema": {"$ref": "#/components/schemas/%s" % req_name}}},
        }
    if res_name:
        operation["responses"]["200"] = {
            "description": "Successful response.",
            "content": {"application/json": {
                "schema": {"$ref": "#/components/schemas/%s" % res_name}}},
        }
    else:
        operation["responses"]["204"] = {
            "description": "Successful. This operation returns no content: the "
                           "Response Body section of the mapping workbook "
                           "carries no attribute mapped to an SOR field.",
        }
    if banner.get("use_case"):
        operation["description"] = text(banner["use_case"])
    if banner.get("behaviour_qualifier"):
        operation["x-bian-behaviour-qualifier"] = banner["behaviour_qualifier"]
    if banner.get("bian_endpoint"):
        operation["x-bian-equivalent-endpoint"] = text(banner["bian_endpoint"])
    if banner.get("sor_endpoint"):
        operation["x-sor-endpoints"] = text(banner["sor_endpoint"])

    document = {
        "openapi": OPENAPI_VERSION,
        "info": {
            "title": title or ("%s %s" % (text(banner.get("service_domain", "")) or
                                          "Service", result.sheet)).strip(),
            "version": version,
            "description": text(banner.get("use_case", "")) or
            "Generated from the SOR mapping workbook.",
            "x-generated-by": "SOR Polymorphizer %s" % __version__,
            "x-generated-at": _dt.datetime.now(_dt.timezone.utc)
            .replace(microsecond=0).isoformat(),
            "x-source-sheet": result.sheet,
            "x-polymorphism-pattern": spec.pattern,
        },
        "paths": {path: {method: operation}},
        "components": {"schemas": schemas},
    }
    if banner.get("service_domain"):
        document["info"]["x-bian-service-domain"] = text(banner["service_domain"])

    leftovers = empty_schemas(document)
    if leftovers:
        spec.status = "failed"
        spec.findings.append(Finding(
            "X002", "error", result.sheet, "",
            "%d schema%s reached the output with nothing inside %s: %s"
            % (len(leftovers), "" if len(leftovers) == 1 else "s",
               "it" if len(leftovers) == 1 else "them",
               ", ".join(leftovers[:8])),
            "every schema to declare properties, a composition keyword or a scalar "
            "type",
            "this is a fault in the generator rather than in the workbook. Send this "
            "sheet to the maintainer. No specification was written for it.",
        ))
        return spec

    spec.document = document
    spec.path, spec.method = path, method
    spec.stats = {
        "sor_endpoints": len(endpoints),
        "declared_request_attributes": _count_nodes(result.request.root),
        "declared_response_attributes": _count_nodes(
            result.response.root) if result.response and result.response.root else 0,
        "request_leaves": req_counts,
        "response_leaves": res_counts,
        "ignored_sections": [lab for lab, _r in result.ignored_sections],
        "ignored_rows": result.ignored_rows,
    }
    return spec


def _mapped_leaves(node, endpoints):
    """Leaves beneath ``node`` that at least one SOR endpoint supplies."""
    out = 0
    for c in node.children:
        if c.children:
            out += _mapped_leaves(c, endpoints)
        elif any(c.sor.get(e) for e in endpoints):
            out += 1
    return out


def empty_schemas(document):
    """Schema names that would publish an empty object.

    The original defect this tool exists to prevent is a specification that
    declares a message and then offers nothing inside it. A schema is empty
    when it has neither properties, nor a composition keyword, nor a scalar
    type. Reaching output with one is either an unmapped section in the
    workbook or a fault in this module, and both must stop the endpoint.
    """
    bad = []
    schemas = document.get("components", {}).get("schemas", {})

    def check(name, node):
        if not isinstance(node, dict):
            return
        if any(k in node for k in ("allOf", "oneOf", "anyOf", "$ref", "enum")):
            return
        t = node.get("type")
        if t == "array":
            items = node.get("items")
            if not isinstance(items, dict):
                bad.append(name)
            else:
                check(name + ".items", items)
            return
        if t in (None, "object"):
            props = node.get("properties")
            if not props:
                bad.append(name)
                return
            for k, v in props.items():
                check("%s.%s" % (name, k), v)

    for name, node in schemas.items():
        check(name, node)
    return bad


def _count_nodes(node):
    if node is None:
        return 0
    return 1 + sum(_count_nodes(c) for c in node.children) - 1 if node else 0


def _count_leaves(node):
    if node is None:
        return 0
    if not node.children:
        return 1
    return sum(_count_leaves(c) for c in node.children)


# --------------------------------------------------------------------------- #
# Whole workbook
# --------------------------------------------------------------------------- #


@dataclass
class RunResult:
    workbook: str
    out_dir: str
    specs: list = field(default_factory=list)
    split: bool = False
    merged: object = None
    sor_index: object = None
    verifications: dict = field(default_factory=dict)
    showcase_path: str = ""

    @property
    def sor_verified(self):
        return bool(self.sor_index is not None and not self.sor_index.empty)

    @property
    def ok(self):
        return [s for s in self.specs if s.status == "ok"]

    @property
    def failed(self):
        return [s for s in self.specs if s.status == "failed"]

    @property
    def skipped(self):
        return [s for s in self.specs if s.status == "skipped"]


def run(workbook, out_dir, *, strict=False, version="1.0.0", fmt="yaml",
        write=True, split=False, merged_name="openapi", title=None,
        sor=None, showcase=True, log=print):
    """Every operation sheet of a workbook.

    By default the sheets are merged into one specification, because
    polymorphism only means anything inside a single namespace. ``split=True``
    writes one file per sheet instead, which is what the earlier releases did
    and is useful when isolating a single endpoint while debugging.

    ``sor`` is one or more System of Record specifications, or a folder of
    them. When given, every SOR field name in the workbook is looked up in the
    SOR endpoint that sheet names, and an element whose field does not exist
    there is excluded from the interface. Without it the workbook's word is
    taken, which every report states plainly.
    """
    results = wbk.read_workbook(workbook, strict=strict)
    out = RunResult(workbook=workbook, out_dir=out_dir, split=split)
    if write:
        os.makedirs(out_dir, exist_ok=True)

    # Verification comes before generation: it clears the mappings that do not
    # resolve, so every later stage prunes those elements without needing to
    # know that verification happened at all.
    if sor:
        import polymorphize_sor as sormod
        out.sor_index = sormod.SorIndex.from_paths(sor, log=log)
        if out.sor_index.load_errors:
            for path, why in out.sor_index.load_errors:
                log("  ERROR  could not read %s: %s" % (os.path.basename(path), why))
        if out.sor_index.empty:
            log("  WARNING: no SOR operations were loaded, so nothing could be "
                "verified. The workbook's word was taken instead.")
        else:
            s = out.sor_index.summary()
            log("  SOR    %d file(s), %d operation(s) across %d path(s)"
                % (s["files"], s["operations"], s["paths"]))
            out.verifications = sormod.verify_workbook(results, out.sor_index,
                                                       log=log)
    else:
        log("  NOTE   no SOR specification was supplied, so every SOR field "
            "name in the workbook was taken on trust.")

    kept = []
    for r in results:
        spec = generate_sheet(r, version=version)
        out.specs.append(spec)
        if spec.status == "skipped":
            log("  skip   %-34s no Level columns, treated as a support sheet"
                % r.sheet)
            continue
        if spec.status == "failed":
            log("  FAILED %-34s %s" % (r.sheet,
                                       spec.findings[0].what[:70] if spec.findings else ""))
            if write:
                # A specification from an earlier run must not survive as though
                # it were current. Remove it and leave the explanation instead.
                for candidate in (_target(out_dir, r.sheet, "yaml"),
                                  _target(out_dir, r.sheet, "json")):
                    if os.path.exists(candidate):
                        os.remove(candidate)
                        log("         removed the stale %s from an earlier run"
                            % os.path.basename(candidate))
                _write_failure(out_dir, spec)
            continue
        log("  ok     %-34s %s  %s %s" % (r.sheet, spec.pattern,
                                          spec.method.upper(), spec.path))
        kept.append((r, spec))
        if write and split:
            target = _target(out_dir, r.sheet, fmt)
            core._write(target, dump_document(spec.document, fmt))
        if write:
            stale_note = os.path.join(out_dir,
                                      "%s.FAILED.md" % (slug(r.sheet) or "sheet"))
            if os.path.exists(stale_note):
                os.remove(stale_note)

    if not split and kept:
        import polymorphize_merge as mrg
        out.merged = mrg.merge([r for r, _s in kept], [s for _r, s in kept],
                               title=title, version=version, log=log)
        if out.merged.errors:
            log("")
            log("  THE MERGE FAILED. No single specification was written.")
            for f in out.merged.errors:
                log("    %s  %s" % (f.code, f.what))
        else:
            m = out.merged.stats
            log("")
            log("  merged %d operations into one specification: %d schemas, "
                "%d groups hoisted, %d specialised across operations"
                % (m["operations"], m["schemas"], m["hoisted_groups"],
                   m["split_groups"]))
            if write:
                target = os.path.join(out_dir, "%s.%s"
                                      % (merged_name,
                                         "yaml" if fmt == "yaml" else "json"))
                core._write(target, dump_document(out.merged.document, fmt))
                log("  wrote  %s" % os.path.basename(target))
    elif not split and write:
        # Nothing generated, so any merged document on disk is stale.
        for ext in ("yaml", "json"):
            stale = os.path.join(out_dir, "%s.%s" % (merged_name, ext))
            if os.path.exists(stale):
                os.remove(stale)
                log("  removed the stale %s from an earlier run"
                    % os.path.basename(stale))

    if write and showcase and (out.ok or out.failed):
        import polymorphize_showcase as show
        try:
            out.showcase_path = show.write(
                out,
                os.path.join(out_dir, show.default_filename(out, merged_name)),
                title=title)
            log("  wrote  %s" % os.path.basename(out.showcase_path))
        except Exception as exc:                               # noqa: BLE001
            # The specification is already on disk. A failure to draw the
            # report is a reporting problem, not a transform problem.
            log("  WARNING: could not build the showcase report: %s: %s"
                % (type(exc).__name__, exc))

    if write:
        core._write(os.path.join(out_dir, "generation_report.md"), build_report(out))
    return out


def _target(out_dir, sheet, fmt):
    """The output path for one sheet."""
    return os.path.join(out_dir, "%s.%s" % (slug(sheet) or "operation",
                                            "yaml" if fmt == "yaml" else "json"))


def _write_failure(out_dir, spec):
    """A failed sheet gets an explanation, never a specification."""
    lines = [
        "# FAILED: %s" % spec.sheet,
        "",
        "No specification was generated for this endpoint. The rest of the "
        "workbook was processed normally.",
        "",
    ]
    for f in spec.findings:
        if f.severity != "error":
            continue
        lines += [
            "## %s at %s" % (f.code, f.where),
            "",
            "**Found:** %s" % f.what,
            "",
            "**Expected:** %s" % f.expected,
            "",
            "**Fix:** %s" % f.fix,
            "",
        ]
    core._write(os.path.join(out_dir, "%s.FAILED.md" % (slug(spec.sheet) or "sheet")),
                "\n".join(lines))


def build_report(out):
    """The run report, failures first and impossible to scroll past."""
    L = []
    total = len(out.specs) - len(out.skipped)
    L.append("# SOR Polymorphizer generation report")
    L.append("")
    L.append("Workbook: `%s`" % os.path.basename(out.workbook))
    L.append("")
    L.append("Generated %d of %d endpoints. **%d failed.**"
             % (len(out.ok), total, len(out.failed)))
    L.append("")
    if out.split:
        L.append("Output: **one specification per sheet**, for debugging. "
                 "Polymorphism across operations is not applied in this mode, "
                 "because each sheet has its own schema namespace.")
    else:
        L.append("Output: **one specification** for the whole workbook.")
    L.append("")

    if out.merged is not None and out.merged.errors:
        L.append("## The merge failed")
        L.append("")
        L.append("The individual endpoints were generated, but they could not "
                 "be combined into one specification, so none was written.")
        L.append("")
        for f in out.merged.errors:
            L.append("- **%s**" % f.code)
            L.append("  - Found: %s" % f.what)
            L.append("  - Expected: %s" % f.expected)
            L.append("  - Fix: %s" % f.fix)
        L.append("")

    if out.failed:
        L.append("## Failed endpoints")
        L.append("")
        L.append("No specification was written for these sheets. Each one needs a "
                 "correction in the workbook.")
        L.append("")
        for s in out.failed:
            L.append("### %s" % s.sheet)
            L.append("")
            for f in s.findings:
                if f.severity != "error":
                    continue
                L.append("- **%s at %s**" % (f.code, f.where))
                L.append("  - Found: %s" % f.what)
                L.append("  - Expected: %s" % f.expected)
                L.append("  - Fix: %s" % f.fix)
            L.append("")

    L.append("## Endpoints generated")
    L.append("")
    L.append("| Sheet | Pattern | SOR endpoints | Method | Path |")
    L.append("|---|---|---|---|---|")
    for s in out.ok:
        L.append("| %s | %s | %d | %s | `%s` |"
                 % (s.sheet, s.pattern, s.stats.get("sor_endpoints", 0),
                    s.method.upper(), s.path))
    L.append("")

    L.append("### Patterns")
    L.append("")
    L.append("- **P1**, one SOR endpoint: a single schema per message, pruned to "
             "what SOR supplies.")
    L.append("- **P2**, several SOR endpoints with a `%s` enumeration: base plus "
             "`allOf` variants, wrapped in `oneOf` with a discriminator on the "
             "request." % VARIANT_PROPERTY)
    L.append("- **P3**, several SOR endpoints with no tag: `oneOf` without a "
             "discriminator.")
    L.append("")

    if out.merged is not None and out.merged.document is not None:
        m = out.merged
        L.append("## The merged specification")
        L.append("")
        L.append("| | |")
        L.append("|---|---|")
        L.append("| Operations | %d |" % m.stats["operations"])
        L.append("| Schemas | %d |" % m.stats["schemas"])
        L.append("| Groups hoisted into components | %d |"
                 % m.stats["hoisted_groups"])
        L.append("| Of those, specialised across operations | %d |"
                 % m.stats["split_groups"])
        L.append("")
        split = [(name, base, derived) for name, base, derived in m.profiles
                 if len(derived) > 1]
        if split:
            L.append("### Groups specialised across operations")
            L.append("")
            L.append("These groups are used by more than one operation and do "
                     "not agree on what they publish. Each keeps its own name "
                     "for the shared core, and every operation that uses it "
                     "gets a derived schema inheriting through `allOf`. A group "
                     "with no shared base has nothing in common between its "
                     "uses; the P001 and P002 warnings say whether that is "
                     "intended or a workbook inconsistency.")
            L.append("")
            L.append("| Group | Shared base | Derived schemas |")
            L.append("|---|---|---|")
            for name, base, derived in sorted(split):
                L.append("| `%s` | %s | %d |"
                         % (name, ("`%s`" % base) if base else "none", len(derived)))
            L.append("")
        if m.findings:
            L.append("### Findings from the merge")
            L.append("")
            L.append("| Code | Found | Fix |")
            L.append("|---|---|---|")
            for f in m.findings:
                L.append("| %s | %s | %s |"
                         % (f.code, f.what.replace("|", "/").replace("\n", " "),
                            f.fix.replace("|", "/").replace("\n", " ")))
            L.append("")

    warn = [(s, f) for s in out.specs for f in s.findings if f.severity == "warning"]
    if warn:
        L.append("## Warnings")
        L.append("")
        L.append("These did not stop generation. Each one is worth an analyst's "
                 "attention.")
        L.append("")
        L.append("| Code | Where | Found | Fix |")
        L.append("|---|---|---|---|")
        for s, f in warn:
            L.append("| %s | `%s` | %s | %s |"
                     % (f.code, f.where, f.what.replace("|", "/"),
                        f.fix.replace("|", "/")))
        L.append("")

    ignored = [(s.sheet, s.stats.get("ignored_sections", []), s.stats.get("ignored_rows", 0))
               for s in out.ok if s.stats.get("ignored_sections") or s.stats.get("ignored_rows")]
    if ignored:
        L.append("## Content ignored by design")
        L.append("")
        L.append("Only Request Body and Response Body are read. Header and "
                 "parameter sections are skipped, and anything below the mapping "
                 "grid, such as a pasted sample payload, is not read at all.")
        L.append("")
        L.append("| Sheet | Sections skipped | Rows below the grid |")
        L.append("|---|---|---|")
        for sheet, secs, rows in ignored:
            L.append("| %s | %s | %d |" % (sheet, ", ".join(secs) or "none", rows))
        L.append("")

    if out.skipped:
        L.append("## Support sheets")
        L.append("")
        L.append("No Level columns, so these are not operation sheets: %s."
                 % ", ".join("`%s`" % s.sheet for s in out.skipped))
        L.append("")
    return "\n".join(L)


__all__ = ["RunResult", "dump_document", "SheetSpec", "Variant", "build_report", "derive_method",
           "derive_path", "emit", "generate_sheet", "pair_variants", "pascal",
           "run", "slug", "supported"]
