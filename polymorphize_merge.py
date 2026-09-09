"""
polymorphize_merge — one specification for the whole workbook.

Version 6.4.

Why one document
================

Polymorphism only means anything inside a single namespace. Twelve separate
specifications each write out their own copy of ``IssuedDeviceIdentifier`` and
never have to agree; one specification forces them to agree, and that is where
the reduction comes from.

Two axes of variance
====================

Merging the sheets makes a second axis of variance visible, and the two are
resolved by different mechanisms because they are settled at different times.

Which operation
    Design time. The client picks a path, and the path settles the shape before
    a byte is sent. Resolved by ``allOf`` specialisation of the shared group,
    with **no discriminator**: there is no runtime ambiguity to resolve, and a
    discriminator would demand a tag property no SOR supplies.

Which variant
    Run time. The client sets ``requestVariant`` and that chooses which SOR
    endpoint serves the call. Resolved by ``oneOf`` with
    ``discriminator: requestVariant``, exactly as before the merge.

The two compose, because a shared group can sit inside a variant of a
variant-driven operation.

How hoisting makes both axes one calculation
============================================

Every group with children is hoisted out of its parent and published as a
named component, so a property's value in any schema is either a scalar
fragment or a ``$ref``. Both axes then reduce to the same operation on
property maps:

    the base declares a property when its emitted value is identical
    everywhere; everything else goes in the delta of each derived schema

which is sound because ``allOf`` conjoins rather than overrides. A group whose
content varies cannot appear in the base, because its ``$ref`` differs; a
group whose content agrees appears once, because its ``$ref`` is the same.

Naming
======

A group with one shape keeps its own name. Where a group splits, the shape
seen in the most places **keeps the plain name**, the intersection base takes
``<Name>Base``, and the remaining shapes become ``<Name>For<Operation>``, or
``<Name>Profile<n>`` where several operations share a shape. Every derived
schema carries ``x-used-by`` naming the operations it serves.

Changed in 6.4. Up to 6.3 the plain name was reserved for the intersection
base, which no property ever references: on the reference workbook not one of
the ten bases was pointed at directly, so every consumer-visible reference
carried a suffix. Because a generator names its classes after component
schemas, that suffix reached the consumer even though the wire did not change,
a property key coming from the workbook's Level column rather than from the
component name. Giving the plain name to the shape most references point at
took suffixed references on the reference workbook from 36 of 68 to 23.
"""

from __future__ import annotations

import datetime as _dt
import json
import re
from collections import OrderedDict, defaultdict
from dataclasses import dataclass, field

import polymorphize_generate as gen
import polymorphize_workbook as wbk
from polymorphize_workbook import Finding, VARIANT_PROPERTY, cell_ref, key, text

__version__ = "6.4"

OPENAPI_VERSION = gen.OPENAPI_VERSION


# --------------------------------------------------------------------------- #
# Contexts
# --------------------------------------------------------------------------- #


@dataclass
class Context:
    """One (operation, message, variant) triple, and the SOR endpoints it uses."""

    sheet: str
    operation_id: str
    message: str                 # "Request" | "Response"
    endpoints: list
    variant: object = None       # gen.Variant or None
    root: object = None          # wbk.Node

    @property
    def is_request(self):
        return self.message == "Request"

    @property
    def protect(self):
        return (VARIANT_PROPERTY,) if self.is_request else ()

    @property
    def variant_codes(self):
        if self.variant is None:
            return None
        return [self.variant.code] if self.is_request else None

    @property
    def label(self):
        if self.variant is None:
            return gen.pascal(self.operation_id)
        return "%s%s" % (gen.pascal(self.operation_id), self.variant.code)


# --------------------------------------------------------------------------- #
# Shapes
# --------------------------------------------------------------------------- #


def scalar_fragment(node, endpoints, variant_codes):
    """The emitted fragment for a leaf, documentation included."""
    return gen._scalar(node, endpoints, variant_codes=variant_codes)


#: Keys of a leaf fragment that form part of the wire contract. A shape is
#: compared on these alone.
STRUCTURAL_KEYS = ("type", "format", "maxLength", "enum", "x-conditional")


def structural_fragment(node, variant_codes):
    """The part of a leaf that a consumer is actually bound by.

    Descriptions, examples and SOR field names are deliberately excluded.
    Two occurrences of one attribute that agree on type and obligation are the
    same interface even when the prose differs or the SOR sources them from
    different fields, and fragmenting the namespace on documentation would
    defeat the reuse the merge exists to achieve. The differences are folded
    back in at emission, and an SOR field that varies by operation is recorded
    per operation rather than dropped.
    """
    full = gen._scalar(node, (), variant_codes=variant_codes)
    return {k: full[k] for k in STRUCTURAL_KEYS if k in full}


def array_item(node):
    """The item of a repeating group, as ``(component base name, node)``.

    An array is never published as a component itself, only as
    ``{type: array, items: {$ref: ...}}`` pointing at its item. That keeps every
    hoisted component an object, so ``allOf`` specialisation applies uniformly:
    there is no way to specialise an array through ``allOf``.

    An analyst who writes a repeating group as an ``Array`` row with one
    ``object`` row beneath it means that row to be the item, so it takes its own
    name. An array with several rows beneath it has an anonymous item, which is
    named after the array with ``Item`` appended.

    The choice is made on the declared children rather than the surviving ones,
    so it cannot change between variants of the same array.
    """
    kids = node.children
    if len(kids) == 1 and kids[0].children and kids[0].json_type == "object":
        return kids[0].name, kids[0]
    return node.name + "Item", node


def datasig(node, endpoints, protect, variant_codes):
    """A canonical signature of the object ``node`` publishes, from data alone.

    Component names do not appear, so the signature can be computed before
    anything has been named. Two occurrences with equal signatures are
    interchangeable and share one component.

    An array node passed here is being treated as its own anonymous item, which
    is what :func:`array_item` decides.
    """
    parts = [_child_sig(c, endpoints, protect, variant_codes)
             for c in gen.retained(node, endpoints, protect=protect)]
    return json.dumps(["object", sorted(parts),
                       sorted(_required_of(node, endpoints, protect))])


def _child_sig(child, endpoints, protect, variant_codes):
    if not child.children:
        return [child.name, "S",
                json.dumps(structural_fragment(child, variant_codes),
                           sort_keys=True)]
    if child.json_type == "array":
        base_name, item = array_item(child)
        return [child.name, "A", base_name,
                datasig(item, endpoints, protect, variant_codes)]
    return [child.name, "G", datasig(child, endpoints, protect, variant_codes)]


def _required_of(node, endpoints, protect):
    out = []
    for c in gen.retained(node, endpoints, protect=protect):
        if c.usage == wbk.USAGE_REQUIRED or c.name in protect:
            out.append(c.name)
    return out


# --------------------------------------------------------------------------- #
# The registry
# --------------------------------------------------------------------------- #


class Registry:
    """Every hoisted group, keyed by name and shape, with its assigned names."""

    def __init__(self):
        # name -> sig -> list of (Context, node)
        self.shapes = defaultdict(lambda: defaultdict(list))
        self.assigned = {}          # (name, sig) -> component name
        self.base_names = {}        # name -> base component name, when split
        self.reserved = set()

    # -- collection -------------------------------------------------------

    def collect(self, ctx):
        """Walk one context's tree and record every group it contains."""
        if ctx.root is None:
            return

        def walk(node):
            for child in gen.retained(node, ctx.endpoints, protect=ctx.protect):
                if not child.children:
                    continue
                base_name, target = (array_item(child)
                                     if child.json_type == "array"
                                     else (child.name, child))
                sig = datasig(target, ctx.endpoints, ctx.protect,
                              ctx.variant_codes)
                self.shapes[base_name][sig].append((ctx, target))
                walk(target)
        walk(ctx.root)

    def reserve(self, name):
        self.reserved.add(name)

    # -- naming -----------------------------------------------------------

    def assign_names(self, findings):
        """Give every (name, shape) pair a component name.

        Shapes are ordered by how many contexts use them, then by signature,
        so numbering is stable between runs.

        Where a group splits, the **most-used shape keeps the plain name** and
        the shared core takes ``<Name>Base``. Releases up to 6.3 did the
        opposite, reserving the plain name for the intersection base. That base
        is referenced only through ``allOf`` and never by a property, so on the
        credit card workbook not one of the ten bases was pointed at directly
        and every consumer-visible reference carried a suffix. The name a
        consumer meets in generated client code and in rendered documentation
        should be the plain one wherever it can be.

        This changes published schema names, and therefore the class names a
        generator produces. It does not change the wire: a property key comes
        from the workbook's Level column, never from the component name.
        """
        for name in sorted(self.shapes):
            sigs = self.shapes[name]
            ordered = sorted(sigs, key=lambda s: (-len(sigs[s]), s))
            if len(ordered) == 1:
                self.assigned[(name, ordered[0])] = self._unique(name)
                continue
            # Claim the plain name for the commonest shape first, so the base
            # cannot take it, then name the remaining shapes.
            self.assigned[(name, ordered[0])] = self._unique(name)
            self.base_names[name] = self._unique("%sBase" % name)
            profile = 0
            for sig in ordered[1:]:
                operations = OrderedDict(
                    (c.operation_id, None) for c, _n in sigs[sig])
                if len(operations) == 1:
                    only = next(iter(operations))
                    candidate = "%sFor%s" % (name, gen.pascal(only))
                else:
                    profile += 1
                    candidate = "%sProfile%d" % (name, profile)
                self.assigned[(name, sig)] = self._unique(candidate)

    def _unique(self, candidate):
        if candidate not in self.reserved:
            self.reserved.add(candidate)
            return candidate
        n = 2
        while "%s%d" % (candidate, n) in self.reserved:
            n += 1
        out = "%s%d" % (candidate, n)
        self.reserved.add(out)
        return out

    def name_for(self, base_name, node, ctx):
        sig = datasig(node, ctx.endpoints, ctx.protect, ctx.variant_codes)
        return self.assigned[(base_name, sig)]

    def name_for_child(self, child, ctx):
        """The component a container property points at."""
        if child.json_type == "array":
            base_name, item = array_item(child)
            return self.name_for(base_name, item, ctx)
        return self.name_for(child.name, child, ctx)

    def contexts_for(self, name, sig):
        return self.shapes[name][sig]

    @property
    def split_names(self):
        return sorted(self.base_names)


# --------------------------------------------------------------------------- #
# Emission with refs
# --------------------------------------------------------------------------- #


def emit_body(node, ctx, registry):
    """The body of one hoisted component. Always an object.

    An array node reaching here is standing in for its own anonymous item, so
    it too emits an object: the array wrapper lives on the property that points
    at this component, not on the component.
    """
    props, required = _property_map(node, ctx, registry)
    out = {"type": "object", "properties": props}
    if required:
        out["required"] = required
    if node.description and node.json_type != "array":
        out["description"] = node.description
    return out


def _property_map(node, ctx, registry):
    """``{property: fragment, $ref or array of $ref}`` plus the required list."""
    props, required = OrderedDict(), []
    for child in gen.retained(node, ctx.endpoints, protect=ctx.protect):
        if not child.children:
            props[child.name] = scalar_fragment(child, ctx.endpoints,
                                                ctx.variant_codes)
        elif child.json_type == "array":
            entry = {"type": "array",
                     "items": {"$ref": _ref(
                         registry.name_for_child(child, ctx))}}
            if child.description:
                entry["description"] = child.description
            mapped = {e: child.sor[e] for e in ctx.endpoints
                      if child.sor.get(e)}
            if mapped:
                entry["x-sor-field" if len(mapped) == 1 else "x-sor-fields"] = (
                    next(iter(mapped.values())) if len(mapped) == 1 else mapped)
            props[child.name] = entry
        else:
            props[child.name] = {"$ref": _ref(
                registry.name_for_child(child, ctx))}
        if child.usage == wbk.USAGE_REQUIRED or child.name in ctx.protect:
            required.append(child.name)
    return props, required


def _ref(name):
    return "#/components/schemas/%s" % name


# --------------------------------------------------------------------------- #
# Folding documentation back in
# --------------------------------------------------------------------------- #


def _leaf_index(node, ctx, prefix=()):
    """``{path: node}`` for every retained leaf, stopping at hoisted groups."""
    out = {}
    for child in gen.retained(node, ctx.endpoints, protect=ctx.protect):
        path = prefix + (child.name,)
        if child.children:
            # A group of its own, published as a separate component and
            # documented when that component is emitted.
            continue
        out[path] = child
    return out


def _best(values):
    """The most-used value, longest first on a tie. Empty values ignored."""
    counts = defaultdict(int)
    for v in values:
        if v:
            counts[v] += 1
    if not counts:
        return None
    return max(sorted(counts), key=lambda v: (counts[v], len(v)))


def enrich(body, uses, registry):
    """Fold documentation and SOR mappings from every use into one component.

    The shape was compared structurally, so the uses agree on what the
    component publishes and differ only in prose and in which SOR field backs
    each attribute. Where an SOR field differs between operations it is
    recorded against each operation by name, because the traceability is the
    reason the annotation exists.
    """
    indexes = [(ctx, _leaf_index(node, ctx)) for ctx, node in uses]
    best_desc = _best([node.description for _c, node in uses
                       if node.json_type != "array"])
    target = body
    if best_desc:
        target["description"] = best_desc
    elif "description" in target:
        del target["description"]

    props = target.get("properties")
    if not isinstance(props, dict):
        return body
    for name, fragment in props.items():
        if not isinstance(fragment, dict) or "$ref" in fragment:
            continue
        path = (name,)
        nodes = [(ctx, index[path]) for ctx, index in indexes if path in index]
        if not nodes:
            continue
        desc = _best([n.description for _c, n in nodes])
        if desc:
            fragment["description"] = desc
        else:
            fragment.pop("description", None)
        example = _best([n.example for _c, n in nodes])
        if example:
            fragment["example"] = example
        else:
            fragment.pop("example", None)

        fragment.pop("x-sor-field", None)
        fragment.pop("x-sor-fields", None)
        by_operation = OrderedDict()
        for ctx, n in nodes:
            mapped = {e: n.sor[e] for e in ctx.endpoints if n.sor.get(e)}
            if not mapped:
                continue
            label = ctx.operation_id if ctx.variant is None else \
                "%s[%s]" % (ctx.operation_id, ctx.variant.code)
            by_operation[label] = (next(iter(mapped.values()))
                                   if len(mapped) == 1 else mapped)
        distinct = {json.dumps(v, sort_keys=True) for v in by_operation.values()}
        if len(distinct) == 1:
            fragment["x-sor-field"] = next(iter(by_operation.values()))
        elif by_operation:
            fragment["x-sor-field-by-operation"] = dict(by_operation)
    return body


# --------------------------------------------------------------------------- #
# The base and delta calculation, used for both axes
# --------------------------------------------------------------------------- #


DOC_KEYS = ("description", "example", "x-sor-field", "x-sor-fields",
            "x-sor-field-by-operation")


def contract_of(value):
    """A property value reduced to what a consumer is bound by.

    Two shapes of one group may document the same attribute differently and
    source it from different SOR fields while publishing an identical
    contract. The base/delta split compares this reduction, so documentation
    never pushes an attribute out of the shared base.
    """
    if not isinstance(value, dict):
        return value
    if "$ref" in value:
        return {"$ref": value["$ref"]}
    out = {k: v for k, v in value.items() if k not in DOC_KEYS}
    if out.get("type") == "array" and isinstance(out.get("items"), dict):
        out["items"] = contract_of(out["items"])
    return out


def _merge_docs(values):
    """One property value carrying the documentation of all its occurrences."""
    base = dict(values[0])
    for k in ("description", "example"):
        best = _best([v.get(k) for v in values if isinstance(v, dict)])
        if best:
            base[k] = best
        else:
            base.pop(k, None)
    by_operation = OrderedDict()
    single = []
    for v in values:
        if not isinstance(v, dict):
            continue
        if "x-sor-field-by-operation" in v:
            by_operation.update(v["x-sor-field-by-operation"])
        elif "x-sor-field" in v:
            single.append(v["x-sor-field"])
        elif "x-sor-fields" in v:
            single.append(v["x-sor-fields"])
    for k in DOC_KEYS[2:]:
        base.pop(k, None)
    distinct = {json.dumps(v, sort_keys=True) for v in single}
    if by_operation:
        for s in single:
            by_operation.setdefault("(unattributed)", s)
        base["x-sor-field-by-operation"] = dict(by_operation)
    elif len(distinct) == 1:
        base["x-sor-field"] = single[0]
    elif distinct:
        base["x-sor-fields"] = sorted(
            {v if isinstance(v, str) else json.dumps(v) for v in single})
    return base


def split_shapes(bodies, base_description):
    """Given several bodies of one group, return ``(base, {label: delta})``.

    A property lands in the base when its *contract* is identical in every
    body, whatever the prose says. Everything else goes in the delta of the
    body it came from. This is sound because ``allOf`` conjoins: a property in
    the base applies to every derived schema, so one that varies must not be
    there.

    Returns ``(None, bodies)`` when nothing is common, because an empty base
    would be an empty object and there is nothing to inherit.
    """
    labels = list(bodies)
    first = bodies[labels[0]]
    common = OrderedDict()
    for name, value in (first.get("properties") or {}).items():
        contract = contract_of(value)
        others = []
        ok = True
        for l in labels[1:]:
            props = bodies[l].get("properties") or {}
            if name not in props or contract_of(props[name]) != contract:
                ok = False
                break
            others.append(props[name])
        if ok:
            common[name] = _merge_docs([value] + others) \
                if isinstance(value, dict) else value

    if not common:
        return None, {l: _as_object(bodies[l]) for l in labels}

    common_required = [
        r for r in (first.get("required") or [])
        if r in common and all(r in (bodies[l].get("required") or [])
                               for l in labels[1:])]
    base = {"type": "object", "properties": common}
    if common_required:
        base["required"] = common_required
    if base_description:
        base["description"] = base_description

    deltas = OrderedDict()
    for label in labels:
        body = bodies[label]
        props = OrderedDict(
            (n, v) for n, v in (body.get("properties") or {}).items()
            if n not in common)
        required = [r for r in (body.get("required") or [])
                    if r not in common_required]
        delta = {"type": "object"}
        if props:
            delta["properties"] = props
        if required:
            delta["required"] = required
        deltas[label] = delta
    return base, deltas


def _as_object(body):
    out = dict(body)
    out.setdefault("type", "object")
    return out


# --------------------------------------------------------------------------- #
# The merged document
# --------------------------------------------------------------------------- #


@dataclass
class MergedSpec:
    document: dict = None
    findings: list = field(default_factory=list)
    operations: list = field(default_factory=list)   # (sheet, method, path)
    stats: dict = field(default_factory=dict)
    profiles: list = field(default_factory=list)     # (group, base, [derived])

    @property
    def errors(self):
        return [f for f in self.findings if f.severity == "error"]


def _contexts_for(result, spec):
    """The contexts one sheet contributes: one per message, per variant."""
    out = []
    op_id = spec.document["paths"][spec.path][spec.method]["operationId"] \
        if spec.document else gen.pascal(result.sheet)
    endpoints = [c.endpoint for c in result.layout.sor_cols]
    variants = spec.variants or []
    for message, section in (("Request", result.request),
                             ("Response", result.response)):
        if section is None or section.root is None:
            continue
        if not gen._mapped_leaves(section.root, endpoints):
            continue
        if spec.pattern == "P1" or not variants:
            out.append(Context(result.sheet, op_id, message, endpoints,
                               None, section.root))
        else:
            for v in variants:
                out.append(Context(result.sheet, op_id, message, [v.endpoint],
                                   v, section.root))
    return out


def _inconsistency_findings(registry, findings):
    """Why a group fragmented, when the reason looks like a mistake.

    A group splitting because different operations genuinely carry different
    attributes is the point of the exercise and is reported in the run report,
    not as a finding. A group splitting because the *same* attribute is
    declared with two different types, lengths or obligations is a workbook
    inconsistency, and fixing it collapses the family into fewer schemas.
    """
    for name in sorted(registry.base_names):
        sigs = registry.shapes[name]
        declared = defaultdict(set)          # attribute -> {canonical fragment}
        where = defaultdict(set)             # attribute -> {operation}
        for _sig, uses in sigs.items():
            for ctx, node in uses:
                for child in gen.retained(node, ctx.endpoints,
                                          protect=ctx.protect):
                    if child.children:
                        continue
                    frag = json.dumps(
                        dict(structural_fragment(child, ctx.variant_codes),
                             required=(child.usage == wbk.USAGE_REQUIRED)),
                        sort_keys=True)
                    declared[child.name].add(frag)
                    where[child.name].add(ctx.operation_id)
        for attribute in sorted(declared):
            variants = sorted(declared[attribute])
            if len(variants) < 2:
                continue
            findings.append(Finding(
                "P002", "warning", "(merged)", "",
                "%s.%s is declared %d different ways across %s: %s"
                % (name, attribute, len(variants),
                   ", ".join(sorted(where[attribute])),
                   " against ".join(_readable(v) for v in variants)),
                "one declaration of one attribute wherever it appears",
                "make the Data Type and Usage of %s.%s the same on every sheet. "
                "Doing so lets the tool publish one shared schema instead of %d."
                % (name, attribute, len(sigs)),
            ))


def _readable(fragment_json):
    """A structural fragment as something an analyst can compare by eye."""
    d = json.loads(fragment_json)
    bits = [str(d.get("type", "?"))]
    if d.get("maxLength"):
        bits.append("(%d)" % d["maxLength"])
    if d.get("format"):
        bits.append("format %s" % d["format"])
    bits.append("mandatory" if d.get("required") else "optional")
    if d.get("x-conditional"):
        bits.append("conditional")
    return " ".join(bits)


def _case_variant_findings(registry, roots, findings):
    """Names that differ only by case are a hazard in one namespace.

    Two components a keystroke apart invite a developer to reference the wrong
    one. They are reported rather than normalised, because normalising would
    silently rename fields in the published interface.
    """
    seen = defaultdict(set)
    for name in list(registry.shapes) + list(roots):
        seen[name.lower()].add(name)
    for _low, spellings in sorted(seen.items()):
        if len(spellings) > 1:
            names = sorted(spellings)
            findings.append(Finding(
                "N001", "warning", "(merged)", "",
                "%s differ only by case, and in one specification they become "
                "separate schemas one keystroke apart"
                % " and ".join(repr(n) for n in names),
                "one spelling for one concept across the whole workbook",
                "settle on one spelling, %r, and correct the others in the "
                "workbook. The tool does not rename them for you, because that "
                "would change the published field names without your seeing it."
                % names[0],
            ))


def merge(results, specs, *, title=None, version="1.0.0", description=None,
          log=print):
    """Build one document from every sheet that generated."""
    out = MergedSpec()
    ok = [(r, s) for r, s in zip(results, specs) if s.status == "ok"]
    if not ok:
        out.findings.append(Finding(
            "Z001", "error", "(merged)", "",
            "no sheet produced a specification, so there is nothing to merge",
            "at least one operation sheet that generates",
            "correct the failures reported for the individual sheets first",
        ))
        return out

    # -- contexts and the group registry ---------------------------------
    registry = Registry()
    contexts = []
    roots = OrderedDict()          # message root name -> [(result, spec, ctxs)]
    for result, spec in ok:
        sheet_ctxs = _contexts_for(result, spec)
        contexts.extend(sheet_ctxs)
        for message, section in (("Request", result.request),
                                 ("Response", result.response)):
            if section is None or section.root is None:
                continue
            endpoints = [c.endpoint for c in result.layout.sor_cols]
            if not gen._mapped_leaves(section.root, endpoints):
                continue
            roots.setdefault(section.root.name, []).append(
                (result, spec, message))

    # Message root names are reserved first: a hoisted group must never take a
    # name a message already owns.
    for name in roots:
        registry.reserve(name)
    for name, uses in roots.items():
        if len(uses) > 1:
            sheets = ", ".join(sorted({r.sheet for r, _s, _m in uses}))
            out.findings.append(Finding(
                "Z002", "error", "(merged)", "",
                "the schema name %r is declared at Level 1 on more than one "
                "sheet: %s" % (name, sheets),
                "one message schema name per specification",
                "rename the Level 1 row on all but one of those sheets so each "
                "message has its own name",
            ))

    for ctx in contexts:
        registry.collect(ctx)
    registry.assign_names(out.findings)
    _case_variant_findings(registry, roots, out.findings)
    _inconsistency_findings(registry, out.findings)

    schemas = OrderedDict()

    # -- axis one: hoisted groups ----------------------------------------
    for name in sorted(registry.shapes):
        sigs = registry.shapes[name]
        bodies, labels_to_sig = OrderedDict(), {}
        for sig, uses in sigs.items():
            ctx, node = uses[0]
            component = registry.assigned[(name, sig)]
            bodies[component] = enrich(emit_body(node, ctx, registry), uses,
                                       registry)
            labels_to_sig[component] = sig

        if len(bodies) == 1:
            component, body = next(iter(bodies.items()))
            schemas[component] = _annotate(body, sigs[labels_to_sig[component]])
            continue

        base_name = registry.base_names[name]
        base, deltas = split_shapes(
            bodies,
            "Attributes of %s common to every operation that uses it." % name)
        derived_names = []
        if base is None:
            # Nothing in common. Publishing an empty base would be an empty
            # object, so each shape stands alone.
            for component, body in deltas.items():
                schemas[component] = _annotate(body,
                                               sigs[labels_to_sig[component]])
                derived_names.append(component)
            # Why there is no base matters. Shapes that share no attribute name
            # are genuinely different groups that happen to share a label, and
            # nothing is wrong. Shapes that share a name but declare it
            # differently are a workbook inconsistency, and fixing it produces
            # a base.
            shared_names = set.intersection(*[
                set(b.get("properties") or {}) for b in bodies.values()])
            if shared_names:
                out.findings.append(Finding(
                    "P001", "warning", "(merged)", "",
                    "%r is used in %d shapes which all declare %s, but not "
                    "identically, so no shared base could be published"
                    % (name, len(bodies),
                       " and ".join(sorted(shared_names)[:4])),
                    "one declaration of a shared attribute wherever it appears",
                    "align the Data Type and Usage of %s across every sheet. "
                    "The P002 findings name the exact differences. Doing so "
                    "replaces %d schemas with a base and %d small derived "
                    "schemas."
                    % (" and ".join(sorted(shared_names)[:4]),
                       len(bodies), len(bodies)),
                ))
        else:
            schemas[base_name] = base
            for component, delta in deltas.items():
                schemas[component] = _annotate(
                    {"allOf": [{"$ref": _ref(base_name)}, delta]},
                    sigs[labels_to_sig[component]],
                    base=base_name)
                derived_names.append(component)
        out.profiles.append((name, base_name if base is not None else None,
                             derived_names))

    # -- axis two: message schemas ---------------------------------------
    for result, spec in ok:
        endpoints = [c.endpoint for c in result.layout.sor_cols]
        variants = spec.variants or []
        for message, section in (("Request", result.request),
                                 ("Response", result.response)):
            if section is None or section.root is None:
                continue
            if not gen._mapped_leaves(section.root, endpoints):
                continue
            root = section.root
            is_request = message == "Request"
            protect = (VARIANT_PROPERTY,) if is_request else ()

            if spec.pattern == "P1" or not variants:
                ctx = Context(result.sheet, _op_id(spec), message, endpoints,
                              None, root)
                schemas[root.name] = emit_body(root, ctx, registry)
                continue

            bodies = OrderedDict()
            for v in variants:
                ctx = Context(result.sheet, _op_id(spec), message, [v.endpoint],
                              v, root)
                component = "%sFor%s" % (root.name, v.name_fragment)
                component = registry._unique(component)
                bodies[component] = emit_body(root, ctx, registry)
                if is_request:
                    gen._narrow_variant_description(bodies[component], v)

            base_name = registry._unique("%sBase" % root.name)
            base, deltas = split_shapes(
                bodies,
                (root.description + " " if root.description else "") +
                "Attributes common to every variant of %s." % root.name)

            refs, mapping = [], OrderedDict()
            for (component, delta), v in zip(deltas.items(), variants):
                if base is None:
                    entry = _as_object(delta)
                else:
                    entry = {"allOf": [{"$ref": _ref(base_name)}, delta]}
                if is_request:
                    _require_tag(entry)
                entry["description"] = "%s as served by %s." % (
                    root.name, v.endpoint or v.label)
                entry["x-sor-endpoint"] = v.endpoint
                entry["x-variant"] = v.code
                entry["x-variant-name"] = v.label
                schemas[component] = entry
                refs.append({"$ref": _ref(component)})
                mapping[v.code] = _ref(component)
            if base is not None:
                schemas[base_name] = base

            wrapper = {"oneOf": refs}
            if is_request and spec.pattern == "P2":
                wrapper["discriminator"] = {"propertyName": VARIANT_PROPERTY,
                                            "mapping": dict(mapping)}
            elif not is_request and spec.pattern == "P2":
                wrapper["x-selected-by"] = VARIANT_PROPERTY
                wrapper["description"] = (
                    "The subtype returned follows the %s sent on the request. "
                    "OpenAPI cannot express a response discriminator that "
                    "depends on a request property, so the dependency is "
                    "recorded here." % VARIANT_PROPERTY)
            else:
                wrapper["description"] = (
                    "Exactly one subtype applies. The SOR endpoint is selected "
                    "by which identifier the caller supplies, so there is no "
                    "property to discriminate on.")
            schemas[root.name] = wrapper

    # -- paths ------------------------------------------------------------
    domains = _service_domains(ok, out.findings)
    paths, tags, seen = OrderedDict(), OrderedDict(), {}
    for result, spec in ok:
        operation = dict(spec.document["paths"][spec.path][spec.method])
        raw_domain = text(result.banner.get("service_domain", ""))
        if raw_domain:
            operation["tags"] = [_canonical_domain(raw_domain, domains)]
        pm = (spec.path, spec.method)
        if pm in seen:
            out.findings.append(Finding(
                "Z003", "error", "(merged)", "",
                "%s %s is claimed by both %r and %r"
                % (spec.method.upper(), spec.path, seen[pm], result.sheet),
                "one path and method per operation",
                "give one of the two sheets a different Proposed Business API "
                "Endpoint, so the two operations do not collide",
            ))
            continue
        seen[pm] = result.sheet
        paths.setdefault(spec.path, OrderedDict())[spec.method] = operation
        for t in operation.get("tags", []):
            tags.setdefault(t, {"name": t})
        out.operations.append((result.sheet, spec.method, spec.path))

    document = OrderedDict([
        ("openapi", OPENAPI_VERSION),
        ("info", OrderedDict([
            ("title", title or (" and ".join(domains) + " API" if domains
                                else "SOR-aligned API")),
            ("version", version),
            ("description", description or _default_description(ok, domains)),
            ("x-generated-by", "SOR Polymorphizer %s" % __version__),
            ("x-generated-at", _dt.datetime.now(_dt.timezone.utc)
             .replace(microsecond=0).isoformat()),
            ("x-source-sheets", [r.sheet for r, _s in ok]),
        ])),
        ("tags", list(tags.values())),
        ("paths", paths),
        ("components", {"schemas": schemas}),
    ])
    if domains:
        document["info"]["x-bian-service-domains"] = domains

    leftovers = gen.empty_schemas(document)
    if leftovers:
        out.findings.append(Finding(
            "X003", "error", "(merged)", "",
            "%d schema%s reached the merged document with nothing inside: %s"
            % (len(leftovers), "" if len(leftovers) == 1 else "s",
               ", ".join(leftovers[:8])),
            "every schema to declare properties, a composition keyword or a "
            "scalar type",
            "this is a fault in the merge rather than in the workbook. Send the "
            "workbook to the maintainer.",
        ))
        return out

    dangling = _dangling_refs(document)
    if dangling:
        out.findings.append(Finding(
            "X004", "error", "(merged)", "",
            "%d reference%s point at a schema that is not in the document: %s"
            % (len(dangling), "" if len(dangling) == 1 else "s",
               ", ".join(sorted(dangling)[:8])),
            "every $ref to resolve within components/schemas",
            "this is a fault in the merge rather than in the workbook. Send the "
            "workbook to the maintainer.",
        ))
        return out

    out.document = document
    out.stats = {
        "operations": len(out.operations),
        "schemas": len(schemas),
        "hoisted_groups": len(registry.shapes),
        "split_groups": len(registry.base_names),
        "tags": len(tags),
    }
    return out


def _service_domains(ok, findings):
    """The distinct service domains, one canonical spelling each.

    The reference workbook writes the same domain three ways, as ``Credit
    Card``, ``Credit card`` and ``CreditCard``. Left alone that produces three
    tags and a nonsense title, so the spellings are folded together and the
    most-used one wins. Unlike a schema name, a tag is not part of the
    published payload, so folding it is safe; the inconsistency is still
    reported so the workbook can be tidied.
    """
    counts = defaultdict(lambda: defaultdict(int))
    for r, _s in ok:
        raw = text(r.banner.get("service_domain", ""))
        if raw:
            counts[key(raw)][raw] += 1
    out = []
    for k in sorted(counts):
        spellings = counts[k]
        chosen = max(sorted(spellings), key=lambda s: spellings[s])
        out.append(chosen)
        if len(spellings) > 1:
            findings.append(Finding(
                "N002", "warning", "(merged)", "",
                "the service domain is spelled %d ways across the workbook: %s"
                % (len(spellings),
                   ", ".join(repr(s) for s in sorted(spellings))),
                "one spelling of each service domain",
                "settle on %r in the Service Domain cell of every sheet. The "
                "generated tag uses that spelling in the meantime." % chosen,
            ))
    return out


def _canonical_domain(raw, domains):
    """The chosen spelling for a sheet's service domain."""
    k = key(raw)
    for d in domains:
        if key(d) == k:
            return d
    return text(raw)


def _op_id(spec):
    return spec.document["paths"][spec.path][spec.method]["operationId"]


def _require_tag(entry):
    """Make the discriminator property required on a variant schema."""
    target = entry["allOf"][1] if "allOf" in entry else entry
    required = target.setdefault("required", [])
    if VARIANT_PROPERTY not in required:
        required.append(VARIANT_PROPERTY)


def _annotate(body, uses, base=None):
    """Record which operations a component serves."""
    out = dict(body)
    operations = sorted({c.operation_id for c, _n in uses})
    if len(operations) > 1 or base is not None:
        out["x-used-by"] = operations
    return out


def _default_description(ok, domains):
    return (
        "Generated from the SOR mapping workbook. %d operation%s across %s. "
        "Only attributes mapped to a System of Record field are published; "
        "shared groups that differ between operations are specialised with "
        "allOf, and operations served by several SOR endpoints carry a "
        "requestVariant discriminator."
        % (len(ok), "" if len(ok) == 1 else "s",
           ", ".join(domains) if domains else "the workbook"))


def _dangling_refs(document):
    """Every ``$ref`` that does not resolve inside components/schemas."""
    known = set(document.get("components", {}).get("schemas", {}))
    missing = set()

    def walk(node):
        if isinstance(node, dict):
            ref = node.get("$ref")
            if isinstance(ref, str) and ref.startswith("#/components/schemas/"):
                name = ref.rsplit("/", 1)[-1]
                if name not in known:
                    missing.add(name)
            for v in node.values():
                walk(v)
        elif isinstance(node, list):
            for v in node:
                walk(v)
    walk(document)
    return missing


__all__ = ["Context", "MergedSpec", "Registry", "datasig", "emit_body",
           "merge", "split_shapes"]
