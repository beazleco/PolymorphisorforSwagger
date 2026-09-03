#!/usr/bin/env python3
"""
Build a synthetic SOR specification from a mapping workbook.

This exists so the verification path can be exercised end to end without real
SOR documentation. It is a **test fixture, not SOR documentation**: it is
derived from the workbook's own SOR column values, so of course most of them
resolve against it. Three defects are introduced on purpose so that the
interesting paths are exercised too:

* a handful of leaves are omitted, so those elements are excluded (``R010``)
* one leaf is duplicated in a second place, so a match is ambiguous (``R011``)
* one endpoint declares only ``PUT`` where a sheet states ``DELETE`` (``R005``)

Usage::

    python tools/build_sor_fixture.py <workbook.xlsx> <out.yaml>
"""

from __future__ import annotations

import os
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import polymorphize_generate as gen                            # noqa: E402
import polymorphize_sor as sor                                 # noqa: E402
import polymorphize_workbook as wbk                            # noqa: E402

#: Leaf names deliberately left out of the fixture, so their elements are
#: excluded from the interface and reported. Chosen to be spread across sheets
#: and to include one that several sheets rely on.
OMIT = {
    "cancelDate",          # response, several sheets
    "issueReason",         # response
    "customerLimitCurrency",
    "postingTime",         # one half of a two-field mapping, so the other half
                           # still resolves and the element survives
    "accountName",
}

#: A leaf duplicated in a second place, to make one match ambiguous.
DUPLICATE = ("currencyCode", "auditTrail")

#: Endpoints for which only these methods are declared, whatever the workbook
#: states. Used to provoke exactly one method mismatch.
METHOD_OVERRIDE = {"/v1/card/cancel": ["put"]}


def collect(workbook):
    """``{(endpoint, side): {dotted path}}`` from every sheet's SOR columns."""
    fields = defaultdict(set)
    methods = {}
    results = wbk.read_workbook(workbook)
    for r in results:
        if r.status == "skipped" or r.layout is None:
            continue
        decisions = sor.decide_endpoints(r, [])
        by_column = {d.column.endpoint: d for d in decisions}
        for side, section in (("request", r.request), ("response", r.response)):
            if section is None or section.root is None:
                continue

            def walk(node):
                for child in node.children:
                    for column in r.layout.sor_cols:
                        raw = child.sor.get(column.endpoint)
                        if not raw:
                            continue
                        d = by_column.get(column.endpoint)
                        if d is None or not d.endpoint:
                            continue
                        for declared, _leaf in sor.candidates_of(raw):
                            fields[(d.endpoint, side)].add(declared)
                        methods.setdefault(d.endpoint, d.stated_method or "get")
                    walk(child)
            walk(section.root)
    return fields, methods


LEAF = "\x00leaf"


def nest(paths):
    """A set of dotted paths as a nested OpenAPI object schema.

    A name can be both a leaf and a branch across different rows, for instance
    ``accountInfo.limit`` alongside ``accountInfo.limit.value``, so every node
    is a dictionary and a leaf is marked rather than represented by ``None``.
    """
    tree = {}
    for path in sorted(paths):
        parts = [p for p in path.split(".") if p]
        node = tree
        for part in parts[:-1]:
            node = node.setdefault(part, {})
            if not isinstance(node, dict):
                node = {}
        if parts:
            node.setdefault(parts[-1], {})[LEAF] = True

    def build(node):
        props = {}
        for name, child in sorted(node.items()):
            if name == LEAF or name in OMIT:
                continue
            children = {k: v for k, v in child.items() if k != LEAF} \
                if isinstance(child, dict) else {}
            props[name] = build(children) if children else {"type": "string"}
        return {"type": "object", "properties": props}
    return build(tree)


def main(argv=None):
    args = list(argv if argv is not None else sys.argv[1:])
    if len(args) < 2:
        print(__doc__)
        return 2
    workbook, out_path = args[0], args[1]

    fields, methods = collect(workbook)
    endpoints = sorted({e for e, _s in fields})
    paths = {}
    for endpoint in endpoints:
        method = METHOD_OVERRIDE.get(endpoint, [methods.get(endpoint, "get")])[0]
        request = nest(fields.get((endpoint, "request"), set()))
        response = nest(fields.get((endpoint, "response"), set()))

        operation = {
            "operationId": "sor" + "".join(
                p.title() for p in endpoint.strip("/").replace("{}", "By")
                .split("/")),
            "summary": "SOR operation %s" % endpoint,
            "responses": {"200": {
                "description": "Success.",
                "content": {"application/json": {"schema": response}},
            }},
        }
        if method in sor.PARAMETER_METHODS:
            operation["parameters"] = [
                {"name": name, "in": "query", "required": False,
                 "schema": {"type": "string"}}
                for name in sorted((request.get("properties") or {}))]
        else:
            operation["requestBody"] = {
                "required": True,
                "content": {"application/json": {"schema": request}},
            }
        entry = {method: operation}
        if "{}" in endpoint:
            entry["parameters"] = [
                {"name": "id", "in": "path", "required": True,
                 "schema": {"type": "string"}}]
        paths[endpoint.replace("{}", "{id}")] = entry

    # The duplicate, added to one response so a leaf name matches twice.
    leaf, holder = DUPLICATE
    for endpoint in endpoints:
        entry = paths[endpoint.replace("{}", "{id}")]
        for method, operation in entry.items():
            if method == "parameters":
                continue
            schema = ((operation.get("responses") or {}).get("200") or {}) \
                .get("content", {}).get("application/json", {}).get("schema")
            if not isinstance(schema, dict):
                continue
            props = schema.setdefault("properties", {})
            if leaf in props:
                props[holder] = {"type": "object",
                                 "properties": {leaf: {"type": "string"}}}
                break
        else:
            continue
        break

    document = {
        "openapi": "3.0.3",
        "info": {
            "title": "Vertexon CMS, synthetic SOR specification",
            "version": "0.0.1-fixture",
            "description":
                "TEST FIXTURE, NOT SOR DOCUMENTATION. Generated by "
                "tools/build_sor_fixture.py from the SOR column values of the "
                "mapping workbook, so most of them resolve against it by "
                "construction. Three defects are deliberate: %d leaf names are "
                "omitted, %r is duplicated under %r, and %s declares only PUT. "
                "Replace this with the real SOR specification before drawing "
                "any conclusion from a run."
                % (len(OMIT), DUPLICATE[0], DUPLICATE[1],
                   ", ".join(METHOD_OVERRIDE)),
            "x-fixture": True,
            "x-omitted-leaves": sorted(OMIT),
        },
        "paths": paths,
    }
    with open(out_path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(gen.dump_document(document))
    print("Wrote %s: %d endpoints" % (out_path, len(paths)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
