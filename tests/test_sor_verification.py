#!/usr/bin/env python3
"""
Regression suite for verification against the System of Record (v6.2).

Run from the package root::

    python -m tests.test_sor_verification

Groups
======

V1   path and endpoint-list parsing from the messy banner forms
V2   indexing an SOR specification: bodies, parameters, refs, allOf, arrays
V3   resolving a field name by its leaf, including multi-value cells
V4   endpoint identification: the banner is definitive
V5   exclusion: an unresolved field removes the element from the interface
V6   the guards: nothing resolved, and an endpoint absent from the SOR
V7   the showcase report
V8   the reference workbook against its SOR fixture, end to end
"""

from __future__ import annotations

import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import polymorphize_generate as gen                            # noqa: E402
import polymorphize_showcase as show                           # noqa: E402
import polymorphize_sor as sor                                 # noqa: E402
import polymorphize_workbook as wbk                            # noqa: E402
from tests.test_workbook_contract import (                     # noqa: E402
    STD_BANNER, build, one_sheet_workbook, read_one, sheet_rows, tmp,
)

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
CREDIT_CARD = os.path.join(ROOT, "samples", "creditcard_v2.0.0.xlsx")
SOR_FIXTURE = os.path.join(ROOT, "samples", "sor_creditcard_fixture.yaml")

PASSED, FAILED = 0, 0


def group(title):
    print("")
    print(title)


def check(label, condition, detail=""):
    global PASSED, FAILED
    if condition:
        PASSED += 1
        print("  PASS  %s" % label)
    else:
        FAILED += 1
        print("  FAIL  %s%s" % (label, ("   " + str(detail)) if detail else ""))


def eq(label, actual, expected):
    check(label, actual == expected, "got %r, expected %r" % (actual, expected))


# --------------------------------------------------------------------------- #
# Fixtures
# --------------------------------------------------------------------------- #


def write_spec(document, path):
    with open(path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(gen.dump_document(document))
    return path


def simple_sor(path, *, request_leaves=("sorId", "ccy"),
               response_leaves=("balAmt", "balCcy"), method="post",
               endpoint="/v1/balance", extra_paths=None):
    """A small SOR specification with the leaves named."""
    op = {
        "operationId": "sorOp",
        "responses": {"200": {
            "description": "ok",
            "content": {"application/json": {"schema": {
                "type": "object",
                "properties": {n: {"type": "string"} for n in response_leaves},
            }}},
        }},
    }
    if method in sor.PARAMETER_METHODS:
        op["parameters"] = [{"name": n, "in": "query",
                             "schema": {"type": "string"}}
                            for n in request_leaves]
    else:
        op["requestBody"] = {"content": {"application/json": {"schema": {
            "type": "object",
            "properties": {n: {"type": "string"} for n in request_leaves},
        }}}}
    paths = {endpoint: {method: op}}
    paths.update(extra_paths or {})
    return write_spec({"openapi": "3.0.3",
                       "info": {"title": "SOR", "version": "1"},
                       "paths": paths}, path)


REQ = [
    (1, "ThingRequest", "", "object", "", "Wrapper."),
    (2, "Identifier", "", "object", "", "Group."),
    (3, "idType", "IBAN", "String (12)", "", "No SOR field at all."),
    (3, "idValue", "123", "String (34)", "Mandatory", "Maps to sorId.", "sorId"),
    (2, "currencyCode", "GBP", "String (3)", "", "Maps to ccy.", "ccy"),
]

RES = [
    (1, "ThingResponse", "", "object", "", "Wrapper."),
    (2, "balanceAmount", "1.00", "Number(18)", "Mandatory", "Maps.", "balAmt"),
    (2, "balanceCurrency", "GBP", "String (3)", "", "Maps.", "balCcy"),
]


def workbook_with(path, request_rows=None, response_rows=None, banner=None,
                  sor_headers=None, title="Thing_Retrieve"):
    return one_sheet_workbook(
        path, title=title,
        banner=banner or STD_BANNER,
        sor_headers=sor_headers if sor_headers is not None else ["/v1/balance"],
        sections=[("Request Body", request_rows or REQ),
                  ("Response Body", response_rows or RES)])


# --------------------------------------------------------------------------- #
# V1 paths
# --------------------------------------------------------------------------- #


def v1():
    group("V1 — reading the endpoint out of the banner")
    eq("a plain path is unchanged",
       sor.normalise_path("/v1/card/availableFunds"), "/v1/card/availableFunds")
    eq("a method prefix with a colon is dropped",
       sor.normalise_path("GET: /v1/a"), "/v1/a")
    eq("a method prefix without a colon is dropped",
       sor.normalise_path("PUT /v1/a"), "/v1/a")
    eq("a missing leading slash is added",
       sor.normalise_path("v1/account/search/trans"), "/v1/account/search/trans")
    eq("a trailing slash is dropped", sor.normalise_path("/v1/a/"), "/v1/a")
    eq("a path parameter is generalised, so its name need not match",
       sor.normalise_path("/v1/card/{cardNumber}/x"), "/v1/card/{}/x")
    eq("and matches whatever the SOR calls it",
       sor.normalise_path("/v1/card/{id}/x"),
       sor.normalise_path("/v1/card/{cardNumber}/x"))
    eq("an empty cell yields nothing", sor.normalise_path(""), "")

    # Every banner form in the reference workbook.
    eq("comma separated with methods",
       sor.split_endpoints("GET: v1/account/search/trans, GET: "
                           "v1/account/search/auth, GET: /v1/account/search/mps"),
       ["/v1/account/search/trans", "/v1/account/search/auth",
        "/v1/account/search/mps"])
    eq("space separated, no methods",
       sor.split_endpoints("/v1/account/contact /v1/account/delinquency"),
       ["/v1/account/contact", "/v1/account/delinquency"])
    eq("space separated with methods",
       sor.split_endpoints("GET: /v1/account/cards GET: /v1/card/details"),
       ["/v1/account/cards", "/v1/card/details"])
    eq("newline separated",
       sor.split_endpoints("/v1/a\n/v1/b"), ["/v1/a", "/v1/b"])
    eq("one endpoint", sor.split_endpoints("PUT /v1/card/cancel"),
       ["/v1/card/cancel"])
    eq("nothing at all", sor.split_endpoints(""), [])

    eq("a stated method is picked up", sor.stated_method("DELETE: /v1/a"),
       "delete")
    eq("and absent when there is none", sor.stated_method("/v1/a"), "")


# --------------------------------------------------------------------------- #
# V2 indexing
# --------------------------------------------------------------------------- #


def v2():
    group("V2 — indexing an SOR specification")
    path = simple_sor(tmp("sor.yaml"))
    index = sor.SorIndex.from_paths([path])
    eq("one file", index.summary()["files"], 1)
    eq("one operation", index.summary()["operations"], 1)
    check("the path is known", index.has_path("/v1/balance"))
    check("and known through a messy spelling",
          index.has_path("POST: v1/balance/"))
    op = index.pick("/v1/balance")
    eq("the request leaves are indexed", sorted(op.request.leaves),
       ["ccy", "sorid"])
    eq("the response leaves are indexed", sorted(op.response.leaves),
       ["balamt", "balccy"])

    # Parameters are the request of a GET.
    p2 = simple_sor(tmp("sor2.yaml"), method="get")
    op2 = sor.SorIndex.from_paths([p2]).pick("/v1/balance")
    eq("a GET's query parameters are its request", sorted(op2.request.leaves),
       ["ccy", "sorid"])

    # Refs, allOf and arrays are all followed.
    doc = {
        "openapi": "3.0.3", "info": {"title": "S", "version": "1"},
        "paths": {"/v1/deep": {"get": {"responses": {"200": {
            "description": "ok",
            "content": {"application/json": {"schema": {
                "$ref": "#/components/schemas/Top"}}}}}}}},
        "components": {"schemas": {
            "Top": {"allOf": [
                {"$ref": "#/components/schemas/Base"},
                {"type": "object", "properties": {
                    "rows": {"type": "array", "items": {
                        "$ref": "#/components/schemas/Row"}}}}]},
            "Base": {"type": "object",
                     "properties": {"id": {"type": "string"}}},
            "Row": {"type": "object", "properties": {
                "deepValue": {"type": "string"},
                "nested": {"type": "object", "properties": {
                    "innermost": {"type": "string"}}}}},
        }},
    }
    p3 = write_spec(doc, tmp("sor3.yaml"))
    op3 = sor.SorIndex.from_paths([p3]).pick("/v1/deep")
    for leaf in ("id", "rows", "deepValue", "nested", "innermost"):
        check("%s is found through refs, allOf and arrays" % leaf,
              bool(op3.response.find(leaf)), sorted(op3.response.leaves))
    eq("and its path is recorded", op3.response.find("innermost"),
       ["rows.nested.innermost"])

    # A recursive schema must not hang.
    rec = {
        "openapi": "3.0.3", "info": {"title": "S", "version": "1"},
        "paths": {"/v1/rec": {"get": {"responses": {"200": {
            "description": "ok",
            "content": {"application/json": {"schema": {
                "$ref": "#/components/schemas/Node"}}}}}}}},
        "components": {"schemas": {"Node": {"type": "object", "properties": {
            "name": {"type": "string"},
            "child": {"$ref": "#/components/schemas/Node"}}}}},
    }
    p4 = write_spec(rec, tmp("sor4.yaml"))
    op4 = sor.SorIndex.from_paths([p4]).pick("/v1/rec")
    check("a recursive schema terminates", bool(op4.response.find("name")))

    # Several files, and a folder.
    folder = os.path.dirname(simple_sor(tmp("a.yaml"), endpoint="/v1/one"))
    simple_sor(os.path.join(folder, "b.yaml"), endpoint="/v1/two")
    both = sor.SorIndex.from_paths([folder])
    eq("a folder loads every specification in it",
       sorted(both.by_path), ["/v1/one", "/v1/two"])

    # A file that is not a specification is reported, not fatal.
    junk = tmp("junk.yaml")
    with open(junk, "w", encoding="utf-8") as fh:
        fh.write("just: some\nother: yaml\n")
    bad = sor.SorIndex.from_paths([junk])
    eq("a file with no paths section is reported", len(bad.load_errors), 1)
    check("and the index is empty rather than broken", bad.empty)


# --------------------------------------------------------------------------- #
# V3 resolving
# --------------------------------------------------------------------------- #


def v3():
    group("V3 — resolving a field name by its leaf")
    doc = {
        "openapi": "3.0.3", "info": {"title": "S", "version": "1"},
        "paths": {"/v1/x": {"get": {"responses": {"200": {
            "description": "ok",
            "content": {"application/json": {"schema": {
                "type": "object", "properties": {
                    "accountInfo": {"type": "object", "properties": {
                        "currency": {"type": "string"},
                        "limit": {"type": "string"}}},
                    "cards": {"type": "array", "items": {
                        "type": "object", "properties": {
                            "currency": {"type": "string"},
                            "cardToken": {"type": "string"}}}},
                }}}}}}}}},
    }
    op = sor.SorIndex.from_paths([write_spec(doc, tmp("s.yaml"))]).pick("/v1/x")

    r = sor.resolve("cardToken", op, "response")
    eq("a bare name resolves", r.status, sor.RESOLVED)
    eq("and records where", r.at, ["cards.cardToken"])

    r = sor.resolve("accountInfo.limit", op, "response")
    eq("a dotted path resolves on its last segment", r.status, sor.RESOLVED)

    r = sor.resolve("something.else.entirely.limit", op, "response")
    eq("the path itself is ignored, by decision", r.status, sor.RESOLVED)
    eq("and it matched the leaf", r.matched, "limit")

    r = sor.resolve("accountInfo.currency", op, "response")
    eq("a leaf in two places is ambiguous", r.status, sor.AMBIGUOUS)
    eq("and both places are recorded", sorted(r.at),
       ["accountInfo.currency", "cards.currency"])

    r = sor.resolve("noSuchField", op, "response")
    eq("an absent field is missing", r.status, sor.MISSING)

    r = sor.resolve("noSuchField, cardToken", op, "response")
    eq("a comma separated cell resolves if any candidate does",
       r.status, sor.RESOLVED)
    eq("and reports which", r.matched, "cardToken")

    r = sor.resolve("a.limit b.cardToken", op, "response")
    eq("a space separated cell resolves too", r.status, sor.RESOLVED)

    r = sor.resolve("nothingA, nothingB", op, "response")
    eq("and is missing when none of them do", r.status, sor.MISSING)
    eq("with every candidate listed", r.candidates, ["nothingA", "nothingB"])

    eq("the request side is searched separately",
       sor.resolve("cardToken", op, "request").status, sor.MISSING)


# --------------------------------------------------------------------------- #
# V4 endpoint identification
# --------------------------------------------------------------------------- #


def v4():
    group("V4 — which endpoint: the banner is definitive")
    r = read_one(workbook_with(tmp("a.xlsx")))
    findings = []
    decisions = sor.decide_endpoints(r, findings)
    eq("one column", len(decisions), 1)
    eq("the banner names the endpoint", decisions[0].endpoint, "/v1/balance")
    eq("and the banner is always the source", decisions[0].source, "banner")
    eq("nothing is reported when the header agrees", findings, [])

    # A header naming nothing changes nothing: only the banner is read.
    r2 = read_one(workbook_with(tmp("b.xlsx"), sor_headers=[""]))
    f2 = []
    d2 = sor.decide_endpoints(r2, f2)
    eq("an unheaded column still gets the banner's endpoint",
       d2[0].endpoint, "/v1/balance")
    eq("with nothing reported", f2, [])

    # The disagreement, which is the reference workbook's Card RSAEncrypted.
    # The banner wins and the difference is reported, not fatal.
    banner = list(STD_BANNER)
    banner[4] = ("SOR API Endpoint:", "GET /v1/card/{cardNumber}/rsa", "")
    r3 = read_one(workbook_with(tmp("c.xlsx"), banner=banner,
                                sor_headers=["/v1/card/cancel"]))
    f3 = []
    d3 = sor.decide_endpoints(r3, f3)
    eq("the banner wins", d3[0].endpoint, "/v1/card/{}/rsa")
    check("and the difference is reported", any(f.code == "R001" for f in f3))
    finding = next(f for f in f3 if f.code == "R001")
    eq("as a warning, not an error", finding.severity, "warning")
    check("naming what the header says",
          "/v1/card/cancel" in finding.what, finding.what)
    check("what the banner says", "rsa" in finding.what, finding.what)
    check("and that the banner is definitive",
          "banner is definitive" in finding.what, finding.what)
    check("with the correction pointed at the header",
          "correct the column header" in finding.fix, finding.fix)

    # The banner's endpoints must be pairable with the columns.
    banner2 = list(STD_BANNER)
    banner2[4] = ("SOR API Endpoint:", "/v1/a /v1/b /v1/c", "")
    r4 = read_one(workbook_with(tmp("d.xlsx"), banner=banner2,
                                sor_headers=[""]))
    f4 = []
    d4 = sor.decide_endpoints(r4, f4)
    check("three banner endpoints against one column is an error",
          any(f.code == "R002" for f in f4))
    eq("and no endpoint is chosen", [d.endpoint for d in d4], [""])

    # Nothing in the banner is now the only way to have no endpoint.
    banner3 = list(STD_BANNER)
    banner3[4] = ("SOR API Endpoint:", "", "")
    r5 = read_one(workbook_with(tmp("e.xlsx"), banner=banner3,
                                sor_headers=["/v1/balance"]))
    f5 = []
    d5 = sor.decide_endpoints(r5, f5)
    check("an empty banner is an error even when a header names one",
          any(f.code == "R003" for f in f5))
    eq("and the header is not used as a fallback", d5[0].endpoint, "")
    r003 = next(f for f in f5 if f.code == "R003")
    check("the message says the banner is the only source read",
          "only place" in r003.what, r003.what)

    # One banner endpoint serves every column.
    banner4 = list(STD_BANNER)
    banner4[4] = ("SOR API Endpoint:", "GET: /v1/shared", "")
    multi_req = [
        (1, "MultiRequest", "", "object", "", "Root."),
        (2, "shared", "1", "String (5)", "Mandatory", "Both.", "sorId", "sorId"),
    ]
    r6 = read_one(workbook_with(tmp("f.xlsx"), request_rows=multi_req,
                                banner=banner4, sor_headers=["", ""]))
    f6 = []
    d6 = sor.decide_endpoints(r6, f6)
    eq("one endpoint serves both columns", [d.endpoint for d in d6],
       ["/v1/shared", "/v1/shared"])
    eq("with nothing reported", f6, [])

    # Several are paired with the columns in order.
    banner5 = list(STD_BANNER)
    banner5[4] = ("SOR API Endpoint:", "GET: /v1/a, GET: /v1/b", "")
    r7 = read_one(workbook_with(tmp("g.xlsx"), request_rows=multi_req,
                                banner=banner5, sor_headers=["/v1/a", "/v1/b"]))
    f7 = []
    d7 = sor.decide_endpoints(r7, f7)
    eq("paired in order", [d.endpoint for d in d7], ["/v1/a", "/v1/b"])
    eq("with nothing reported", f7, [])
    eq("and the method comes from that endpoint's own banner token",
       [d.stated_method for d in d7], ["get", "get"])


# --------------------------------------------------------------------------- #
# V5 exclusion
# --------------------------------------------------------------------------- #


def v5():
    group("V5 — an unresolved field removes the element")
    workbook = workbook_with(tmp("a.xlsx"))
    # ccy is absent from the SOR, so currencyCode must go.
    spec_path = simple_sor(tmp("sor.yaml"), request_leaves=("sorId",))
    index = sor.SorIndex.from_paths([spec_path])

    r = read_one(workbook)
    before = dict(r.request.root.children[1].sor)
    findings = []
    v = sor.verify_sheet(r, index, findings)
    eq("the sheet was checked", v.checked, True)
    eq("one field resolved on the request", v.counts[sor.RESOLVED], 3)
    eq("and one did not", v.counts[sor.MISSING], 1)

    check("the removal is reported as R010",
          any(f.code == "R010" for f in findings))
    f = next(f for f in findings if f.code == "R010")
    check("naming the element", "currencyCode" in f.what, f.what)
    check("the field the workbook wrote", "'ccy'" in f.what, f.what)
    check("the message searched", "request" in f.what, f.what)
    check("the endpoint searched", "/v1/balance" in f.what, f.what)
    check("and that the element was removed",
          "removed from the interface" in f.what, f.what)
    check("with the cell to correct, on the row the attribute is on",
          f.cell[1:].isdigit() and int(f.cell[1:]) == 12, f.cell)

    eq("the effective mapping is cleared",
       r.request.root.children[1].sor.get("/v1/balance"), "")
    eq("but the workbook's word is kept for the report",
       r.request.root.children[1].sor_declared.get("/v1/balance"), "ccy")
    check("which is what it was before verification", bool(before))

    spec = gen.generate_sheet(r)
    props = spec.document["components"]["schemas"]["ThingRequest"]["properties"]
    check("so the element is not in the interface", "currencyCode" not in props,
          sorted(props))
    check("while the resolved one is",
          "idValue" in props["Identifier"]["properties"])

    # A summary finding names the totals.
    check("a per-sheet summary is reported",
          any(f.code == "R021" for f in findings))

    # An ambiguity keeps the element and says so.
    amb = simple_sor(tmp("sor2.yaml"), request_leaves=("sorId", "ccy"),
                     response_leaves=("balAmt", "balCcy"),
                     extra_paths=None)
    doc_path = tmp("sor3.yaml")
    write_spec({"openapi": "3.0.3", "info": {"title": "S", "version": "1"},
                "paths": {"/v1/balance": {"post": {
                    "requestBody": {"content": {"application/json": {"schema": {
                        "type": "object", "properties": {
                            "sorId": {"type": "string"},
                            "a": {"type": "object", "properties": {
                                "ccy": {"type": "string"}}},
                            "b": {"type": "object", "properties": {
                                "ccy": {"type": "string"}}}}}}}},
                    "responses": {"200": {
                        "description": "ok",
                        "content": {"application/json": {"schema": {
                            "type": "object", "properties": {
                                "balAmt": {"type": "string"},
                                "balCcy": {"type": "string"}}}}}}}}}}},
               doc_path)
    r2 = read_one(workbook_with(tmp("b.xlsx")))
    f2 = []
    sor.verify_sheet(r2, sor.SorIndex.from_paths([doc_path]), f2)
    check("an ambiguous match is reported as R011",
          any(f.code == "R011" for f in f2))
    amb_f = next(f for f in f2 if f.code == "R011")
    check("naming both places", "a.ccy" in amb_f.what and "b.ccy" in amb_f.what,
          amb_f.what)
    eq("but the element is kept",
       r2.request.root.children[1].sor.get("/v1/balance"), "ccy")

    # Without an SOR specification nothing is verified and nothing is lost.
    r3 = read_one(workbook_with(tmp("c.xlsx")))
    f3 = []
    v3_ = sor.verify_sheet(r3, None, f3)
    eq("no SOR means no verification", v3_.checked, False)
    eq("and no findings", f3, [])
    eq("and the mapping is untouched",
       r3.request.root.children[1].sor.get("/v1/balance"), "ccy")


# --------------------------------------------------------------------------- #
# V6 the guards
# --------------------------------------------------------------------------- #


def v6():
    group("V6 — the guards")
    # An endpoint the SOR specifications do not define. The banner has to
    # agree with the header, otherwise the conflict is reported first and
    # verification stops before it gets this far.
    banner_nowhere = list(STD_BANNER)
    banner_nowhere[4] = ("SOR API Endpoint:", "GET /v1/nowhere", "")
    workbook = workbook_with(tmp("a.xlsx"), banner=banner_nowhere,
                             sor_headers=["/v1/nowhere"])
    index = sor.SorIndex.from_paths([simple_sor(tmp("sor.yaml"))])
    r = read_one(workbook)
    findings = []
    v = sor.verify_sheet(r, index, findings)
    check("an unknown endpoint is an error",
          any(f.code == "R004" for f in findings))
    f = next(f for f in findings if f.code == "R004")
    check("naming the endpoint", "/v1/nowhere" in f.what, f.what)
    check("and telling the analyst what to supply",
          "supply the SOR specification" in f.fix.lower() or
          "supply the sor specification" in f.fix.lower(), f.fix)
    eq("nothing was checked", v.checked, False)

    # Every field absent means the wrong endpoint was named.
    empty = simple_sor(tmp("sor2.yaml"), request_leaves=("nothing",),
                       response_leaves=("nothingElse",))
    r2 = read_one(workbook_with(tmp("b.xlsx")))
    f2 = []
    sor.verify_sheet(r2, sor.SorIndex.from_paths([empty]), f2)
    check("nothing resolving at all is an error, not a silent emptying",
          any(f.code == "R020" for f in f2))
    g = next(f for f in f2 if f.code == "R020")
    check("and it says why that is almost certainly a pointing error",
          "wrong endpoint" in g.what, g.what)

    # A method the SOR does not define is a warning, and the run continues.
    banner = list(STD_BANNER)
    banner[4] = ("SOR API Endpoint:", "DELETE /v1/balance", "")
    r3 = read_one(workbook_with(tmp("c.xlsx"), banner=banner,
                                sor_headers=[""]))
    f3 = []
    v3_ = sor.verify_sheet(r3, sor.SorIndex.from_paths(
        [simple_sor(tmp("sor3.yaml"), method="post")]), f3)
    check("a method mismatch is a warning",
          any(f.code == "R005" and f.severity == "warning" for f in f3))
    eq("and the sheet is still checked", v3_.checked, True)


# --------------------------------------------------------------------------- #
# V7 the showcase
# --------------------------------------------------------------------------- #


def v7():
    group("V7 — the showcase report")
    workbook = workbook_with(tmp("a.xlsx"))
    spec_path = simple_sor(tmp("sor.yaml"), request_leaves=("sorId",))
    out_dir = tmp("out")
    out = gen.run(workbook, out_dir, sor=[spec_path], log=lambda *_a: None)
    eq("the endpoint generated", len(out.ok), 1)
    check("a showcase was written", os.path.exists(out.showcase_path),
          out.showcase_path)
    html = open(out.showcase_path, encoding="utf-8").read()

    check("it says the mappings were verified",
          "checked against the System of Record" in html)
    check("it names the workbook", os.path.basename(workbook) in html)
    for label in ("Published", "No SOR field", "Not in the SOR", "Container"):
        check("the %r disposition appears" % label, label in html)
    check("the element removed for want of a real SOR field is shown",
          "currencyCode" in html)
    check("with the field name the workbook wrote", ">ccy<" in html or
          "ccy" in html)
    check("the element with no SOR field at all is shown", "idType" in html)
    check("the published element is shown", "idValue" in html)
    check("nothing was silently lost", "LOST" not in html.replace(
        "LOST", "", 0) or ">LOST<" not in html)
    check("the reasoning is spelled out", "How this was decided" in html)
    check("it is self contained, with no external asset",
          "http://" not in html and "https://" not in html)

    views = show.build_views(out)
    eq("one view per generated endpoint", len(views), 1)
    v = views[0]
    dispositions = {}
    for m in v.messages:
        for cls in m.classes:
            for row in cls.rows:
                dispositions[row.node.name] = row.disposition
    eq("idType has no SOR field", dispositions["idType"], show.NO_FIELD)
    eq("currencyCode is not in the SOR", dispositions["currencyCode"],
       show.NOT_IN_SOR)
    eq("idValue is published", dispositions["idValue"], show.PUBLISHED)
    eq("Identifier is a container", dispositions["Identifier"], show.CONTAINER)
    check("nothing is marked lost",
          show.LOST not in dispositions.values(), dispositions)

    # Without an SOR specification the report says so, loudly.
    out2 = gen.run(workbook, tmp("out2"), log=lambda *_a: None)
    html2 = open(out2.showcase_path, encoding="utf-8").read()
    check("the unverified report warns that nothing was checked",
          "No System of Record specification was supplied" in html2)
    check("and says the workbook was taken on trust",
          "taken on trust" in html2)


# --------------------------------------------------------------------------- #
# V8 the reference workbook
# --------------------------------------------------------------------------- #


def v8():
    group("V8 — the reference workbook against its SOR fixture")
    if not (os.path.exists(CREDIT_CARD) and os.path.exists(SOR_FIXTURE)):
        check("the reference workbook and its SOR fixture are present", False,
              "%s / %s" % (CREDIT_CARD, SOR_FIXTURE))
        return

    index = sor.SorIndex.from_paths([SOR_FIXTURE])
    eq("the fixture loads", index.load_errors, [])
    check("covering every endpoint the workbook names",
          index.summary()["operations"] >= 14, index.summary())

    out = gen.run(CREDIT_CARD, tmp("cc"), sor=[SOR_FIXTURE],
                  log=lambda *_a: None)
    eq("verification ran", out.sor_verified, True)
    eq("eleven endpoints generate", len(out.ok), 11)
    failed = sorted(s.sheet for s in out.failed)
    eq("and only the mislabelled sheet fails", failed,
       ["Card Issued Device_Retrieve"])

    codes = {s.sheet: {f.code for f in s.findings if f.severity == "error"}
             for s in out.failed}
    check("the mislabelled section fails on S001",
          "S001" in codes["Card Issued Device_Retrieve"], codes)

    # The sheet whose header and banner disagree now generates against the
    # banner's endpoint, with the difference reported.
    rsa = next(s for s in out.ok if s.sheet == "Card RSAEncrypted")
    r001 = [f for f in rsa.findings if f.code == "R001"]
    check("the banner and header difference is a warning on a generated "
          "endpoint", bool(r001) and r001[0].severity == "warning", r001)
    check("and the banner's endpoint was the one used",
          "cardRSAEncrypted" in r001[0].what, r001[0].what)

    r010 = [f for s in out.specs for f in s.findings if f.code == "R010"]
    check("the deliberately omitted SOR fields are caught", len(r010) >= 2,
          len(r010))
    names = " ".join(f.what for f in r010)
    check("including customerLimitCurrency", "customerLimitCurrency" in names)
    check("and accountName", "accountName" in names)

    r011 = [f for s in out.specs for f in s.findings if f.code == "R011"]
    check("and the duplicated leaf is reported as ambiguous", bool(r011),
          len(r011))

    # A two-candidate cell where only one candidate is absent keeps the element.
    adj = next(s for s in out.ok if s.sheet == "Transaction Adjustment_Initiate")
    doc = gen.dump_document(adj.document)
    check("a cell naming two SOR fields survives one of them being absent",
          "systemEffectiveDate" in doc)

    check("a showcase was written", os.path.exists(out.showcase_path))
    eq("named after the first service domain",
       os.path.basename(out.showcase_path), "credit-card.html")
    html = open(out.showcase_path, encoding="utf-8").read()
    check("it reports the removals for want of a real SOR field",
          "Not in the SOR" in html)
    check("and names the endpoints that were not generated",
          "Endpoints not generated" in html)
    check("nothing was silently lost", ">LOST<" not in html)
    check("the eliminations have a section of their own",
          "Elements eliminated from the interface" in html)
    check("the navigation lists the integration APIs, not the SOR endpoints",
          "INTEGRATION APIS" in html.upper())
    for path in ("/CreditCard/Account/Retrieve",
                 "/credit-card/billing/retrieve"):
        check("the integration path %s appears" % path, path in html)

    for spec in out.ok:
        eq("%s carries no empty object" % spec.sheet,
           gen.empty_schemas(spec.document), [])


# --------------------------------------------------------------------------- #


def main():
    for fn in (v1, v2, v3, v4, v5, v6, v7, v8):
        try:
            fn()
        except Exception:                                      # noqa: BLE001
            global FAILED
            FAILED += 1
            import traceback
            print("  FAIL  %s raised" % fn.__name__)
            traceback.print_exc()
    print("")
    print("%d passed, %d failed" % (PASSED, FAILED))
    return 1 if FAILED else 0


if __name__ == "__main__":
    raise SystemExit(main())
