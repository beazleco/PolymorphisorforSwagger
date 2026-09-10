#!/usr/bin/env python3
"""
Regression suite for the Level-format workbook contract (v6.0).

Run from the package root::

    python -m tests.test_workbook_contract

Groups
======

W1   layout detection: the header row, the Level columns, the column roles
W2   the banner
W3   sections: only Request Body and Response Body are read
W4   the structure break: pasted payloads below the grid are ignored
W5   the Level tree: depth, parents, gaps, duplicates
W6   Data Type and Usage parsing
W7   SOR columns and pruning
W8   P1, one SOR endpoint
W9   P2, several SOR endpoints with a requestVariant discriminator
W10  P3, several SOR endpoints with no tag
W11  failure isolation: one bad sheet does not take the workbook down
W12  the empty-object guard
W13  the template validates clean and generates
W14  the reference credit card workbook, end to end
W15  method and path derivation
W16  the merged document: hoisting, cross-operation profiling, collisions
W17  drag and drop path parsing
W18  classifying the input format
W19  reading the field mapping format
W20  the Apigee error set
W21  class and context in schema names
W22  the API lifecycle status on every operation
W23  the field mapping export: shape, difference columns, provenance
W24  the round trip: export, re-import, and export again
"""

from __future__ import annotations

import os
import sys
import re
import json
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl                                                # noqa: E402

import polymorphize_generate as gen                            # noqa: E402
import polymorphize_merge as mrg                               # noqa: E402
import polymorphize_template as tmpl                           # noqa: E402
import polymorphize_validate as val                            # noqa: E402
import polymorphize_workbook as wbk                            # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
CREDIT_CARD = os.path.join(ROOT, "samples", "creditcard_v2.0.0.xlsx")

PASSED, FAILED = 0, 0
_GROUP = None


def group(title):
    global _GROUP
    _GROUP = title
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
# Building workbooks in memory
# --------------------------------------------------------------------------- #


def build(sheets, path):
    """``sheets`` is ``{title: list_of_rows}``, each row a list of cells."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title[:31])
        for row in rows:
            ws.append(list(row))
    wb.save(path)
    return path


def sheet_rows(levels, sor_headers, banner, sections, roles=None):
    """A well-formed operation sheet.

    ``sections`` is a list of ``(label, rows)`` where each row is
    ``(depth, name, example, dtype, usage, description, *sor)`` with depth
    1-based.
    """
    roles = roles or ["Dummy Value", "Data Type", "Usage", "Descriptions"]
    header = [""] + ["Level %d" % i for i in range(1, levels + 1)] + roles
    header += ["SOR %s" % s if s else "SOR Field" for s in sor_headers]
    rows = [header]
    width = len(header)

    def blank():
        return [""] * width

    for label, value, extra in banner:
        r = blank()
        r[0] = label
        r[1] = value
        if extra:
            r[2] = extra
        rows.append(r)
    for label, body in sections:
        r = blank()
        r[0] = label
        rows.append(r)
        for entry in body:
            depth, name, example, dtype, usage, desc = entry[:6]
            sor = list(entry[6:])
            r = blank()
            r[depth] = name
            base = levels + 1
            r[base + 0] = example
            r[base + 1] = dtype
            r[base + 2] = usage
            r[base + 3] = desc
            for j, v in enumerate(sor):
                if base + len(roles) + j < width:
                    r[base + len(roles) + j] = v
            rows.append(r)
    return rows


STD_BANNER = [
    ("Use Case:", "Retrieve a thing.", ""),
    ("Service Domain:", "Deposit Account", "BQ: Account Balance"),
    ("Equivalent BIAN API Endpoint:", "/DepositAccount/Balance/Retrieve", ""),
    ("Proposed Business API Endpoint:", "/deposit-account/balance/retrieve", ""),
    ("SOR API Endpoint:", "GET: /v1/balance", ""),
]

SIMPLE_REQ = [
    (1, "ThingRequest", "", "object", "", "Wrapper."),
    (2, "Identifier", "", "object", "", "Group."),
    (3, "idType", "IBAN", "String (12)", "", "Unmapped, so dropped."),
    (3, "idValue", "123", "String (34)", "Mandatory", "Mapped.", "sorId"),
]

SIMPLE_RES = [
    (1, "ThingResponse", "", "object", "", "Wrapper."),
    (2, "amount", "1.00", "Number(18)", "Mandatory", "Mapped.", "sorAmt"),
]


def one_sheet_workbook(path, **kw):
    rows = sheet_rows(
        kw.get("levels", 4),
        kw.get("sor_headers", ["/v1/balance"]),
        kw.get("banner", STD_BANNER),
        kw.get("sections", [("Request Body", SIMPLE_REQ),
                            ("Response Body", SIMPLE_RES)]),
        roles=kw.get("roles"))
    return build({kw.get("title", "Thing_Retrieve"): rows}, path)


def read_one(path):
    results = [r for r in wbk.read_workbook(path) if r.status != "skipped"]
    return results[0]


def tmp(name):
    return os.path.join(tempfile.mkdtemp(prefix="polymorphize6_"), name)


# --------------------------------------------------------------------------- #
# W1 layout
# --------------------------------------------------------------------------- #


def w1():
    group("W1 — layout detection")
    r = read_one(one_sheet_workbook(tmp("a.xlsx")))
    eq("the header row is row 1", r.layout.header_row, 0)
    eq("four Level columns are found", len(r.layout.level_cols), 4)
    eq("Level 1 is column B", r.layout.level_cols[0], 1)
    eq("the label column is column A", r.layout.label_cols, [0])
    check("every column role is recognised",
          set(r.layout.roles) >= {"type", "usage", "description", "example"},
          sorted(r.layout.roles))
    eq("one SOR column", len(r.layout.sor_cols), 1)
    eq("its endpoint comes from the header", r.layout.sor_cols[0].endpoint,
       "/v1/balance")

    # Synonyms an analyst actually uses.
    r2 = read_one(one_sheet_workbook(
        tmp("b.xlsx"), roles=["Example", "Data Type", "Required", "Description"]))
    check("Required is accepted for Usage", "usage" in r2.layout.roles)
    check("Example is accepted for Dummy Value", "example" in r2.layout.roles)
    check("Description is accepted for Descriptions",
          "description" in r2.layout.roles)

    # A sheet with no Level columns is a support sheet, not a failure.
    path = build({"Notes": [["Some notes"], ["More notes"]]}, tmp("c.xlsx"))
    res = wbk.read_workbook(path)
    eq("a sheet with no Level columns is skipped", res[0].status, "skipped")

    # Level numbering gaps are reported.
    rows = sheet_rows(3, ["/v1/x"], STD_BANNER,
                      [("Request Body", SIMPLE_REQ), ("Response Body", SIMPLE_RES)])
    rows[0][3] = "Level 4"        # Level 1, Level 2, Level 4
    r3 = read_one(build({"Gap_Retrieve": rows}, tmp("d.xlsx")))
    check("a gap in the Level numbering is reported",
          any(f.code == "L001" for f in r3.findings))


# --------------------------------------------------------------------------- #
# W2 banner
# --------------------------------------------------------------------------- #


def w2():
    group("W2 — the banner")
    r = read_one(one_sheet_workbook(tmp("a.xlsx")))
    eq("the use case is read", r.banner.get("use_case"), "Retrieve a thing.")
    eq("the service domain is read", r.banner.get("service_domain"),
       "Deposit Account")
    eq("the behaviour qualifier is read from the BQ cell",
       r.banner.get("behaviour_qualifier"), "Account Balance")
    eq("the BIAN endpoint is read", r.banner.get("bian_endpoint"),
       "/DepositAccount/Balance/Retrieve")
    eq("the SOR endpoint is read", r.banner.get("sor_endpoint"),
       "GET: /v1/balance")

    banner = [("Proposed Re-usable API Endpoint:", "/x/y", ""),
              ("Use Case:", "Something.", "")]
    r2 = read_one(one_sheet_workbook(tmp("b.xlsx"), banner=banner))
    eq("the hyphenated re-usable label is accepted",
       r2.banner.get("business_endpoint"), "/x/y")

    reps = val.validate_workbook(one_sheet_workbook(tmp("c.xlsx"), banner=[
        ("Use Case:", "", ""), ("Service Domain:", "", "")]), strict=True)
    check("strict reports a missing banner value",
          any(f.code == "B001" for f in reps[0].findings))


# --------------------------------------------------------------------------- #
# W3 sections
# --------------------------------------------------------------------------- #


def w3():
    group("W3 — only Request Body and Response Body are read")
    header_rows = [
        (1, "HeaderSchema", "", "object", "", "Should not be read."),
        (2, "correlationId", "abc", "String (36)", "Mandatory", "Ignored.",
         "corrId"),
    ]
    sections = [("Request Header", header_rows),
                ("Request Body", SIMPLE_REQ),
                ("ResponseHeader", header_rows),
                ("Response Body", SIMPLE_RES)]
    r = read_one(one_sheet_workbook(tmp("a.xlsx"), sections=sections))
    eq("the sheet reads clean", r.status, "ok")
    eq("the request root is the body schema", r.request.root.name, "ThingRequest")
    eq("the response root is the body schema", r.response.root.name,
       "ThingResponse")
    eq("both header sections are recorded as skipped",
       len(r.ignored_sections), 2)
    spec = gen.generate_sheet(r)
    names = set(spec.document["components"]["schemas"])
    check("no header schema reaches the output", "HeaderSchema" not in names,
          sorted(names))

    # Bare Request and Response are accepted as body aliases.
    r2 = read_one(one_sheet_workbook(
        tmp("b.xlsx"),
        sections=[("Request", SIMPLE_REQ), ("Response", SIMPLE_RES)]))
    eq("bare Request is read as Request Body", r2.request.root.name,
       "ThingRequest")
    eq("bare Response is read as Response Body", r2.response.root.name,
       "ThingResponse")

    # A parameter section is skipped too.
    r3 = read_one(one_sheet_workbook(
        tmp("c.xlsx"),
        sections=[("Request Parameter", header_rows),
                  ("Request Body", SIMPLE_REQ), ("Response Body", SIMPLE_RES)]))
    eq("a parameter section is skipped", len(r3.ignored_sections), 1)


# --------------------------------------------------------------------------- #
# W4 structure break
# --------------------------------------------------------------------------- #


def w4():
    group("W4 — content below the grid is ignored")
    rows = sheet_rows(4, ["/v1/balance"], STD_BANNER,
                      [("Request Body", SIMPLE_REQ),
                       ("Response Body", SIMPLE_RES)])
    width = len(rows[0])
    for _ in range(4):
        rows.append([""] * width)
    # A pasted payload sample. The braces land in Level columns, which is
    # exactly how the reference workbook breaks a naive reader.
    payload = [["/v1/account/cards", "", "{", "", "{"],
               ["{", "", '"accountCards": [', "", '"cards": ['],
               ['"cardToken": "1555"', "", "{", "", "{"],
               ["}", "", "]", "", "]"],
               ["}", "", "}", "", "}"]]
    for line in payload:
        rows.append(line + [""] * (width - len(line)))
    r = read_one(build({"Break_Retrieve": rows}, tmp("a.xlsx")))
    eq("the sheet still reads clean", r.status, "ok")
    check("rows below the grid are counted as ignored", r.ignored_rows >= 5,
          r.ignored_rows)

    def names(node):
        return [node.name] + [n for c in node.children for n in names(c)]
    all_names = names(r.request.root) + names(r.response.root)
    check("no brace becomes an attribute",
          not any(n in ("{", "}", "]", "[") for n in all_names), all_names)
    check("no payload key becomes an attribute",
          "accountCards" not in all_names and "cards" not in all_names,
          all_names)

    # Two blank rows inside a section must not end it. The reference workbook
    # has a genuine run of two.
    rows2 = sheet_rows(4, ["/v1/balance"], STD_BANNER,
                       [("Request Body", SIMPLE_REQ)])
    w2_ = len(rows2[0])
    rows2 += [[""] * w2_, [""] * w2_]
    rows2 += sheet_rows(4, ["/v1/balance"], [],
                        [("Response Body", SIMPLE_RES)])[1:]
    r2 = read_one(build({"TwoBlank_Retrieve": rows2}, tmp("b.xlsx")))
    check("two blank rows do not end the grid",
          r2.response is not None and r2.response.root is not None)
    eq("the break threshold is three", wbk.STRUCTURE_BREAK, 3)


# --------------------------------------------------------------------------- #
# W5 the Level tree
# --------------------------------------------------------------------------- #


def w5():
    group("W5 — the Level tree")
    deep = [
        (1, "DeepRequest", "", "object", "", "Root."),
        (2, "A", "", "object", "", "Level 2 group."),
        (3, "B", "", "object", "", "Level 3 group."),
        (4, "c", "1", "String (5)", "Mandatory", "Level 4 leaf.", "sorC"),
        (2, "D", "2", "String (5)", "", "Back to Level 2.", "sorD"),
    ]
    r = read_one(one_sheet_workbook(
        tmp("a.xlsx"), sections=[("Request Body", deep),
                                 ("Response Body", SIMPLE_RES)]))
    root = r.request.root
    eq("the root is the Level 1 row", root.name, "DeepRequest")
    eq("the root has two children", [c.name for c in root.children], ["A", "D"])
    a = root.children[0]
    eq("A holds B", [c.name for c in a.children], ["B"])
    eq("B holds c", [c.name for c in a.children[0].children], ["c"])
    eq("returning to Level 2 attaches to the root", root.children[1].name, "D")

    # A skipped Level is an error, and the row is still attached.
    gapped = [
        (1, "GapRequest", "", "object", "", "Root."),
        (3, "orphan", "1", "String (5)", "Mandatory", "Two Levels in.", "sorO"),
    ]
    r2 = read_one(one_sheet_workbook(
        tmp("b.xlsx"), sections=[("Request Body", gapped),
                                 ("Response Body", SIMPLE_RES)]))
    check("a skipped Level is reported",
          any(f.code == "A003" for f in r2.findings))
    finding = next(f for f in r2.findings if f.code == "A003")
    check("the message names the row and both Levels",
          "orphan" in finding.what and "Level" in finding.what, finding.what)
    check("the row is still attached rather than lost",
          any(c.name == "orphan" for c in r2.request.root.children))

    # Duplicate siblings.
    dup = [
        (1, "DupRequest", "", "object", "", "Root."),
        (2, "same", "1", "String (5)", "", "First.", "sorA"),
        (2, "same", "2", "String (5)", "", "Second.", "sorB"),
    ]
    r3 = read_one(one_sheet_workbook(
        tmp("c.xlsx"), sections=[("Request Body", dup),
                                 ("Response Body", SIMPLE_RES)]))
    check("a duplicate sibling is reported",
          any(f.code == "A007" for f in r3.findings))

    # A second Level 1 row.
    two_roots = [
        (1, "FirstRequest", "", "object", "", "Root."),
        (2, "x", "1", "String (5)", "", "Leaf.", "sorX"),
        (1, "SecondRequest", "", "object", "", "A second root."),
    ]
    r4 = read_one(one_sheet_workbook(
        tmp("d.xlsx"), sections=[("Request Body", two_roots),
                                 ("Response Body", SIMPLE_RES)]))
    check("a second Level 1 row is reported",
          any(f.code == "A004" for f in r4.findings))

    # Pasted junk inside a section is rejected by name, not consumed.
    junk = [
        (1, "JunkRequest", "", "object", "", "Root."),
        (2, "{", "", "", "", ""),
        (2, "ok", "1", "String (5)", "", "Leaf.", "sorOk"),
    ]
    r5 = read_one(one_sheet_workbook(
        tmp("e.xlsx"), sections=[("Request Body", junk),
                                 ("Response Body", SIMPLE_RES)]))
    check("an unusable name inside a section is reported",
          any(f.code == "A001" for f in r5.findings))
    eq("and is not attached", [c.name for c in r5.request.root.children], ["ok"])


# --------------------------------------------------------------------------- #
# W6 types and usage
# --------------------------------------------------------------------------- #


def w6():
    group("W6 — Data Type and Usage")
    rows = [
        (1, "TypesRequest", "", "object", "", "Root."),
        (2, "a", "x", "String (19)", "Mandatory", "Length in brackets.", "s1"),
        (2, "b", "1", "Number(20)", "", "No space before the bracket.", "s2"),
        (2, "c", "1", "integer", "", "Plain integer.", "s3"),
        (2, "d", "true", "boolean", "", "Boolean.", "s4"),
        (2, "e", "2026-01-01", "Date", "", "Date.", "s5"),
        (2, "f", "2026-01-01T00:00:00Z", "DateTime", "ConditionalMandatory",
         "Date and time.", "s6"),
        (2, "g", "1", "string(2)", "", "Lower case with a length.", "s7"),
    ]
    r = read_one(one_sheet_workbook(
        tmp("a.xlsx"), sections=[("Request Body", rows),
                                 ("Response Body", SIMPLE_RES)]))
    kids = {c.name: c for c in r.request.root.children}
    eq("String (19) is a string", kids["a"].json_type, "string")
    eq("its length becomes maxLength", kids["a"].max_length, 19)
    eq("Number(20) is a number", kids["b"].json_type, "number")
    eq("integer stays an integer", kids["c"].json_type, "integer")
    eq("boolean stays a boolean", kids["d"].json_type, "boolean")
    eq("Date is a string", kids["e"].json_type, "string")
    eq("with format date", kids["e"].fmt, "date")
    eq("DateTime has format date-time", kids["f"].fmt, "date-time")
    eq("string(2) parses without a space", kids["g"].max_length, 2)

    eq("Mandatory becomes required", kids["a"].usage, wbk.USAGE_REQUIRED)
    eq("ConditionalMandatory does not", kids["f"].usage, wbk.USAGE_CONDITIONAL)
    eq("a blank Usage is optional", kids["b"].usage, wbk.USAGE_OPTIONAL)

    spec = gen.generate_sheet(r)
    req = spec.document["components"]["schemas"]["TypesRequest"]
    eq("only Mandatory lands in required", req.get("required"), ["a"])
    check("ConditionalMandatory is annotated instead",
          req["properties"]["f"].get("x-conditional") is True)

    bad = [
        (1, "BadRequest", "", "object", "", "Root."),
        (2, "a", "x", "Charstring£", "Compulsory", "Both cells are wrong.", "s1"),
    ]
    r2 = read_one(one_sheet_workbook(
        tmp("b.xlsx"), sections=[("Request Body", bad),
                                 ("Response Body", SIMPLE_RES)]))
    check("an unrecognised Data Type is reported",
          any(f.code == "T001" for f in r2.findings))
    check("an unrecognised Usage is reported",
          any(f.code == "U001" for f in r2.findings))
    t = next(f for f in r2.findings if f.code == "T001")
    check("the type message quotes the cell contents",
          "Charstring" in t.what, t.what)
    check("and names the cell", t.cell.startswith("G"), t.cell)

    blank_leaf = [
        (1, "BlankRequest", "", "object", "", "Root."),
        (2, "a", "x", "", "", "No Data Type at all.", "s1"),
    ]
    r3 = read_one(one_sheet_workbook(
        tmp("c.xlsx"), sections=[("Request Body", blank_leaf),
                                 ("Response Body", SIMPLE_RES)]))
    check("a blank Data Type on a leaf is reported",
          any(f.code == "T002" for f in r3.findings))
    eq("and defaults to string",
       r3.request.root.children[0].json_type, "string")
    eq("a blank Data Type on a row with children becomes object",
       r3.request.root.json_type, "object")


# --------------------------------------------------------------------------- #
# W7 SOR columns and pruning
# --------------------------------------------------------------------------- #


def w7():
    group("W7 — SOR columns and pruning")
    r = read_one(one_sheet_workbook(tmp("a.xlsx")))
    spec = gen.generate_sheet(r)
    req = spec.document["components"]["schemas"]["ThingRequest"]
    ident = req["properties"]["Identifier"]["properties"]
    check("an unmapped attribute is removed", "idType" not in ident, sorted(ident))
    check("a mapped attribute is kept", "idValue" in ident, sorted(ident))
    eq("the SOR field name is carried through",
       ident["idValue"]["x-sor-field"], "sorId")

    # A container whose only mapped descendant is removed goes with it.
    rows = [
        (1, "PruneRequest", "", "object", "", "Root."),
        (2, "Empty", "", "object", "", "Nothing beneath is mapped."),
        (3, "x", "1", "String (5)", "", "Unmapped."),
        (2, "Kept", "", "object", "", "Something beneath is mapped."),
        (3, "y", "1", "String (5)", "", "Mapped.", "sorY"),
    ]
    r2 = read_one(one_sheet_workbook(
        tmp("b.xlsx"), sections=[("Request Body", rows),
                                 ("Response Body", SIMPLE_RES)]))
    props = gen.generate_sheet(r2).document["components"]["schemas"][
        "PruneRequest"]["properties"]
    check("a container with nothing mapped beneath it is removed",
          "Empty" not in props, sorted(props))
    check("a container with something mapped is kept", "Kept" in props)

    # Not-used markers are the same as an empty cell.
    for marker in ("N/A", "Not Used", "-", "TBC"):
        rows3 = [
            (1, "MarkRequest", "", "object", "", "Root."),
            (2, "a", "1", "String (5)", "", "Marked not used.", marker),
            (2, "b", "1", "String (5)", "", "Mapped.", "sorB"),
        ]
        r3 = read_one(one_sheet_workbook(
            tmp("m%s.xlsx" % abs(hash(marker))),
            sections=[("Request Body", rows3), ("Response Body", SIMPLE_RES)]))
        p = gen.generate_sheet(r3).document["components"]["schemas"][
            "MarkRequest"]["properties"]
        check("%r means not supplied" % marker, "a" not in p and "b" in p,
              sorted(p))

    # A sheet with no SOR column at all cannot be generated.
    rows4 = sheet_rows(4, [], STD_BANNER,
                       [("Request Body", SIMPLE_REQ),
                        ("Response Body", SIMPLE_RES)])
    del rows4[0][-1]
    r4 = read_one(build({"NoSor_Retrieve": rows4}, tmp("d.xlsx")))
    check("a sheet with no SOR column fails",
          r4.status == "failed" and any(f.code == "M001" for f in r4.findings))


# --------------------------------------------------------------------------- #
# W8 pattern P1
# --------------------------------------------------------------------------- #


def w8():
    group("W8 — P1, one SOR endpoint")
    r = read_one(one_sheet_workbook(tmp("a.xlsx")))
    spec = gen.generate_sheet(r)
    eq("the pattern is P1", spec.pattern, "P1")
    schemas = spec.document["components"]["schemas"]
    eq("two message schemas plus the shared error schema",
       sorted(n for n in schemas if n != "ErrorResponse"),
       ["ThingRequest", "ThingResponse"])
    check("the Apigee error schema is present", "ErrorResponse" in schemas)
    check("there is no oneOf",
          not any("oneOf" in s for s in schemas.values()))
    check("there is no allOf",
          not any("allOf" in s for s in schemas.values()))
    check("there is no discriminator",
          "discriminator" not in gen.dump_document(spec.document))


# --------------------------------------------------------------------------- #
# W9 pattern P2
# --------------------------------------------------------------------------- #

P2_REQ = [
    (1, "MultiRequest", "", "object", "", "Root."),
    (2, wbk.VARIANT_PROPERTY, "01", "String (2)", "Mandatory",
     "01. Alpha\n02. Beta"),
    (2, "common", "1", "String (5)", "Mandatory", "Both endpoints.",
     "sorCommon", "sorCommon"),
    (2, "onlyAlpha", "1", "String (5)", "", "Alpha only.", "sorAlpha", ""),
    (2, "onlyBeta", "1", "String (5)", "", "Beta only.", "", "sorBeta"),
]

P2_RES = [
    (1, "MultiResponse", "", "object", "", "Root."),
    (2, "shared", "1", "String (5)", "Mandatory", "Both endpoints.",
     "resCommon", "resCommon"),
    (2, "alphaOnly", "1", "String (5)", "", "Alpha only.", "resAlpha", ""),
]


def w9():
    group("W9 — P2, several SOR endpoints with a requestVariant")
    r = read_one(one_sheet_workbook(
        tmp("a.xlsx"), sor_headers=["/v1/alpha", "/v1/beta"],
        sections=[("Request Body", P2_REQ), ("Response Body", P2_RES)]))
    eq("two variants are declared", len(r.variants), 2)
    spec = gen.generate_sheet(r)
    eq("the pattern is P2", spec.pattern, "P2")
    s = spec.document["components"]["schemas"]

    eq("the base holds only the common attributes",
       sorted(s["MultiRequest__Base"]["properties"]),
       sorted([wbk.VARIANT_PROPERTY, "common"]))
    check("the base requires the discriminator property",
          wbk.VARIANT_PROPERTY in s["MultiRequest__Base"].get("required", []))

    for name, own, other in (("MultiRequest__Alpha", "onlyAlpha", "onlyBeta"),
                             ("MultiRequest__Beta", "onlyBeta", "onlyAlpha")):
        check("%s exists" % name, name in s, sorted(s))
        entry = s[name]
        eq("%s inherits through allOf" % name,
           entry["allOf"][0]["$ref"],
           "#/components/schemas/MultiRequest__Base")
        delta = entry["allOf"][1]["properties"]
        check("%s declares %s" % (name, own), own in delta, sorted(delta))
        check("%s does not declare %s" % (name, other), other not in delta,
              sorted(delta))
        check("%s does not repeat the common attribute" % name,
              "common" not in delta, sorted(delta))
        eq("%s narrows the tag to one code" % name,
           delta[wbk.VARIANT_PROPERTY]["enum"],
           ["01"] if own == "onlyAlpha" else ["02"])

    wrapper = s["MultiRequest"]
    eq("the request wrapper is a oneOf of two", len(wrapper["oneOf"]), 2)
    eq("with a discriminator on the tag",
       wrapper["discriminator"]["propertyName"], wbk.VARIANT_PROPERTY)
    eq("mapped by variant code",
       sorted(wrapper["discriminator"]["mapping"]), ["01", "02"])
    eq("01 maps to the alpha schema",
       wrapper["discriminator"]["mapping"]["01"],
       "#/components/schemas/MultiRequest__Alpha")

    res = s["MultiResponse"]
    eq("the response is also a oneOf", len(res["oneOf"]), 2)
    check("but carries no discriminator", "discriminator" not in res)
    eq("the dependency on the request tag is recorded",
       res.get("x-selected-by"), wbk.VARIANT_PROPERTY)

    check("the endpoint is recorded on each variant",
          s["MultiRequest__Alpha"]["x-sor-endpoint"] == "/v1/alpha")

    # The tag is never dropped by pruning even though no SOR column names it.
    check("the tag survives pruning despite having no SOR field",
          wbk.VARIANT_PROPERTY in s["MultiRequest__Base"]["properties"])

    # Mismatched counts are reported.
    bad_req = list(P2_REQ)
    bad_req[1] = (2, wbk.VARIANT_PROPERTY, "01", "String (2)", "Mandatory",
                  "01. Alpha\n02. Beta\n03. Gamma")
    r2 = read_one(one_sheet_workbook(
        tmp("b.xlsx"), sor_headers=["/v1/alpha", "/v1/beta"],
        sections=[("Request Body", bad_req), ("Response Body", P2_RES)]))
    check("three variants against two SOR columns is reported",
          any(f.code == "V002" for f in r2.findings))

    # A variant whose name matches nothing in its column is reported.
    odd = list(P2_REQ)
    odd[1] = (2, wbk.VARIANT_PROPERTY, "01", "String (2)", "Mandatory",
              "01. Zebra\n02. Walrus")
    r3 = read_one(one_sheet_workbook(
        tmp("c.xlsx"), sor_headers=["/v1/alpha", "/v1/beta"],
        sections=[("Request Body", odd), ("Response Body", P2_RES)]))
    s3 = gen.generate_sheet(r3)
    check("a variant paired only by position is reported",
          any(f.code == "V003" for f in s3.findings))


# --------------------------------------------------------------------------- #
# W10 pattern P3
# --------------------------------------------------------------------------- #


def w10():
    group("W10 — P3, several SOR endpoints with no tag")
    req = [row for row in P2_REQ if row[1] != wbk.VARIANT_PROPERTY]
    r = read_one(one_sheet_workbook(
        tmp("a.xlsx"), sor_headers=["/v1/alpha", "/v1/beta"],
        sections=[("Request Body", req), ("Response Body", P2_RES)]))
    check("the missing tag is reported as a warning",
          any(f.code == "V001" for f in r.findings))
    eq("but the sheet still reads", r.status, "ok")
    spec = gen.generate_sheet(r)
    eq("the pattern is P3", spec.pattern, "P3")
    s = spec.document["components"]["schemas"]
    wrapper = s["MultiRequest"]
    eq("the wrapper is a oneOf", len(wrapper["oneOf"]), 2)
    check("with no discriminator", "discriminator" not in wrapper)
    check("and says why", "identifier" in wrapper.get("description", ""),
          wrapper.get("description"))
    eq("the variants still inherit through allOf",
       s["MultiRequest__Alpha"]["allOf"][0]["$ref"],
       "#/components/schemas/MultiRequest__Base")


# --------------------------------------------------------------------------- #
# W11 failure isolation
# --------------------------------------------------------------------------- #


def w11():
    group("W11 — one bad sheet does not take the workbook down")
    good = sheet_rows(4, ["/v1/balance"], STD_BANNER,
                      [("Request Body", SIMPLE_REQ),
                       ("Response Body", SIMPLE_RES)])
    # The reference workbook's own defect: the request schema sits under a
    # section labelled Request Header, and there is no Request Body.
    mislabelled = sheet_rows(4, ["/v1/balance"], STD_BANNER,
                             [("Request Header", SIMPLE_REQ),
                              ("Response Body", SIMPLE_RES)])
    path = build({"Good_Retrieve": good,
                  "Bad_Retrieve": mislabelled,
                  "Good2_Retrieve": good}, tmp("a.xlsx"))

    results = wbk.read_workbook(path)
    statuses = {r.sheet: r.status for r in results}
    eq("the first good sheet reads", statuses["Good_Retrieve"], "ok")
    eq("the mislabelled sheet fails", statuses["Bad_Retrieve"], "failed")
    eq("the sheet after it still reads", statuses["Good2_Retrieve"], "ok")

    bad = next(r for r in results if r.sheet == "Bad_Retrieve")
    f = next(f for f in bad.findings if f.code == "S001")
    check("the message names the mislabelled section",
          "Request Header" in f.what, f.what)
    check("and gives the cell to change", f.cell == "A7", f.cell)
    check("and says exactly what to change it to",
          "Request Body" in f.fix, f.fix)

    out_dir = os.path.join(os.path.dirname(path), "out")
    out = gen.run(path, out_dir, split=True, log=lambda *_a: None)
    eq("two endpoints generate", len(out.ok), 2)
    eq("one fails", len(out.failed), 1)
    check("the good specifications are on disk",
          os.path.exists(os.path.join(out_dir, "good-retrieve.yaml")) and
          os.path.exists(os.path.join(out_dir, "good2-retrieve.yaml")))
    check("no specification is written for the failure",
          not os.path.exists(os.path.join(out_dir, "bad-retrieve.yaml")))
    check("an explanation is written instead",
          os.path.exists(os.path.join(out_dir, "bad-retrieve.FAILED.md")))
    report = open(os.path.join(out_dir, "generation_report.md"),
                  encoding="utf-8").read()
    check("the report leads with the failure",
          report.index("Failed endpoints") < report.index("Endpoints generated"))
    check("and names the sheet", "Bad_Retrieve" in report)

    # A stale specification from an earlier run must not survive.
    stale = os.path.join(out_dir, "bad-retrieve.yaml")
    with open(stale, "w", encoding="utf-8") as fh:
        fh.write("openapi: 3.0.3\n")
    gen.run(path, out_dir, split=True, log=lambda *_a: None)
    check("a stale specification is removed on the next run",
          not os.path.exists(stale))


# --------------------------------------------------------------------------- #
# W12 the empty-object guard
# --------------------------------------------------------------------------- #


def w12():
    group("W12 — nothing empty ever reaches the output")
    unmapped_response = [
        (1, "EmptyResponse", "", "object", "", "Root."),
        (2, "a", "1", "String (5)", "", "Nobody filled in the SOR column."),
        (2, "b", "1", "String (5)", "", "Nor this one."),
    ]
    r = read_one(one_sheet_workbook(
        tmp("a.xlsx"), sections=[("Request Body", SIMPLE_REQ),
                                 ("Response Body", unmapped_response)]))
    spec = gen.generate_sheet(r)
    eq("an unmapped response no longer fails the endpoint", spec.status, "ok")
    f = next(f for f in spec.findings if f.code == "E001")
    eq("it is a warning", f.severity, "warning")
    check("the message names the section", "Response Body" in f.what, f.what)
    check("and counts the rows examined", "2 attributes" in f.what, f.what)
    check("and says no payload was published",
          "no payload was published" in f.what, f.what)
    responses = spec.document["paths"][spec.path][spec.method]["responses"]
    eq("the response is 204 No Content rather than an empty object",
       [c for c in sorted(responses) if c.startswith("2")], ["204"])
    eq("and the Apigee error set is attached alongside it",
       [c for c in sorted(responses) if not c.startswith("2")],
       ["400", "401", "403", "404", "405", "409", "422", "429",
        "500", "502", "503", "504"])
    check("and nothing empty was published",
          gen.empty_schemas(spec.document) == [])

    # An entirely empty section, label present, is permitted in silence.
    rows = sheet_rows(4, ["/v1/balance"], STD_BANNER,
                      [("Request Body", []), ("Response Body", SIMPLE_RES)])
    r2 = read_one(build({"NoReq_Retrieve": rows}, tmp("b.xlsx")))
    check("an empty Request Body section is permitted",
          r2.status == "ok" and any(f.code == "S004" for f in r2.findings))
    eq("and it is a warning, not an error",
       [f.severity for f in r2.findings if f.code == "S004"], ["warning"])
    spec2 = gen.generate_sheet(r2)
    eq("the endpoint still generates", spec2.status, "ok")
    operation = spec2.document["paths"][spec2.path][spec2.method]
    check("with no requestBody at all", "requestBody" not in operation,
          sorted(operation))
    eq("and the natural verb is kept, since there is no body to carry",
       spec2.method, "get")

    # A missing label is still a hard failure.
    rows3 = sheet_rows(4, ["/v1/balance"], STD_BANNER,
                       [("Response Body", SIMPLE_RES)])
    r3 = read_one(build({"Gone_Retrieve": rows3}, tmp("c.xlsx")))
    eq("a missing Request Body label still fails the sheet", r3.status, "failed")
    check("with S001", any(f.code == "S001" for f in r3.errors))

    check("the guard finds an empty object in any document",
          gen.empty_schemas({"components": {"schemas": {
              "A": {"type": "object", "properties": {}}}}}) == ["A"])
    check("a nested empty object is found too",
          gen.empty_schemas({"components": {"schemas": {
              "A": {"type": "object",
                    "properties": {"b": {"type": "object",
                                         "properties": {}}}}}}}) == ["A.b"])
    check("a composed schema is not empty",
          gen.empty_schemas({"components": {"schemas": {
              "A": {"allOf": [{"$ref": "#/x"}]}}}}) == [])
    check("a scalar is not empty",
          gen.empty_schemas({"components": {"schemas": {
              "A": {"type": "string"}}}}) == [])

    # And the guard holds across the reference workbook.
    if os.path.exists(CREDIT_CARD):
        out = gen.run(CREDIT_CARD, tmp("cc_out"), log=lambda *_a: None)
        offenders = [s.sheet for s in out.ok if gen.empty_schemas(s.document)]
        eq("no generated specification carries an empty object", offenders, [])


# --------------------------------------------------------------------------- #
# W13 the template
# --------------------------------------------------------------------------- #


def w13():
    group("W13 — the template is the contract")
    path = tmpl.write_template(tmp("SOR_mapping_template.xlsx"))
    wb = openpyxl.load_workbook(path)
    titles = wb.sheetnames
    for expected in ("How to use", "Vocabulary", "Example_SingleSor",
                     "Example_MultiSor", "Operation_Template"):
        check("the template has a %r sheet" % expected, expected in titles,
              titles)

    ws = wb["Operation_Template"]
    eq("the header is row 1", ws.cell(row=1, column=2).value, "Level 1")
    eq("six Level columns", ws.cell(row=1, column=7).value, "Level 6")
    eq("the banner starts on row 2", ws.cell(row=2, column=1).value, "Use Case:")
    eq("the SOR column is headed with the endpoint",
       ws.cell(row=1, column=12).value, "SOR /v1/replace/me")
    check("Request Body appears in the label column",
          any(ws.cell(row=r, column=1).value == "Request Body"
              for r in range(7, 12)))
    check("dropdowns are attached", len(ws.data_validations.dataValidation) >= 2)
    wb.close()

    for strict in (False, True):
        reps = val.validate_workbook(path, strict=strict)
        s = val.summarise(reps)
        eq("the template validates clean, strict=%s: no errors" % strict,
           s["errors"], 0)
        eq("the template validates clean, strict=%s: no warnings" % strict,
           s["warnings"], 0)
        eq("three operation sheets, strict=%s" % strict, s["operations"], 3)
        eq("two support sheets, strict=%s" % strict, s["skipped"], 2)

    out = gen.run(path, tmp("tmpl_out"), log=lambda *_a: None)
    eq("every template sheet generates", len(out.failed), 0)
    patterns = sorted(s.pattern for s in out.ok)
    eq("the examples cover P1 and P2", patterns, ["P1", "P1", "P2"])

    # The template must also merge cleanly, and demonstrate the reuse its own
    # warnings ask for rather than tripping them.
    eq("the template merges with no findings at all",
       [f.code for f in out.merged.findings], [])
    merged = out.merged.document
    eq("the three examples land in one document", len(merged["paths"]), 3)
    ms = merged["components"]["schemas"]
    eq("a group declared identically on two sheets is published once",
       [n for n in ms if n.startswith("AccountIdentifier")],
       ["AccountIdentifier"])
    eq("and records both operations that use it",
       sorted(ms["AccountIdentifier"].get("x-used-by") or []),
       ["exampleMultiSor", "exampleSingleSor"])
    eq("so nothing needed specialising across operations",
       out.merged.stats["split_groups"], 0)
    check("and a Method cell keeps a body-carrying Retrieve off GET",
          not any(f.code == "G001" for r in out.ok for f in r.findings))


# --------------------------------------------------------------------------- #
# W14 the reference workbook
# --------------------------------------------------------------------------- #


def w14():
    group("W14 — the reference credit card workbook")
    if not os.path.exists(CREDIT_CARD):
        check("the reference workbook is present", False, CREDIT_CARD)
        return
    results = wbk.read_workbook(CREDIT_CARD)
    eq("thirteen sheets", len(results), 13)
    eq("one support sheet is skipped",
       sum(1 for r in results if r.status == "skipped"), 1)
    skipped = next(r for r in results if r.status == "skipped")
    eq("and it is the payload sheet", skipped.sheet, "SOR Payload")

    out = gen.run(CREDIT_CARD, tmp("cc"), log=lambda *_a: None)
    eq("eleven endpoints generate", len(out.ok), 11)
    eq("one fails", len(out.failed), 1)
    failed = sorted(s.sheet for s in out.failed)
    eq("and it is the mislabelled sheet", failed,
       ["Card Issued Device_Retrieve"])

    codes = {s.sheet: next(f.code for f in s.findings if f.severity == "error")
             for s in out.failed}
    eq("the mislabelled section is S001",
       codes["Card Issued Device_Retrieve"], "S001")
    rsa = next(s for s in out.ok if s.sheet == "Card RSAEncrypted")
    check("the unmapped response is now a warning on a generated endpoint",
          any(f.code == "E001" and f.severity == "warning" for f in rsa.findings))
    eq("and that endpoint returns 204",
       [c for c in sorted(rsa.document["paths"][rsa.path][rsa.method]["responses"])
        if c.startswith("2")], ["204"])

    patterns = sorted(s.pattern for s in out.ok)
    eq("two endpoints are variant driven",
       patterns.count("P2"), 2)
    eq("the rest have a single SOR endpoint", patterns.count("P1"), 9)

    details = next(s for s in out.ok if s.sheet == "Card Details_Retrieve")
    s = details.document["components"]["schemas"]
    eq("Card Details has five request variants",
       sum(1 for n in s
           if n.startswith("CreditCardAccountRetrieveRequest__")
           and not n.endswith("__Base")), 5)
    disc = s["CreditCardAccountRetrieveRequest"]["discriminator"]
    eq("its discriminator is on requestVariant", disc["propertyName"],
       wbk.VARIANT_PROPERTY)
    eq("with five codes mapped", len(disc["mapping"]), 5)

    # The reduction that justifies the exercise. In split mode the response
    # base exists per sheet; where the five variant responses share nothing at
    # all there is no base, and that absence is correct rather than an
    # omission, because an empty base would be an empty object.
    base_name = "CreditCardAccountRetrieveResponse__Base"
    if base_name in s:
        base = s[base_name]["properties"]
        check("the response base is much smaller than the union",
              len(base) <= 2, len(base))
    else:
        shapes = [n for n in s
                  if n.startswith("CreditCardAccountRetrieveResponse__")]
        common = (set.intersection(*[set(s[n].get("properties") or {})
                                     for n in shapes]) if shapes else set())
        check("no response base, and the variants genuinely share nothing",
              shapes and not common, sorted(common))

    for spec in out.ok:
        eq("%s carries no empty object" % spec.sheet,
           gen.empty_schemas(spec.document), [])
        eq("%s produces exactly one path" % spec.sheet,
           len(spec.document["paths"]), 1)


# --------------------------------------------------------------------------- #
# W15 method and path
# --------------------------------------------------------------------------- #


def w15():
    group("W15 — method and path derivation")
    eq("Initiate is a POST",
       gen.derive_method("Card_Initiate", {}), "post")
    eq("Update is a PUT", gen.derive_method("Card_Update", {}), "put")
    eq("Control is a PUT", gen.derive_method("Card_Control", {}), "put")
    eq("Retrieve is a GET before the body check",
       gen.derive_method("Card_Retrieve", {}), "get")
    eq("the BIAN endpoint wins over the sheet name",
       gen.derive_method("Card_Retrieve",
                         {"bian_endpoint": "/CreditCard/IssuedDevice/Update"}),
       "put")
    eq("an explicit Method banner cell wins outright",
       gen.derive_method("Card_Retrieve", {"method": "PATCH"}), "patch")

    # The SOR endpoint's own verb must never be borrowed: the reference
    # workbook fulfils an Initiate with PUT and an Update with DELETE.
    eq("the SOR verb is ignored for an Initiate",
       gen.derive_method("Card_Initiate",
                         {"sor_endpoint": "PUT: /v1/card/activate"}), "post")
    eq("the SOR verb is ignored for an Update",
       gen.derive_method("Card_Update",
                         {"sor_endpoint": "DELETE: /v1/card/cancel"}), "put")

    eq("the business endpoint is preferred for the path",
       gen.derive_path("X_Retrieve", {"business_endpoint": "/a/b",
                                      "bian_endpoint": "/c/d"}), "/a/b")
    eq("N/A falls through to the BIAN endpoint",
       gen.derive_path("X_Retrieve", {"business_endpoint": "N/A",
                                      "bian_endpoint": "/c/d"}), "/c/d")
    eq("with neither, one is built from the service domain",
       gen.derive_path("Thing_Retrieve", {"service_domain": "Deposit Account"}),
       "/deposit-account/thing-retrieve")

    # A GET carrying a body is emitted as POST, loudly.
    r = read_one(one_sheet_workbook(tmp("a.xlsx"), title="Thing_Retrieve",
                                    banner=[("Use Case:", "x", "")]))
    spec = gen.generate_sheet(r)
    eq("a GET with a request body becomes a POST", spec.method, "post")
    check("and the substitution is reported",
          any(f.code == "G001" for f in spec.findings))




# --------------------------------------------------------------------------- #
# W16 the merged document
# --------------------------------------------------------------------------- #


def _merge(path, **kw):
    results = wbk.read_workbook(path)
    specs = [gen.generate_sheet(r, **kw) for r in results]
    ok = [(r, s) for r, s in zip(results, specs) if s.status == "ok"]
    return mrg.merge([r for r, _s in ok], [s for _r, s in ok],
                     log=lambda *_a: None)


def w16():
    group("W16 — the merged document")

    # Two operations sharing a group name with different mapped subsets. This
    # is the case the merge exists to resolve, and it cannot arise while each
    # sheet has its own namespace.
    shared_a = [
        (1, "AlphaRequest", "", "object", "", "Root."),
        (2, "Party", "", "object", "", "Shared group."),
        (3, "partyId", "1", "String (12)", "Mandatory", "In both.", "sorId"),
        (3, "alphaOnly", "1", "String (5)", "", "Alpha only.", "sorAlpha"),
    ]
    shared_b = [
        (1, "BetaRequest", "", "object", "", "Root."),
        (2, "Party", "", "object", "", "Shared group."),
        (3, "partyId", "1", "String (12)", "Mandatory", "In both.", "sorId"),
        (3, "betaOnly", "1", "String (5)", "", "Beta only.", "sorBeta"),
    ]
    res_a = [(1, "AlphaResponse", "", "object", "", "Root."),
             (2, "ok", "1", "String (5)", "", "Mapped.", "sorOk")]
    res_b = [(1, "BetaResponse", "", "object", "", "Root."),
             (2, "ok", "1", "String (5)", "", "Mapped.", "sorOk")]
    banner_a = list(STD_BANNER)
    banner_a[3] = ("Proposed Business API Endpoint:", "/alpha/retrieve", "")
    banner_b = list(STD_BANNER)
    banner_b[3] = ("Proposed Business API Endpoint:", "/beta/retrieve", "")
    path = build({
        "Alpha_Retrieve": sheet_rows(4, ["/v1/a"], banner_a,
                                     [("Request Body", shared_a),
                                      ("Response Body", res_a)]),
        "Beta_Retrieve": sheet_rows(4, ["/v1/b"], banner_b,
                                    [("Request Body", shared_b),
                                     ("Response Body", res_b)]),
    }, tmp("a.xlsx"))

    m = _merge(path)
    eq("the merge produced a document", m.errors, [])
    d = m.document
    eq("one document holds both operations", len(d["paths"]), 2)
    s = d["components"]["schemas"]

    # Since 6.4 the shared core takes <Name>Base and the most-used shape keeps
    # the plain name, because the plain name is what a consumer meets in
    # generated code and in rendered documentation. Reserving it for a base
    # that no property references left every visible name suffixed.
    check("the shared core is Party__Base", "Party__Base" in s, sorted(s))
    eq("and holds only what both operations publish",
       sorted(s["Party__Base"]["properties"]), ["partyId"])
    eq("its class and context are machine readable",
       (s["Party__Base"]["x-class"], s["Party__Base"]["x-context"],
        s["Party__Base"]["x-qualified-name"]),
       ("Party", "Base", "Party~Base"))
    check("the plain name is a real, usable shape", "Party" in s, sorted(s))
    check("and is a subtype of the shared core",
          s["Party"]["allOf"][0]["$ref"] == "#/components/schemas/Party__Base",
          s["Party"].get("allOf"))

    shapes = [n for n in s if n == "Party" or
              (n.startswith("Party" + "__") and not n.endswith("__Base"))]
    eq("two shapes are published, one of them plainly named", len(shapes), 2)
    suffixed = [n for n in shapes if n != "Party"]
    eq("exactly one carries a suffix", len(suffixed), 1)
    check("and it names the operation it serves",
          suffixed[0] in ("Party__AlphaRetrieve", "Party__BetaRetrieve"),
          suffixed)

    for name in shapes:
        eq("%s inherits from the shared core" % name,
           s[name]["allOf"][0]["$ref"], "#/components/schemas/Party__Base")
        delta = s[name]["allOf"][1]["properties"]
        check("%s carries exactly one of the two differing attributes" % name,
              sorted(delta) in (["alphaOnly"], ["betaOnly"]), sorted(delta))
        check("%s does not repeat the shared attribute" % name,
              "partyId" not in delta, sorted(delta))

    check("no discriminator is introduced for the operation axis",
          not [n for n, v in s.items() if "discriminator" in v],
          [n for n, v in s.items() if "discriminator" in v])

    # The point of the change: each request still points at its own shape, and
    # one of the two now reads as the plain concept name.
    pointed = sorted([s["AlphaRequest"]["properties"]["Party"]["$ref"],
                      s["BetaRequest"]["properties"]["Party"]["$ref"]])
    eq("the two requests point at the two shapes", pointed,
       sorted(["#/components/schemas/Party",
               "#/components/schemas/" + suffixed[0]]))
    check("neither points at the shared core directly",
          "#/components/schemas/Party__Base" not in pointed, pointed)
    # And whichever schema they point at, the property key is unchanged, which
    # is what keeps the wire identical.
    eq("the property key is the concept name on both",
       [list(s["AlphaRequest"]["properties"]).count("Party"),
        list(s["BetaRequest"]["properties"]).count("Party")], [1, 1])

    # A group that agrees everywhere is published once.
    same = [
        (1, "GammaRequest", "", "object", "", "Root."),
        (2, "Party", "", "object", "", "Shared group."),
        (3, "partyId", "1", "String (12)", "Mandatory", "In both.", "sorId"),
    ]
    banner_c = list(STD_BANNER)
    banner_c[3] = ("Proposed Business API Endpoint:", "/gamma/retrieve", "")
    path2 = build({
        "Alpha_Retrieve": sheet_rows(4, ["/v1/a"], banner_a,
                                     [("Request Body", same),
                                      ("Response Body", res_a)]),
        "Gamma_Retrieve": sheet_rows(4, ["/v1/c"], banner_c,
                                     [("Request Body", same),
                                      ("Response Body", res_b)]),
    }, tmp("b.xlsx"))
    m2 = _merge(path2)
    s2 = m2.document["components"]["schemas"]
    eq("a group that agrees is published once",
       [n for n in s2 if n.startswith("Party")], ["Party"])
    eq("and records both operations that use it",
       sorted(s2["Party"].get("x-used-by") or []),
       ["alphaRetrieve", "gammaRetrieve"])

    # An array never becomes a component: its item does.
    arr = [
        (1, "ListRequest", "", "object", "", "Root."),
        (2, "Lines", "", "Array", "", "Repeating group.", "sorLines"),
        (3, "lineId", "1", "String (5)", "Mandatory", "Mapped.", "sorLineId"),
        (3, "lineText", "x", "String (40)", "", "Mapped.", "sorLineText"),
    ]
    path3 = build({"List_Retrieve": sheet_rows(
        4, ["/v1/l"], STD_BANNER,
        [("Request Body", arr), ("Response Body", res_a)])}, tmp("c.xlsx"))
    m3 = _merge(path3)
    s3 = m3.document["components"]["schemas"]
    eq("no component is an array",
       [n for n in s3 if s3[n].get("type") == "array"], [])
    check("the array item is hoisted under its own name",
          "LinesItem" in s3, sorted(s3))
    eq("and the property is an array of references",
       s3["ListRequest"]["properties"]["Lines"]["items"]["$ref"],
       "#/components/schemas/LinesItem")

    # An Array row with a single object row beneath it means that row.
    arr2 = [
        (1, "List2Request", "", "object", "", "Root."),
        (2, "Lines", "", "Array", "", "Repeating group."),
        (3, "Line", "", "object", "", "The element."),
        (4, "lineId", "1", "String (5)", "Mandatory", "Mapped.", "sorLineId"),
    ]
    path4 = build({"List2_Retrieve": sheet_rows(
        4, ["/v1/l"], STD_BANNER,
        [("Request Body", arr2), ("Response Body", res_a)])}, tmp("d.xlsx"))
    s4 = _merge(path4).document["components"]["schemas"]
    eq("a single object row beneath an Array is the item",
       s4["List2Request"]["properties"]["Lines"]["items"]["$ref"],
       "#/components/schemas/Line")
    check("and no synthetic item name is invented", "LinesItem" not in s4,
          sorted(s4))

    # Documentation must not fragment the namespace.
    doc_a = [
        (1, "DocARequest", "", "object", "", "Root."),
        (2, "Party", "", "object", "", "Group."),
        (3, "partyId", "1", "String (12)", "Mandatory", "One wording.", "sorX"),
    ]
    doc_b = [
        (1, "DocBRequest", "", "object", "", "Root."),
        (2, "Party", "", "object", "", "Group."),
        (3, "partyId", "2", "String (12)", "Mandatory",
         "Quite another wording entirely.", "sorY"),
    ]
    path5 = build({
        "DocA_Retrieve": sheet_rows(4, ["/v1/a"], banner_a,
                                    [("Request Body", doc_a),
                                     ("Response Body", res_a)]),
        "DocB_Retrieve": sheet_rows(4, ["/v1/b"], banner_b,
                                    [("Request Body", doc_b),
                                     ("Response Body", res_b)]),
    }, tmp("e.xlsx"))
    s5 = _merge(path5).document["components"]["schemas"]
    eq("differing prose does not split a group",
       [n for n in s5 if n.startswith("Party")], ["Party"])
    field = s5["Party"]["properties"]["partyId"]
    check("an SOR field that differs is recorded per operation",
          "x-sor-field-by-operation" in field, sorted(field))
    eq("naming both operations",
       sorted(field["x-sor-field-by-operation"]),
       ["docARetrieve", "docBRetrieve"])

    # A differing contract does split it, and says why.
    clash_b = [
        (1, "DocBRequest", "", "object", "", "Root."),
        (2, "Party", "", "object", "", "Group."),
        (3, "partyId", "2", "String (40)", "", "Longer and optional.", "sorY"),
    ]
    path6 = build({
        "DocA_Retrieve": sheet_rows(4, ["/v1/a"], banner_a,
                                    [("Request Body", doc_a),
                                     ("Response Body", res_a)]),
        "DocB_Retrieve": sheet_rows(4, ["/v1/b"], banner_b,
                                    [("Request Body", clash_b),
                                     ("Response Body", res_b)]),
    }, tmp("f.xlsx"))
    m6 = _merge(path6)
    check("a differing declaration is reported as P002",
          any(f.code == "P002" for f in m6.findings))
    p002 = next(f for f in m6.findings if f.code == "P002")
    check("naming the attribute", "Party.partyId" in p002.what, p002.what)
    check("and both declarations",
          "string (12) mandatory" in p002.what and
          "string (40) optional" in p002.what, p002.what)
    check("and the shared attribute means no base is possible, reported as P001",
          any(f.code == "P001" for f in m6.findings))

    # Names that differ only by case.
    case_b = [
        (1, "DocBRequest", "", "object", "", "Root."),
        (2, "party", "", "object", "", "Same concept, lower case."),
        (3, "partyId", "2", "String (12)", "Mandatory", "Mapped.", "sorY"),
    ]
    path7 = build({
        "DocA_Retrieve": sheet_rows(4, ["/v1/a"], banner_a,
                                    [("Request Body", doc_a),
                                     ("Response Body", res_a)]),
        "DocB_Retrieve": sheet_rows(4, ["/v1/b"], banner_b,
                                    [("Request Body", case_b),
                                     ("Response Body", res_b)]),
    }, tmp("g.xlsx"))
    m7 = _merge(path7)
    check("names differing only by case are reported as N001",
          any(f.code == "N001" for f in m7.findings))
    n001 = next(f for f in m7.findings if f.code == "N001")
    check("naming both spellings",
          "'Party'" in n001.what and "'party'" in n001.what, n001.what)
    check("and saying the tool will not rename them",
          "would change the published field names" in n001.fix, n001.fix)
    s7 = m7.document["components"]["schemas"]
    check("both spellings survive as separate components",
          "Party" in s7 and "party" in s7, sorted(s7))

    # A path claimed twice fails the merge.
    path8 = build({
        "Alpha_Retrieve": sheet_rows(4, ["/v1/a"], banner_a,
                                     [("Request Body", shared_a),
                                      ("Response Body", res_a)]),
        "Clash_Retrieve": sheet_rows(4, ["/v1/b"], banner_a,
                                     [("Request Body", shared_b),
                                      ("Response Body", res_b)]),
    }, tmp("h.xlsx"))
    m8 = _merge(path8)
    check("two sheets claiming one path and method is an error",
          any(f.code == "Z003" for f in m8.findings))

    # A message schema name used twice fails the merge.
    dup_root = [
        (1, "AlphaRequest", "", "object", "", "The same root name again."),
        (2, "x", "1", "String (5)", "Mandatory", "Mapped.", "sorX"),
    ]
    path9 = build({
        "Alpha_Retrieve": sheet_rows(4, ["/v1/a"], banner_a,
                                     [("Request Body", shared_a),
                                      ("Response Body", res_a)]),
        "Other_Retrieve": sheet_rows(4, ["/v1/b"], banner_b,
                                     [("Request Body", dup_root),
                                      ("Response Body", res_b)]),
    }, tmp("i.xlsx"))
    m9 = _merge(path9)
    check("a duplicated message schema name is an error",
          any(f.code == "Z002" for f in m9.findings))

    # The reference workbook, merged.
    if os.path.exists(CREDIT_CARD):
        m10 = _merge(CREDIT_CARD)
        eq("the reference workbook merges cleanly", m10.errors, [])
        d10 = m10.document
        eq("eleven operations in one document", len(d10["paths"]), 11)
        eq("one tag, because the domain is folded to one spelling",
           len(d10["tags"]), 1)
        check("the three spellings of the domain are reported as N002",
              any(f.code == "N002" for f in m10.findings))
        eq("no schema is empty", gen.empty_schemas(d10), [])

        # 6.4: the plain concept name must belong to a shape a property points
        # at, not to an intersection base that nothing references. A generator
        # names its classes after component schemas, so this is the difference
        # between a consumer meeting Account or AccountForCardDetailsRetrieve3.
        import re as _re
        _suf = _re.compile(r"(For[A-Z]\w*?\d?|Profile\d+|Base)$")
        _refs = []
        for _owner, _sch in d10["components"]["schemas"].items():
            for _prop, _v in (_sch.get("properties") or {}).items():
                if not isinstance(_v, dict):
                    continue
                _t = _v.get("$ref") or (_v.get("items") or {}).get("$ref")
                if _t:
                    _refs.append(_t.split("/")[-1])
        _suffixed = [t for t in _refs if _suf.search(t)]
        check("most property references carry no tool-generated suffix",
              len(_suffixed) * 2 < len(_refs),
              "%d of %d suffixed" % (len(_suffixed), len(_refs)))
        _bases = {d10["components"]["schemas"][n]["allOf"][0]["$ref"].split("/")[-1]
                  for n in d10["components"]["schemas"]
                  if "allOf" in d10["components"]["schemas"][n]}
        check("every intersection base is named <Class>__Base",
              all(b.endswith("__Base") for b in _bases), sorted(_bases))
        _plain = {b[:-len("__Base")] for b in _bases}
        check("and the plain name it gave up is a published shape",
              _plain <= set(d10["components"]["schemas"]),
              sorted(_plain - set(d10["components"]["schemas"])))
        eq("no reference dangles", mrg._dangling_refs(d10), set())
        check("groups were specialised across operations",
              m10.stats["split_groups"] >= 10, m10.stats)
        discs = [n for n, v in d10["components"]["schemas"].items()
                 if "discriminator" in v]
        eq("exactly two discriminators, one per variant-driven request",
           len(discs), 2)
        for n in discs:
            eq("%s discriminates on the request tag" % n,
               d10["components"]["schemas"][n]["discriminator"]["propertyName"],
               wbk.VARIANT_PROPERTY)

        # Split mode still produces one file per sheet and no cross-operation
        # specialisation, which is the point of keeping it.
        out_split = gen.run(CREDIT_CARD, tmp("split"), split=True,
                            log=lambda *_a: None)
        eq("split mode writes no merged document", out_split.merged, None)
        eq("and still generates every good endpoint", len(out_split.ok), 11)


# --------------------------------------------------------------------------- #
# W17 drag and drop parsing
# --------------------------------------------------------------------------- #


def _load_dropped_paths():
    """``_dropped_paths`` from the GUI source, without importing tkinter.

    The parsing has no GUI dependency, so it can be checked on interpreters
    that have no tkinter, which is how the build machine is configured.
    """
    import ast
    source = open(os.path.join(ROOT, "polymorphize_gui.py"),
                  encoding="utf-8").read()
    tree = ast.parse(source)
    fn = next(n for n in tree.body
              if isinstance(n, ast.FunctionDef) and n.name == "_dropped_paths")
    namespace = {}
    exec(compile(ast.Module([fn], []), "<gui>", "exec"), namespace)  # noqa: S102
    return namespace["_dropped_paths"]


def w17():
    group("W17 — dropped path parsing")
    try:
        import tkinter                                         # noqa: F401
    except Exception:                                          # noqa: BLE001
        print("  ....  tkinter is not available in this interpreter, so the "
              "GUI module cannot be imported here. Parsing is checked against "
              "the same source below.")
        f = _load_dropped_paths()
    else:
        import importlib.util
        spec = importlib.util.spec_from_file_location(
            "_gui_probe", os.path.join(ROOT, "polymorphize_gui.py"))
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        check("the GUI module imports whether or not tkinterdnd2 is present",
              True)
        f = mod._dropped_paths
    eq("one plain path", f("/tmp/book.xlsx"), ["/tmp/book.xlsx"])
    eq("a braced path with spaces",
       f("{C:/My Folder/book.xlsx}"), ["C:/My Folder/book.xlsx"])
    eq("two paths, one braced",
       f("{C:/My Folder/a.xlsx} /tmp/b.xlsx"),
       ["C:/My Folder/a.xlsx", "/tmp/b.xlsx"])
    eq("two braced paths",
       f("{/a b/one.xlsx} {/c d/two.xlsx}"),
       ["/a b/one.xlsx", "/c d/two.xlsx"])
    eq("nothing dropped", f(""), [])



# --------------------------------------------------------------------------- #
# W18 the format classifier
# --------------------------------------------------------------------------- #

FIELD_MAPPING = os.path.join(ROOT, "samples",
                             "issueddevice_fieldmapping_v1.0.7.xlsx")


def w18():
    group("W18 — classifying the input format")
    import polymorphize_classify as cls

    lvl = cls.classify(CREDIT_CARD)
    eq("the credit card workbook is the Level format", lvl.fmt, cls.LEVEL)
    check("and it is readable", lvl.readable)
    check("it does not ask for a specification to trim",
          not lvl.needs_specification)

    tpl = cls.classify(os.path.join(ROOT, "samples", "SOR_mapping_template.xlsx"))
    eq("so is the template", tpl.fmt, cls.LEVEL)
    check("and its support sheets are counted separately",
          tpl.support_sheets >= 1, tpl.line())

    if os.path.exists(FIELD_MAPPING):
        fm = cls.classify(FIELD_MAPPING)
        eq("the field mapping workbook is the mapping format", fm.fmt, cls.MAPPING)
        check("its header is found on row 10, not in the first six",
              all(v.header_row == 10 for v in fm.per_sheet), 
              [v.header_row for v in fm.per_sheet])
        check("and it asks for a specification to trim", fm.needs_specification)

    # Neither format. This is the case that used to pass silently.
    nothing = tmp("plain.xlsx")
    build({"Sheet1": [["a", "b"], [1, 2]]}, nothing)
    un = cls.classify(nothing)
    eq("a workbook that is neither format is not recognised", un.fmt, cls.UNKNOWN)
    check("and says why", "no Level columns" in un.line(), un.line())

    # Fail soft rather than raise: these files are often open in Excel.
    gone = cls.classify(os.path.join(ROOT, "no-such-file.xlsx"))
    check("a missing file is reported, not raised", not gone.readable)
    check("and the message says it cannot be read yet",
          "Cannot read" in gone.line(), gone.line())

    # The peek must stay bounded. An ordinary open of an 8.7 MB workbook takes
    # about 76 seconds, which at a file field looks like a hang.
    for label, path in (("credit card", CREDIT_CARD), ("mapping", FIELD_MAPPING)):
        if not os.path.exists(path):
            continue
        c = cls.classify(path)
        check("classifying the %s workbook stays under %.0fs (%.2fs)"
              % (label, cls.PEEK_BUDGET, c.elapsed),
              c.elapsed < cls.PEEK_BUDGET, c.elapsed)
    sampled = cls.peek(CREDIT_CARD)
    check("the peek reads at most PEEK_SHEETS sheets",
          len(sampled) <= cls.PEEK_SHEETS, len(sampled))
    check("and at most PEEK_ROWS rows of each",
          all(len(g) <= cls.PEEK_ROWS for _t, g in sampled))


# --------------------------------------------------------------------------- #
# W18 the field mapping format
# --------------------------------------------------------------------------- #


def w19():
    group("W19 — reading the field mapping format")
    if not os.path.exists(FIELD_MAPPING):
        check("the field mapping sample is present", False, FIELD_MAPPING)
        return
    import polymorphize_mapping as mapfmt

    results = mapfmt.read_workbook(FIELD_MAPPING)
    eq("every sheet reads", sorted({r.status for r in results}), ["ok"])
    eq("twenty-two operation sheets", len(results), 22)

    card = next(r for r in results if r.sheet == "Card Create")
    eq("the banner is read from the rows above the header",
       card.banner.get("service_domain"), "Issued Device Administration")
    eq("including the method", card.banner.get("method"), "POST")
    eq("and the SOR name, which the Level format has no cell for",
       card.banner.get("sor_name"), "Thales")
    eq("the SOR endpoint drops its method prefix",
       card.endpoints, ["/v2/issuers/{issuerId}/cards"])

    # The dotted path is the hierarchy.
    root = card.request.root
    eq("the top-level path segment names the schema", root.name,
       "CardCreateRequest")
    party = next(c for c in root.children if c.name == "PartyIdentifier")
    eq("a two-segment path nests under its parent", party.json_type, "object")
    check("and a three-segment path nests under that",
          "partyIdentification" in [g.name for g in party.children],
          [g.name for g in party.children])
    eq("a length in the Schema column is read",
       next(g.max_length for g in party.children
            if g.name == "partyIdentification"), 64)
    eq("Mandatory in the Usage column is read",
       next(g.usage for g in party.children
            if g.name == "partyIdentification"), wbk.USAGE_REQUIRED)
    arr = next(c for c in root.children if c.name == "AccountList")
    eq("an array is typed from the Schema column", arr.json_type, "array")

    # Header rows are recognised and skipped, which is this format's version
    # of the rule that only body sections are read.
    check("header and parameter rows are skipped", card.ignored_rows > 0,
          card.ignored_rows)
    check("and the skipped sections are named",
          any("Header" in lab for lab, _r in card.ignored_sections),
          card.ignored_sections)

    # N/A and an empty cell both mean unmapped.
    mapped = [c for c in root.children if c.sor]
    check("only rows naming an SOR field are mapped", 0 < len(mapped) < len(root.children),
          (len(mapped), len(root.children)))

    # A group declared with no members publishes nothing and says so.
    details = next(r for r in results if r.sheet == "Token_Details")
    check("a group declared with no members is reported as A006",
          any(f.code == "A006" for f in details.findings),
          [f.code for f in details.findings])

    # The whole point: the tree is the same shape the Level reader produces, so
    # the generator needs no knowledge of which format was read.
    spec = gen.generate_sheet(card)
    eq("the sheet generates", spec.status, "ok")
    eq("with no empty schema", gen.empty_schemas(spec.document), [])
    eq("and one path", len(spec.document["paths"]), 1)
    check("at the proposed reusable endpoint",
          spec.path.startswith("/issued-device-administration"), spec.path)


# --------------------------------------------------------------------------- #
# W18 the Apigee error set
# --------------------------------------------------------------------------- #


def w20():
    group("W20 — the Apigee error set")
    import polymorphize_errors as errs

    eq("twelve errors are defined", len(errs.APIGEE_ERRORS), 12)
    eq("and they are the codes the API COE specification carries",
       errs.error_codes(),
       ["400", "401", "403", "404", "405", "409", "422", "429",
        "500", "502", "503", "504"])

    r = read_one(one_sheet_workbook(tmp("a.xlsx")))
    spec = gen.generate_sheet(r)
    d = spec.document
    op = d["paths"][spec.path][spec.method]

    for name, code, _reason in errs.APIGEE_ERRORS:
        eq("%s is attached as %s" % (name, code),
           op["responses"][code],
           {"$ref": "#/components/responses/%s" % name})
    eq("the twelve are shared components, not inlined",
       sorted(d["components"]["responses"]),
       sorted(n for n, _c, _r in errs.APIGEE_ERRORS))
    check("each carries the Apigee disclaimer verbatim",
          all(errs.DISCLAIMER in v["description"]
              for v in d["components"]["responses"].values()))
    eq("the error schema is published once",
       errs.ERROR_SCHEMA_NAME in d["components"]["schemas"], True)

    # errors is an array in the schema, so it must be an array in the example.
    schema = d["components"]["schemas"][errs.ERROR_SCHEMA_NAME]
    eq("errors is declared an array", schema["properties"]["errors"]["type"],
       "array")
    check("and the example provides an array, unlike the source sample",
          isinstance(schema["example"]["errors"], list), schema["example"])
    eq("the example carries every declared property",
       sorted(schema["example"]), sorted(schema["properties"]))
    eq("and its member carries every declared member property",
       sorted(schema["example"]["errors"][0]),
       sorted(schema["properties"]["errors"]["items"]["properties"]))

    # The three headers are a house convention, not an error convention.
    eq("three standard headers are published",
       sorted(d["components"]["headers"]), sorted(errs.STANDARD_HEADERS))
    for code, response in op["responses"].items():
        if "$ref" in response:
            continue
        eq("the %s response carries the standard headers too" % code,
           sorted(response.get("headers", {})), sorted(errs.STANDARD_HEADERS))

    # Attaching twice must not double anything.
    before = json.dumps(d, sort_keys=True, default=str)
    errs.attach(d)
    eq("attaching the set twice changes nothing",
       json.dumps(d, sort_keys=True, default=str), before)


# --------------------------------------------------------------------------- #
# W18 class and context in a component name
# --------------------------------------------------------------------------- #


def w21():
    group("W21 — separating a class from its context")
    import polymorphize_merge as mrg

    eq("the separator is legal in an OpenAPI component key", mrg.SEP, "__")
    eq("and the tilde a following process asked for is a value, not a key",
       mrg.QUALIFIED_SEP, "~")
    eq("a qualified name joins the two", mrg.qualified("Account", "Retrieve"),
       "Account__Retrieve")
    eq("an unqualified name is left alone", mrg.qualified("Account"), "Account")
    eq("and it splits back", mrg.split_qualified("Account__Retrieve"),
       ("Account", "Retrieve"))
    eq("an unqualified name splits to no context",
       mrg.split_qualified("Account"), ("Account", None))

    stamped = mrg.identify({}, "Account__CardDetailsRetrieve")
    eq("x-class carries the class", stamped["x-class"], "Account")
    eq("x-context carries the context", stamped["x-context"],
       "CardDetailsRetrieve")
    eq("x-qualified-name carries the tilde form",
       stamped["x-qualified-name"], "Account~CardDetailsRetrieve")
    plain = mrg.identify({}, "Account")
    check("an unqualified schema has no x-context", "x-context" not in plain)
    eq("and its qualified name is just the class",
       plain["x-qualified-name"], "Account")

    if not os.path.exists(CREDIT_CARD):
        return
    m = _merge(CREDIT_CARD)
    S = m.document["components"]["schemas"]
    eq("no component key contains a tilde",
       [n for n in S if "~" in n], [])
    bad = [n for n in S if not re.fullmatch(r"[a-zA-Z0-9.\-_]+", n)]
    eq("every component key is inside the OpenAPI character set", bad, [])
    check("every schema is identified", all("x-class" in v for v in S.values()),
          [n for n, v in S.items() if "x-class" not in v][:4])
    qualified = {n: v for n, v in S.items() if mrg.SEP in n}
    check("qualified names exist to check", bool(qualified))
    for n, v in qualified.items():
        cls, ctx = mrg.split_qualified(n)
        eq("%s splits to its own class" % n, v["x-class"], cls)
        eq("%s splits to its own context" % n, v["x-context"], ctx)
        eq("%s carries the tilde form" % n, v["x-qualified-name"],
           "%s~%s" % (cls, ctx))
    bases = {v["allOf"][0]["$ref"].split("/")[-1]
             for v in S.values() if "allOf" in v}
    check("every intersection base is <Class>__Base",
          all(b.endswith(mrg.SEP + "Base") for b in bases), sorted(bases))


# --------------------------------------------------------------------------- #
# W22 the API lifecycle status
# --------------------------------------------------------------------------- #


def w22():
    group("W22 — the API lifecycle status on every operation")
    eq("the description is the bold underlined lifecycle line",
       gen.LIFECYCLE_DESCRIPTION,
       "**<u>API Lifecycle Status - Design</u>**")

    if not os.path.exists(CREDIT_CARD):
        check("the reference workbook is present", False, CREDIT_CARD)
        return
    m = _merge(CREDIT_CARD)
    ops = [op for entry in m.document["paths"].values()
           for method, op in entry.items() if isinstance(op, dict)]
    check("there are operations to check", bool(ops), len(ops))
    eq("every operation carries the lifecycle description",
       {op.get("description") for op in ops},
       {"**<u>API Lifecycle Status - Design</u>**"})
    eq("and the status as an extension a machine can read",
       {op.get("x-api-lifecycle-status") for op in ops}, {"Design"})
    # The use case is not lost to make room for it.
    check("the use case survives as the summary",
          all(op.get("summary") for op in ops))
    check("and untruncated as x-use-case",
          any(op.get("x-use-case") for op in ops))
    longest = max((op.get("x-use-case", "") for op in ops), key=len)
    check("x-use-case is not the truncated summary", len(longest) > 0,
          longest[:40])


# --------------------------------------------------------------------------- #
# W23 the field mapping export
# --------------------------------------------------------------------------- #


def _export(workbook, tmp, **kw):
    """Run and export, returning ``(RunResult, path)``."""
    import polymorphize_export as expmod
    out = gen.run(workbook, tmp, write=False, log=lambda *_a: None, **kw)
    path = os.path.join(tmp, "export.xlsx")
    expmod.write(out, path)
    return out, path


def w23():
    group("W23 — the field mapping export")
    import polymorphize_export as expmod
    if not os.path.exists(FIELD_MAPPING):
        check("the field mapping sample is present", False, FIELD_MAPPING)
        return

    eq("the format's own nine columns come first",
       list(expmod.FORMAT_COLUMNS[:2]),
       ["Parameter Type", "Reusable API Field Name"])
    eq("the header sits on row 10, as in the reference document",
       expmod.HEADER_ROW, 10)
    eq("three difference columns are added and no more",
       len(expmod.DIFF_COLUMNS), 3)
    eq("the standard request header block is the seven from the reference",
       len(expmod.REQUEST_HEADERS), 7)

    tmp = tempfile.mkdtemp()
    _out, path = _export(FIELD_MAPPING, tmp)
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        check("the provenance sheet is present",
              expmod.PROVENANCE_SHEET in wb.sheetnames, wb.sheetnames[:3])
        operations = [n for n in wb.sheetnames if n != expmod.PROVENANCE_SHEET]
        eq("one sheet per operation, the sheet names preserved",
           len(operations), 22)

        ws = wb["Token_Search"]
        eq("the banner is nine rows", ws.cell(row=9, column=1).value,
           "SOR API Endpoint:")
        eq("and its first row names the API",
           ws.cell(row=1, column=1).value, "API Name")
        headings = [ws.cell(row=expmod.HEADER_ROW, column=i).value
                    for i in range(1, 13)]
        eq("the header row starts with Parameter Type", headings[0],
           "Parameter Type")
        check("the SOR column heading carries the endpoint",
              "\n" in (headings[4] or ""), headings[4])
        eq("the difference columns are the last three",
           headings[9:12], list(expmod.DIFF_COLUMNS))

        rows = [[c for c in r] for r in
                ws.iter_rows(min_row=expmod.HEADER_ROW + 1, values_only=True)]
        statuses = {r[9] for r in rows if r[9]}
        check("published elements are marked", "Published" in statuses,
              sorted(x for x in statuses if x))
        check("and dropped ones say why",
              any(r[9] == "No SOR field" and r[10] for r in rows))
        check("a dropped element carries an action for the analyst",
              any(r[9] == "No SOR field" and r[11] for r in rows))
        check("the standard header rows are carried",
              any(r[0] == "Header" and r[1] == "Authorization" for r in rows))
        check("both body sections are present",
              {"Request Body", "Response Body"} <=
              {r[0] for r in rows if r[0]})

        prov = wb[expmod.PROVENANCE_SHEET]
        text_of = {str(r[0]): str(r[1]) for r in
                   prov.iter_rows(max_row=12, max_col=2, values_only=True)}
        eq("the provenance sheet records the round", text_of.get("Round"), "1")
        eq("and the source workbook",
           text_of.get("Source workbook"),
           os.path.basename(FIELD_MAPPING))
        eq("and the lifecycle status",
           text_of.get("API lifecycle status"), "Design")
    finally:
        wb.close()

    # A Level workbook with several SOR endpoints becomes several sheets.
    if os.path.exists(CREDIT_CARD):
        tmp2 = tempfile.mkdtemp()
        _out2, path2 = _export(CREDIT_CARD, tmp2)
        wb2 = openpyxl.load_workbook(path2, data_only=True)
        try:
            names = [n for n in wb2.sheetnames
                     if n != expmod.PROVENANCE_SHEET]
            expected = sum(len(r.layout.sor_cols) for r in _out2.results
                           if r.status == "ok" and r.layout is not None)
            eq("one sheet per SOR endpoint, across the whole workbook",
               len(names), expected)
            check("the five endpoints of Card Details each get a sheet",
                  sum(1 for n in names if n.startswith("Card Detail")) == 5,
                  [n for n in names if n.startswith("Card Detail")])
            check("and the three of Card Transaction",
                  sum(1 for n in names if n.startswith("Card Transaction")) == 3,
                  [n for n in names if n.startswith("Card Transaction")])
            check("each of those sheets names a different SOR endpoint",
                  len({wb2[n].cell(row=9, column=2).value
                       for n in names if n.startswith("Card Detail")}) == 5)
            check("every sheet name is Excel-legal and short enough",
                  all(len(n) <= 31 and not set(n) & set("[]:*?/\\")
                      for n in wb2.sheetnames), wb2.sheetnames)
            first = wb2[names[0]]
            check("the header block is supplied where the format had none",
                  any(r[1] == "x-BDO-Application-Id" for r in
                      first.iter_rows(min_row=expmod.HEADER_ROW + 1,
                                      max_col=2, values_only=True)))
        finally:
            wb2.close()


# --------------------------------------------------------------------------- #
# W24 the round trip
# --------------------------------------------------------------------------- #


def _tree_signature(result):
    """Everything about a sheet that decides what the specification says."""
    out = []
    for section in (result.request, result.response):
        if section is None or section.root is None:
            continue

        def walk(node, path):
            for child in node.children:
                here = path + (child.name,)
                out.append((section.kind, ".".join(here), child.json_type,
                            child.fmt, child.max_length, child.usage,
                            tuple(sorted(child.sor_declared.items()))))
                walk(child, here)
        walk(section.root, (section.root.name,))
    return out


def w24():
    group("W24 — the round trip")
    import polymorphize_export as expmod
    import polymorphize_classify as cls
    import polymorphize_mapping as mapfmt
    if not os.path.exists(FIELD_MAPPING):
        check("the field mapping sample is present", False, FIELD_MAPPING)
        return

    tmp = tempfile.mkdtemp()
    _out, first = _export(FIELD_MAPPING, tmp)

    verdict = cls.classify(first)
    eq("the export is itself a field mapping document", verdict.fmt,
       cls.MAPPING)

    before = {r.sheet: r for r in mapfmt.read_workbook(FIELD_MAPPING)
              if r.status == "ok"}
    after = {r.sheet: r for r in mapfmt.read_workbook(first)
             if r.status == "ok"}
    eq("the same sheets read back", sorted(after), sorted(before))
    differing = [n for n in before
                 if _tree_signature(before[n]) != _tree_signature(after[n])]
    eq("and every tree is identical, mapping and all", differing, [])

    # The specification is the real test: the round trip must not change it.
    def spec_of(path):
        run = gen.run(path, tmp, write=False, log=lambda *_a: None)
        body = gen.dump_document(run.merged.document, "yaml")
        return re.sub(r"x-generated-at: .*", "", body)

    eq("the specification from the export equals the original",
       spec_of(first), spec_of(FIELD_MAPPING))

    # Exporting the export must not change it either, or a document would
    # drift a little on every pass through the loop.
    out2 = gen.run(first, tmp, write=False, log=lambda *_a: None)
    second = os.path.join(tmp, "second.xlsx")
    expmod.write(out2, second)

    def grid(path):
        wb = openpyxl.load_workbook(path, data_only=True)
        try:
            return {n: [tuple("" if c is None else str(c) for c in row)
                        for row in wb[n].iter_rows(values_only=True)]
                    for n in wb.sheetnames if n != expmod.PROVENANCE_SHEET}
        finally:
            wb.close()

    g1, g2 = grid(first), grid(second)
    eq("the second export has the same sheets", sorted(g2), sorted(g1))
    drift = [n for n in g1 if g1[n] != g2.get(n)]
    eq("and every cell of every sheet is unchanged", drift, [])
    eq("the round number advances", expmod.previous_round(second), 2)

    # Column creep is the failure this guards against: the difference columns
    # must be recognised on the way back in, not carried as extra columns.
    widths = {len(rows[expmod.HEADER_ROW - 1]) for rows in g2.values()}
    eq("the width is stable across rounds",
       widths, {len(rows[expmod.HEADER_ROW - 1]) for rows in g1.values()})


# --------------------------------------------------------------------------- #


def main():
    for fn in (w1, w2, w3, w4, w5, w6, w7, w8, w9, w10, w11, w12, w13, w14, w15,
               w16, w17, w18, w19, w20, w21, w22, w23, w24):
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
