#!/usr/bin/env python3
"""
Regression tests for the six defects corrected in polymorphize_core v5.0.

Run with:   python tests/test_regressions.py
No test framework required; it exits non-zero on the first failure.
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl                                  # noqa: E402
import polymorphize_core as core                 # noqa: E402

PASSED = []
FAILED = []


def check(name, condition, detail=""):
    (PASSED if condition else FAILED).append(name)
    print(("  PASS  " if condition else "  FAIL  ") + name + (f"   {detail}" if detail else ""))


# --------------------------------------------------------------------------- #
# Fixtures
# --------------------------------------------------------------------------- #
def make_casa_style_mapping(path, rows, header_row=11, banner=True):
    """Header part-way down the sheet; one dotted-path column; SOR in col 5."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Op_Initiate"
    if banner:
        for r, (k, v) in enumerate([("API Name", "Demo"), ("Use Case:", "Demo use case"),
                                    ("Service Domain:", "DemoDomain"),
                                    ("Method", "POST"),
                                    # deliberately the un-hyphenated label variant
                                    ("Proposed Reusable API Endpoint:", "/demo/create"),
                                    ("SOR Name:", "Evolution")], start=1):
            ws.cell(r, 1, k)
            ws.cell(r, 2, v)
    ws.cell(header_row, 1, "Parameter Type")
    ws.cell(header_row, 2, "Re-usable API Field Name")
    ws.cell(header_row, 3, "Usage")
    ws.cell(header_row, 4, "Schema")
    ws.cell(header_row, 5, "SOR API Field Name - /demo/create")
    ws.cell(header_row, 6, "Examples")
    ws.cell(header_row, 7, "Description")
    r = header_row + 1
    ws.cell(r, 1, "Request Body")
    ws.cell(r, 2, "\xa0")
    r += 1
    for path_txt, dtype, sor in rows:
        ws.cell(r, 1, "Body")
        ws.cell(r, 2, path_txt)
        ws.cell(r, 3, "Mandatory")
        ws.cell(r, 4, dtype)
        ws.cell(r, 5, sor)
        r += 1
    wb.save(path)
    return path


def make_legacy_mapping(path, rows):
    """The v4 layout: Level 1..5 in C..G, Data Type in I, SOR Field in J."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Legacy"
    for i in range(5):
        ws.cell(1, 3 + i, f"Level {i + 1}")
    ws.cell(1, 9, "Data Type")
    ws.cell(1, 10, "SOR Field")
    r = 2
    for levels, dtype, sor in rows:
        for i, lv in enumerate(levels):
            ws.cell(r, 3 + i, lv)
        ws.cell(r, 9, dtype)
        ws.cell(r, 10, sor)
        r += 1
    wb.save(path)
    return path


SWAGGER = """openapi: 3.0.3
info:
  title: Demo
  version: 1.0.0
paths:
  /demo/create:
    post:
      requestBody:
        content:
          application/json:
            schema:
              $ref: '#/components/schemas/DemoRequest'
      responses:
        '200':
          content:
            application/json:
              schema:
                $ref: '#/components/schemas/DemoResponse'
components:
  schemas:
    DemoRequest:
      type: object
      properties:
        keptField:
          type: string
          description: has a SOR field
          example: ABC
        droppedField:
          type: string
          description: explicitly not used in the SOR
        unknownField:
          type: string
          description: not mentioned in the mapping at all
        Party:
          $ref: '#/components/schemas/Party'
        Amounts:
          type: array
          items:
            $ref: '#/components/schemas/Amount'
      required:
      - keptField
      - droppedField
    Party:
      type: object
      description: Party details
      properties:
        partyIdType:
          type: string
        partyId:
          type: string
      required:
      - partyIdType
    Amount:
      type: object
      description: An amount
      properties:
        amountType:
          type: string
        amountValue:
          type: integer
    DemoResponse:
      type: object
      properties:
        Txn:
          type: object
          description: Transaction
          properties:
            txnId:
              type: string
            txnType:
              type: string
            txnStatus:
              type: string
        keptField:
          type: string
          description: has a SOR field
          example: ABC
    Txn:
      type: object
      description: Transaction
      properties:
        txnId:
          type: string
        txnType:
          type: string
        txnWhen:
          type: string
"""


def write(path, text):
    with open(path, "w") as fh:
        fh.write(text)
    return path


# --------------------------------------------------------------------------- #
# Tests
# --------------------------------------------------------------------------- #
def main():
    tmp = tempfile.mkdtemp(prefix="polymorphizer_tests_")
    swagger = write(os.path.join(tmp, "demo.yaml"), SWAGGER)

    rows = [
        ("DemoRequest", "object", None),
        ("DemoRequest.keptField", "string", "sor.kept"),
        ("DemoRequest.droppedField", "string", "N/A"),
        ("DemoRequest.Party", "object", None),
        ("DemoRequest.Party.partyIdType", "string", "\xa0"),
        ("DemoRequest.Party.partyId", "string", "sor.partyId"),
        ("DemoRequest.Amounts", "array", None),
        ("DemoRequest.Amounts.amountType", "string", "Not Used In SOR"),
        ("DemoRequest.Amounts.amountValue", "integer", "-"),
        ("DemoResponse", "object", None),
        ("DemoResponse.keptField", "string", "sor.kept"),
        ("DemoResponse.Txn", "object", None),
        ("DemoResponse.Txn.txnId", "string", "sor.txnId"),
        ("DemoResponse.Txn.txnType", "string", "sor.txnType"),
        ("DemoResponse.Txn.txnStatus", "string", "sor.txnStatus"),
        ("Txn.txnId", "string", "sor.txnId"),
        ("Txn.txnType", "string", "sor.txnType"),
        ("Txn.txnWhen", "string", "sor.txnWhen"),
    ]
    mapping = make_casa_style_mapping(os.path.join(tmp, "mapping.xlsx"), rows)

    print("\nD1 — spreadsheet layout is detected, not assumed")
    m = core.load_mapping(mapping)
    lay = m.layouts[0]
    check("header row found part-way down the sheet", lay["header_row"] == 11,
          f"got {lay['header_row']}")
    check("field/dtype/SOR columns resolved by label",
          (lay["cols"].get("field"), lay["cols"].get("dtype"), lay["cols"].get("sor")) == (2, 4, 5),
          str(lay["cols"]))
    check("dotted-path mode selected", lay["mode"] == "path", lay["mode"])
    check("banner and section rows are not read as attributes",
          m.stats["attribute_rows"] == len(rows),
          f"{m.stats['attribute_rows']} of {len(rows)}")

    print("\nD1b — legacy Level 1..n layout still works")
    legacy = make_legacy_mapping(os.path.join(tmp, "legacy.xlsx"),
                                 [(["DemoRequest", "keptField"], "string", "sor.kept"),
                                  (["DemoRequest", "droppedField"], "string", "Not Used In SOR")])
    lm = core.load_mapping(legacy)
    check("levels layout detected", lm.layouts[0]["mode"] == "levels", lm.layouts[0]["mode"])
    check("legacy rows indexed by full path",
          core.norm_path(["DemoRequest", "keptField"]) in lm.by_path)

    print("\nD1c — blank, nbsp, N/A, '-' and 'Not Used In SOR' all mean 'no SOR field'")
    for token in (None, "", "\xa0", "N/A", "n/a", "-", "TBD", "Not Used In SOR"):
        check(f"'{token}' treated as not-used", core.is_not_used(token))
    check("a real SOR path is not treated as not-used", not core.is_not_used("depAcctId.acctType"))

    print("\nD2 — attributes resolve on the qualified path, not a bare leaf")
    diag = core.diagnose(swagger, mapping)["diagnostics"]
    check("every attribute except the deliberately unmapped one resolved",
          diag["attributes_unmatched"] == 1
          and diag["unmatched"][0]["attribute"] == "unknownField",
          f"{diag['attributes_unmatched']} unmatched: "
          f"{[u['attribute'] for u in diag['unmatched']]}")
    check("resolution used the exact qualified path",
          diag["strategies"].get("exact-path", 0) >= 6, str(diag["strategies"]))
    check("shared schema keeps a leaf mapped under one usage path only",
          any(k["attribute"] == "txnWhen" for k in diag["kept"]))

    print("\nD3 — an attribute the mapping does not mention is kept and flagged")
    out = os.path.join(tmp, "out.yaml")
    res = core.run(swagger, mapping, out, auto_dedupe=True, enforce_assertions=True)
    props = res_props(out, "DemoRequest")
    check("unknownField retained", "unknownField" in props, str(sorted(props)))
    check("keptField retained", "keptField" in props)
    check("droppedField eliminated", "droppedField" not in props)
    check("required list no longer names the eliminated attribute",
          "droppedField" not in (schema(out, "DemoRequest").get("required") or []))
    check("on_unmatched='drop' still available for the old behaviour",
          "unknownField" not in res_props(
              run_variant(swagger, mapping, tmp, "drop.yaml", on_unmatched="drop"), "DemoRequest"))

    print("\nD4 — objects emptied by pruning are removed with their references")
    check("Amount removed (both attributes were marked not-used in the SOR)",
          "Amount" not in schemas(out), str([n for n in schemas(out) if "Amount" in n]))
    check("the array property whose item type died is gone",
          "Amounts" not in res_props(out, "DemoRequest"))
    check("Party survives because one attribute is mapped",
          "Party" in schemas(out) and "partyId" in res_props(out, "Party"))
    check("Party.partyIdType eliminated (blank SOR cell)",
          "partyIdType" not in res_props(out, "Party"))
    check("no object schema is left with an empty properties map",
          not [n for n, sc in schemas(out).items()
               if isinstance(sc, dict) and isinstance(sc.get("properties"), dict)
               and not sc["properties"]])

    print("\nD5 — de-duplication never merges schemas that differed in the source")
    inline = res_props(out, "DemoResponse").get("Txn") if isinstance(
        res_props(out, "DemoResponse").get("Txn"), dict) else {}
    check("inline Txn was NOT collapsed into the named Txn schema",
          "$ref" not in (inline or {}), str(inline)[:120])
    check("the attribute unique to the inline object survives",
          "txnStatus" in ((inline or {}).get("properties") or {}),
          str(sorted(((inline or {}).get("properties") or {}).keys())))
    print("     (this is the AccountLimitResponse.PaymentTransaction / "
          "paymentTransactionStatus regression)")

    print("\nD5b — an inline object identical in the source IS still de-duplicated")
    dd_swagger = write(os.path.join(tmp, "dd.yaml"), SWAGGER.replace(
        """            txnStatus:
              type: string
""", """            txnWhen:
              type: string
"""))
    dd_out = os.path.join(tmp, "dd_out.yaml")
    core.run(dd_swagger, mapping, dd_out, auto_dedupe=True, enforce_assertions=False)
    check("identical inline object replaced by a $ref",
          "$ref" in (res_props(dd_out, "DemoResponse").get("Txn") or {}),
          str(res_props(dd_out, "DemoResponse").get("Txn"))[:140])

    print("\nD6 — guards stop a run whose mapping did not join")
    empty_map = make_casa_style_mapping(os.path.join(tmp, "wrong.xlsx"),
                                        [("SomethingElse.fieldA", "string", "sor.a"),
                                         ("SomethingElse.fieldB", "string", "sor.b")])
    raised = None
    try:
        core.run(swagger, empty_map, os.path.join(tmp, "never.yaml"),
                 min_match_rate=0.25, enforce_guards=True)
    except core.MappingCoverageError as e:
        raised = e
    check("a non-joining mapping raises MappingCoverageError", raised is not None,
          str(raised)[:90] if raised else "no exception")
    check("nothing was written when the guard fired",
          not os.path.exists(os.path.join(tmp, "never.yaml")))

    print("\nD6b — an unreadable layout is reported rather than silently emptying the spec")
    blank = openpyxl.Workbook()
    blank.active["A1"] = "nothing useful here"
    blank_path = os.path.join(tmp, "blank.xlsx")
    blank.save(blank_path)
    raised = None
    try:
        core.run(swagger, blank_path, os.path.join(tmp, "never2.yaml"))
    except core.MappingCoverageError as e:
        raised = e
    check("an unrecognised workbook raises rather than eliminating everything",
          raised is not None, str(raised)[:90] if raised else "no exception")

    print("\nD6c — post-run assertions")
    check("the successful run reported no assertion failures", not res["assertions"],
          str(res["assertions"][:3]))
    check("assert_output finds a deliberately emptied schema",
          bool(core.assert_output({"components": {"schemas": {
              "Hollow": {"type": "object", "properties": {}}}}})))
    check("assert_output finds a dangling $ref",
          any("dangling" in p for p in core.assert_output({"components": {"schemas": {
              "A": {"type": "object", "properties": {
                  "b": {"$ref": "#/components/schemas/Missing"}}}}}})))

    print("\nCoverage reporting")
    check("the report states the match rate", "Mapping coverage" in res["report"])
    check("the report lists the detected layout", "Detected spreadsheet layout" in res["report"])

    print("\nD7 — text encoding (reported UnicodeDecodeError in write_html_report)")
    # The specification carries an em dash. Saved as ANSI by a Windows editor
    # that becomes byte 0x97, which a hard-coded UTF-8 read cannot decode.
    utf8_src = SWAGGER.replace("title: Demo", "title: Demo — deposits")
    p_utf8 = os.path.join(tmp, "enc_utf8.yaml")
    open(p_utf8, "w", encoding="utf-8").write(utf8_src)
    p_bom = os.path.join(tmp, "enc_bom.yaml")
    open(p_bom, "wb").write(b"\xef\xbb\xbf" + utf8_src.encode("utf-8"))
    p_ansi = os.path.join(tmp, "enc_ansi.yaml")
    open(p_ansi, "wb").write(utf8_src.encode("cp1252"))

    check("UTF-8 input detected", core.read_text(p_utf8)[1] == "utf-8",
          core.read_text(p_utf8)[1])
    check("UTF-8 BOM detected and stripped",
          core.read_text(p_bom)[1] == "utf-8 (BOM)"
          and core.read_text(p_bom)[0].startswith("openapi"),
          core.read_text(p_bom)[1])
    check("ANSI (cp1252) input detected", core.read_text(p_ansi)[1] == "cp1252",
          core.read_text(p_ansi)[1])
    check("the em dash survives the ANSI round trip",
          "—" in core.read_text(p_ansi)[0])

    enc_out = os.path.join(tmp, "enc_out.yaml")
    enc_html = os.path.join(tmp, "enc_out.html")
    enc_md = os.path.join(tmp, "enc_out.md")
    raised = None
    try:
        enc_res = core.run(p_ansi, mapping, enc_out, enc_md, html_report=enc_html,
                           auto_dedupe=True, enforce_assertions=False)
    except Exception as e:                                    # noqa: BLE001
        raised = e
        enc_res = None
    check("an ANSI specification runs end to end", raised is None,
          f"{type(raised).__name__}: {raised}" if raised else "")
    check("the HTML change report is produced from an ANSI input",
          os.path.exists(enc_html) and os.path.getsize(enc_html) > 0)
    check("the output YAML is written as UTF-8",
          core.read_text(enc_out)[1] == "utf-8", core.read_text(enc_out)[1])
    check("the Markdown report is written as UTF-8",
          core.read_text(enc_md)[1] == "utf-8", core.read_text(enc_md)[1])
    check("no report errors were recorded",
          enc_res is not None and not enc_res["report_errors"],
          str(enc_res["report_errors"]) if enc_res else "")

    # A failure while building a report must not discard a good transform.
    original = core.write_html_report

    def boom(*a, **kw):
        raise UnicodeDecodeError("utf-8", b"\x97", 0, 1, "simulated")

    core.write_html_report = boom
    try:
        keep_out = os.path.join(tmp, "keep_out.yaml")
        kr = core.run(p_utf8, mapping, keep_out,
                      html_report=os.path.join(tmp, "keep.html"),
                      enforce_assertions=False)
        check("a failing HTML report no longer loses the transform",
              os.path.exists(keep_out) and kr["report_errors"],
              str(kr["report_errors"]))
    finally:
        core.write_html_report = original

    print("\nD8 — superseded")
    check("the dotted-path generator has been replaced by the Level-format "
          "reader, covered by tests/test_workbook_contract.py", True)
    print("\nD9 — per-operation allOf profiling (the tool's purpose)")
    # One shared schema, two operations, different mapped subsets, plus a third
    # position that maps nothing at all.
    PROF_SPEC = """openapi: 3.0.3
info: {title: P, version: 1.0.0}
paths:
  /a/initiate:
    post:
      requestBody:
        content: {application/json: {schema: {$ref: '#/components/schemas/AReq'}}}
      responses: {'200': {content: {application/json: {schema: {$ref: '#/components/schemas/ARes'}}}}}
  /b/initiate:
    post:
      requestBody:
        content: {application/json: {schema: {$ref: '#/components/schemas/BReq'}}}
      responses: {'200': {content: {application/json: {schema: {$ref: '#/components/schemas/BRes'}}}}}
components:
  schemas:
    Shared:
      type: object
      properties:
        common: {type: string}
        onlyA:  {type: string}
        onlyB:  {type: string}
      required: [common, onlyA]
    AReq:
      type: object
      properties: {Shared: {$ref: '#/components/schemas/Shared'}}
    BReq:
      type: object
      properties: {Shared: {$ref: '#/components/schemas/Shared'}}
    ARes:
      type: object
      properties: {Shared: {$ref: '#/components/schemas/Shared'}}
    BRes:
      type: object
      properties: {keep: {type: string}}
"""
    pspec = write(os.path.join(tmp, "prof.yaml"), PROF_SPEC)
    pmap = make_casa_style_mapping(os.path.join(tmp, "profmap.xlsx"), [
        ("AReq", "object", None),
        ("AReq.Shared", "object", None),
        ("AReq.Shared.common", "string", "sor.common"),
        ("AReq.Shared.onlyA", "string", "sor.onlyA"),
        ("AReq.Shared.onlyB", "string", "Not Used In SOR"),
        ("BReq", "object", None),
        ("BReq.Shared", "object", None),
        ("BReq.Shared.common", "string", "sor.common"),
        ("BReq.Shared.onlyA", "string", "Not Used In SOR"),
        ("BReq.Shared.onlyB", "string", "sor.onlyB"),
        ("ARes", "object", None),
        ("ARes.Shared", "object", None),
        ("ARes.Shared.common", "string", "Not Used In SOR"),
        ("ARes.Shared.onlyA", "string", "Not Used In SOR"),
        ("ARes.Shared.onlyB", "string", "Not Used In SOR"),
        ("BRes", "object", None),
        ("BRes.keep", "string", "sor.keep"),
    ])
    pout = os.path.join(tmp, "prof_out.yaml")
    pr = core.run(pspec, pmap, pout, profile_shared=True, auto_dedupe=False,
                  enforce_guards=False, enforce_assertions=False)
    pdoc = _load(pout)
    ps = (pdoc.get("components") or {}).get("schemas") or {}
    shared = ps.get("Shared") or {}
    prof = pr["profiling"]

    check("shared schema reduced to the attributes every live route maps",
          sorted((shared.get("properties") or {}).keys()) == ["common"],
          str(sorted((shared.get("properties") or {}).keys())))
    derived = [n for n in ps if n.startswith("Shared") and n != "Shared"]
    check("one derived profile per distinct route", len(derived) == 2, str(derived))
    check("each derived profile inherits the base with allOf",
          all("allOf" in (ps[n] or {}) and
              (ps[n]["allOf"][0] or {}).get("$ref") == "#/components/schemas/Shared"
              for n in derived), str([list((ps[n] or {}).keys()) for n in derived]))
    a_ref = ((ps.get("AReq") or {}).get("properties") or {}).get("Shared", {}).get("$ref", "")
    b_ref = ((ps.get("BReq") or {}).get("properties") or {}).get("Shared", {}).get("$ref", "")
    check("each operation points at its own profile",
          a_ref != b_ref and a_ref.split("/")[-1] in derived
          and b_ref.split("/")[-1] in derived, f"A={a_ref} B={b_ref}")

    def delta(name):
        for part in (ps[name].get("allOf") or []):
            if "properties" in (part or {}):
                return sorted(part["properties"].keys())
        return []
    check("the profile for operation A carries only onlyA",
          delta(a_ref.split("/")[-1]) == ["onlyA"], str(delta(a_ref.split("/")[-1])))
    check("the profile for operation B carries only onlyB",
          delta(b_ref.split("/")[-1]) == ["onlyB"], str(delta(b_ref.split("/")[-1])))
    check("a position that maps nothing is removed entirely",
          "Shared" not in ((ps.get("ARes") or {}).get("properties") or {}),
          str(list(((ps.get("ARes") or {}).get("properties") or {}).keys())))
    check("removal is reported", any("ARes" in x for x in prof["removed_positions"]),
          str(prof["removed_positions"]))
    check("no discriminator is introduced anywhere",
          "discriminator" not in open(pout, encoding="utf-8").read())
    check("required moves with the attribute, not left contradicting the mapping",
          "onlyA" not in (shared.get("required") or []),
          str(shared.get("required")))
    check("switching profiling off leaves the union in place",
          sorted((((_load(run_variant(pspec, pmap, tmp, "noprof.yaml",
                                      profile_shared=False)) or {})
                   .get("components") or {}).get("schemas") or {})
                 .get("Shared", {}).get("properties", {}).keys())
          == ["common", "onlyA", "onlyB"])

    print("\nD10 — superseded")
    check("workbook validation and the template now follow the Level-format "
          "contract, covered by tests/test_workbook_contract.py", True)
    print("\nD11 — workbook reads are bounded")
    wide = openpyxl.Workbook()
    wsx = wide.active
    wsx.cell(1, 1, "data")
    wsx.cell(60000, 1, None)
    wsx.row_dimensions[60000].height = 15          # forces a huge declared dimension
    wide_path = os.path.join(tmp, "wide.xlsx")
    wide.save(wide_path)
    grids = core.load_grids(wide_path)
    check("a sheet padded with empty rows is truncated to its data",
          grids[0].max_row < 400, f"max_row={grids[0].max_row}")
    check("the grid still exposes cell values by row and column",
          grids[0].cell(1, 1).value == "data")

    print(f"\n{len(PASSED)} passed, {len(FAILED)} failed")
    if FAILED:
        for f in FAILED:
            print(f"  FAILED: {f}")
        return 1
    return 0


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _load(path):
    from ruamel.yaml import YAML
    y = YAML(typ="safe")
    with open(path) as fh:
        return y.load(fh)


def schemas(path):
    return (_load(path).get("components") or {}).get("schemas") or {}


def schema(path, name):
    return schemas(path).get(name) or {}


def res_props(path, name):
    return schema(path, name).get("properties") or {}


def run_variant(swagger, mapping, tmp, filename, **kw):
    out = os.path.join(tmp, filename)
    core.run(swagger, mapping, out, enforce_guards=False, enforce_assertions=False, **kw)
    return out


if __name__ == "__main__":
    sys.exit(main())
