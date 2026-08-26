#!/usr/bin/env python3
"""
polymorphize_cli.py  —  command-line front-end for polymorphize_core (v5.1)
===========================================================================

Examples
--------
Coverage check first (writes nothing, tells you whether the mapping actually
joined to the specification):

    python polymorphize_cli.py --diagnose \
        --swagger casa.yaml --mapping casa_mapping.xlsx

Transform:

    python polymorphize_cli.py \
        --swagger casa.yaml --mapping casa_mapping.xlsx \
        --out casa_polymorphic.yaml --report report.md \
        --html-report changes.html --auto-dedupe

Explicit groups / de-duplication:

    python polymorphize_cli.py --swagger casa.yaml --mapping casa_mapping.xlsx \
        --out out.yaml \
        --poly "ResultBase = InitiateResponse:initiate, UpdateResponse:update" \
        --dedupe "AmountBlockControlRequest.AmountBlock = AmountBlock"

Behaviour change in v5.0
------------------------
An attribute the mapping does not mention is KEPT and flagged, not deleted.
Pass --on-unmatched drop to restore the old behaviour, and read the coverage
figures in the report before you do.

Behaviour change in v5.1
------------------------
The input encoding is detected (UTF-8, UTF-8 with BOM, or ANSI/cp1252) and the
output and reports are always written as UTF-8. A report that fails to build is
reported as a warning rather than discarding the transformed document.
"""
import argparse
import sys

import polymorphize_core as core


def _print_layouts(layouts):
    print("Detected spreadsheet layout:")
    for lay in layouts:
        cols = ", ".join(f"{k}=col{v}" for k, v in sorted(lay["cols"].items()))
        lv = f"  levels={lay['levels']}" if lay["levels"] else ""
        print(f"  [{lay['sheet']}] mode={lay['mode']} header_row={lay['header_row']}  {cols}{lv}")


def _diagnose(args):
    res = core.diagnose(args.swagger, args.mapping,
                        sheet_cfg=_sheet_cfg(args),
                        protected=_protected(args),
                        on_unmatched=args.on_unmatched,
                        log=lambda m: print(m))
    d = res["diagnostics"]
    _print_layouts(res["layouts"])
    st = res["mapping_stats"]
    print(f"\nMapping: {st['attribute_rows']} attribute row(s) from {st['sheets']} sheet(s) "
          f"({st['mapped']} mapped, {st['unmapped']} marked not-used, "
          f"{st['section_rows']} non-attribute rows skipped)")
    print(f"Specification: {d['attributes_total']} attribute(s)")
    print(f"Resolved:      {d['attributes_matched']} ({d['match_rate'] * 100:.1f}%)  "
          f"strategies={d['strategies']}")
    print(f"Not found:     {d['attributes_unmatched']}")
    if d["unmatched"]:
        print("\nNot found in the mapping (first 40), with the keys that were tried:")
        for u in d["unmatched"][:40]:
            near = res["nearest"].get(u["path"]) or []
            hint = f"   nearest: {', '.join(near)}" if near else ""
            print(f"  {u['path']}\n      tried: {', '.join(u['keys_tried'])}{hint}")
    if d["match_rate"] < 0.5:
        print("\nA low match rate means the lookup failed, not that the SOR lacks the fields. "
              "Check the field-name column in the layout above before running a transform.")
    return 0


def _sheet_cfg(args):
    return {
        "level_first_col": args.level_first_col,
        "level_last_col": args.level_last_col,
        "dtype_col": args.dtype_col,
        "sor_col": args.sor_col,
        "not_used_text": args.not_used_text,
        "auto_layout": not args.no_auto_layout,
        "header_scan_rows": args.header_scan_rows,
    }


def _protected(args):
    return tuple(s.strip() for s in args.protected.split(",") if s.strip())


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Eliminate attributes with no SOR connection from an OpenAPI file.")
    ap.add_argument("--swagger", required=True, help="input OpenAPI/Swagger YAML")
    ap.add_argument("--mapping", required=True, help="SOR mapping spreadsheet (.xlsx/.xlsm)")
    ap.add_argument("--out", help="output YAML path (required unless --diagnose)")
    ap.add_argument("--report", default=None, help="optional Markdown report path")
    ap.add_argument("--html-report", default=None, help="optional colour-coded HTML change report")
    ap.add_argument("--diagnose", action="store_true",
                    help="report mapping coverage and exit without writing anything")

    ap.add_argument("--sor-swagger", action="append", default=[], metavar="YAML",
                    help="SOR API YAML used to validate mapped fields (repeatable, or ';'-separated)")
    ap.add_argument("--require-sor-field", action="store_true",
                    help="also eliminate a mapped attribute if its SOR field is absent from the SOR YAML(s)")

    ap.add_argument("--on-unmatched", choices=("keep", "drop"), default="keep",
                    help="what to do with an attribute the mapping does not mention "
                         "(default: keep and flag)")
    ap.add_argument("--keep-absent", action="store_true",
                    help="deprecated; keeping unmatched attributes is now the default")

    ap.add_argument("--discriminator", action="store_true",
                    help="add discriminator + mapping to generated base schemas")
    ap.add_argument("--disc-prop", default="resultContext", help="discriminator property name")
    ap.add_argument("--poly", action="append", default=[], metavar="SPEC",
                    help="polymorphism group 'Base = SchemaA:valA, SchemaB:valB' (repeatable)")
    ap.add_argument("--dedupe", action="append", default=[], metavar="SPEC",
                    help="inline de-dupe 'Owner.prop = Target' (repeatable)")
    ap.add_argument("--auto-poly", action="store_true", help="auto-detect families to factor")
    ap.add_argument("--auto-dedupe", action="store_true",
                    help="auto-replace inline objects equivalent to a named schema")
    ap.add_argument("--dedupe-strictness", choices=("deep", "shape"), default="deep",
                    help="'deep' requires full equality; 'shape' ignores prose and examples")
    ap.add_argument("--protected", default="ErrorResponse",
                    help="comma-separated schema names never pruned")

    ap.add_argument("--min-match-rate", type=float, default=0.25,
                    help="abort if the mapping resolves less than this fraction of attributes")
    ap.add_argument("--max-total-loss", type=float, default=0.75,
                    help="abort if more than this fraction of attributes would be eliminated")
    ap.add_argument("--no-guards", action="store_true",
                    help="downgrade the coverage and loss guards to warnings")
    ap.add_argument("--no-assert", action="store_true",
                    help="write the output even if a post-run assertion fails")

    ap.add_argument("--level-first-col", type=int, default=core.DEFAULT_SHEET_CFG["level_first_col"])
    ap.add_argument("--level-last-col", type=int, default=core.DEFAULT_SHEET_CFG["level_last_col"])
    ap.add_argument("--dtype-col", type=int, default=core.DEFAULT_SHEET_CFG["dtype_col"])
    ap.add_argument("--sor-col", type=int, default=core.DEFAULT_SHEET_CFG["sor_col"])
    ap.add_argument("--not-used-text", default=core.DEFAULT_SHEET_CFG["not_used_text"])
    ap.add_argument("--no-auto-layout", action="store_true",
                    help="do not detect the header row and columns; use the fixed columns above")
    ap.add_argument("--header-scan-rows", type=int,
                    default=core.DEFAULT_SHEET_CFG["header_scan_rows"],
                    help="how many rows to scan when looking for the header row")

    args = ap.parse_args(argv)

    if args.diagnose:
        return _diagnose(args)
    if not args.out:
        ap.error("--out is required unless --diagnose is used")

    sor_swaggers = []
    for s in args.sor_swagger:
        sor_swaggers.extend(p.strip() for p in str(s).split(";") if p.strip())

    on_unmatched = "keep" if args.keep_absent else args.on_unmatched

    try:
        result = core.run(
            args.swagger, args.mapping, args.out, args.report,
            html_report=args.html_report, sor_swaggers=sor_swaggers,
            require_sor_field=args.require_sor_field, on_unmatched=on_unmatched,
            add_discriminator=args.discriminator, discriminator_prop=args.disc_prop,
            poly_groups=[core.parse_group_spec(s) for s in args.poly],
            dedupe_pairs=[core.parse_dedupe_spec(s) for s in args.dedupe],
            auto_poly=args.auto_poly, auto_dedupe=args.auto_dedupe,
            dedupe_strictness=args.dedupe_strictness,
            protected=_protected(args), sheet_cfg=_sheet_cfg(args),
            min_match_rate=args.min_match_rate, max_total_loss=args.max_total_loss,
            enforce_guards=not args.no_guards, enforce_assertions=not args.no_assert,
            log=lambda m: print(m),
        )
    except core.MappingCoverageError as e:
        print(f"\nABORTED — mapping coverage: {e}", file=sys.stderr)
        print("Run again with --diagnose to see the keys that were tried.", file=sys.stderr)
        return 3
    except core.ExcessiveLossError as e:
        print(f"\nABORTED — excessive loss: {e}", file=sys.stderr)
        return 4
    except core.OutputAssertionError as e:
        print(f"\nASSERTION FAILURE: {e}", file=sys.stderr)
        return 5

    d = result["diagnostics"]
    print(f"\nMapping resolved {d['attributes_matched']}/{d['attributes_total']} attribute(s) "
          f"({d['match_rate'] * 100:.1f}%).")
    print(f"Eliminated {len(result['dropped'])} attribute(s) with a positive no-SOR signal.")
    print(f"Kept and flagged {d['attributes_unmatched']} attribute(s) absent from the mapping.")
    print(f"Cascade-removed {len(result['cascade']['removed_schemas'])} emptied object(s) and "
          f"{len(result['cascade']['removed_properties'])} reference(s).")
    for line in result["poly_log"]:
        print(f"  poly:   {line}")
    for line in result["dedupe_log"]:
        print(f"  dedupe: {line}")
    if result["cascade"]["empty_body_schemas"]:
        print("  WARNING: operation body schema(s) left empty: "
              + ", ".join(result["cascade"]["empty_body_schemas"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
