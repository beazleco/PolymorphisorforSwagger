"""
polymorphize_cli — command line entry points.

Version 6.3.

Two subcommands:

``check``
    Validate a workbook without writing anything.

``generate``
    Validate, then write one OpenAPI specification per operation sheet.

Exit codes
==========

===  ==========================================================
0    every operation sheet generated
2    the workbook could not be opened
6    at least one endpoint failed, the rest generated
7    no endpoint generated at all
===  ==========================================================

Code 6 is deliberately not zero. A partial run must not pass a build step
silently just because most of the workbook was fine.
"""

from __future__ import annotations

import argparse
import os
import sys

import polymorphize_generate as gen
import polymorphize_validate as val

__version__ = "6.3"

EXIT_OK = 0
EXIT_UNREADABLE = 2
EXIT_PARTIAL = 6
EXIT_TOTAL_FAILURE = 7

RULE = "=" * 72


def _banner(stream, title):
    print(RULE, file=stream)
    print(title, file=stream)
    print(RULE, file=stream)


def _report_failures(reports_or_specs, stream):
    """The failure block. Loud, at the end, where the eye lands."""
    failed = [x for x in reports_or_specs if x.status == "failed"]
    if not failed:
        return
    print("", file=stream)
    _banner(stream, "%d ENDPOINT%s FAILED" % (len(failed),
                                              "" if len(failed) == 1 else "S"))
    for item in failed:
        print("", file=stream)
        print("  %s" % item.sheet, file=stream)
        for f in item.findings:
            if f.severity != "error":
                continue
            print("    %s at %s" % (f.code, f.where), file=stream)
            for label, body in (("found", f.what), ("expected", f.expected),
                                ("fix", f.fix)):
                for i, line in enumerate(_wrap(body, 62)):
                    print("      %-9s %s" % (label if i == 0 else "", line),
                          file=stream)
    print("", file=stream)
    print("  No specification was written for the endpoints above.", file=stream)
    print(RULE, file=stream)


def _wrap(body, width):
    out, line = [], ""
    for word in str(body).replace("\n", " ").split():
        if len(line) + len(word) + 1 > width:
            out.append(line)
            line = word
        else:
            line = (line + " " + word).strip()
    if line:
        out.append(line)
    return out or [""]


def cmd_check(args):
    try:
        reports = val.validate_workbook(args.workbook,
                                        strict=(args.level == "strict"),
                                        sor=args.sor)
    except Exception as exc:                                   # noqa: BLE001
        print("Could not read %s: %s: %s"
              % (args.workbook, type(exc).__name__, exc), file=sys.stderr)
        return EXIT_UNREADABLE

    s = val.summarise(reports)
    print("Workbook: %s" % args.workbook)
    print("Level:    %s" % args.level)
    print("SOR:      %s" % (", ".join(args.sor) if args.sor
                            else "not supplied, so the workbook is taken on "
                                 "trust"))
    print("")
    for r in reports:
        if r.status == "skipped":
            print("  skip    %-34s support sheet" % r.sheet)
        elif r.status == "failed":
            print("  FAILED  %-34s %d error(s), %d warning(s)"
                  % (r.sheet, len(r.errors), len(r.warnings)))
        else:
            print("  ok      %-34s %-3s %d warning(s)"
                  % (r.sheet, r.pattern or "-", len(r.warnings)))
    print("")
    print("%d operation sheet(s): %d would generate, %d would fail. "
          "%d error(s), %d warning(s)."
          % (s["operations"], s["ok"], s["failed"], s["errors"], s["warnings"]))

    if args.report:
        text = val.format_report(args.workbook, reports, args.level == "strict")
        with open(args.report, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(text)
        print("Wrote %s" % args.report)

    if args.verbose:
        for r in reports:
            for f in r.findings:
                print("  %-7s %-5s %-28s %s"
                      % (f.severity, f.code, f.where, f.what.replace("\n", " ")))

    _report_failures(reports, sys.stderr)
    if s["failed"] and s["ok"]:
        return EXIT_PARTIAL
    if s["failed"]:
        return EXIT_TOTAL_FAILURE
    return EXIT_OK


def cmd_generate(args):
    try:
        out = gen.run(args.workbook, args.out_dir, strict=args.strict,
                      version=args.api_version, fmt=args.format,
                      split=args.split, merged_name=args.name,
                      title=args.title, sor=args.sor,
                      showcase=not args.no_showcase)
    except Exception as exc:                                   # noqa: BLE001
        print("Could not process %s: %s: %s"
              % (args.workbook, type(exc).__name__, exc), file=sys.stderr)
        return EXIT_UNREADABLE

    print("")
    print("Generated %d of %d endpoint(s) into %s"
          % (len(out.ok), len(out.ok) + len(out.failed),
             os.path.abspath(args.out_dir)))
    if not args.sor:
        print("NOTE: no SOR specification was supplied, so every SOR field "
              "name in the workbook was taken on trust.")
    elif out.sor_verified:
        ss = out.sor_index.summary()
        print("Verified against %d SOR file(s), %d operation(s)."
              % (ss["files"], ss["operations"]))
    if out.showcase_path:
        print("Showcase: %s" % out.showcase_path)
    if out.split:
        print("Wrote one specification per sheet (--split).")
    elif out.merged is not None and out.merged.document is not None:
        m = out.merged.stats
        print("Wrote one specification: %d operations, %d schemas, %d groups "
              "hoisted, %d specialised across operations."
              % (m["operations"], m["schemas"], m["hoisted_groups"],
                 m["split_groups"]))
    if out.failed:
        print("%d FAILED. See generation_report.md and the .FAILED.md files."
              % len(out.failed))
    if out.merged is not None and out.merged.errors:
        print("")
        _banner(sys.stderr, "THE MERGE FAILED")
        for f in out.merged.errors:
            print("  %s at %s" % (f.code, f.where), file=sys.stderr)
            for label, body in (("found", f.what), ("expected", f.expected),
                                ("fix", f.fix)):
                for i, line in enumerate(_wrap(body, 62)):
                    print("    %-9s %s" % (label if i == 0 else "", line),
                          file=sys.stderr)
        print("  No single specification was written.", file=sys.stderr)
        print(RULE, file=sys.stderr)
        return EXIT_TOTAL_FAILURE

    _report_failures(out.specs, sys.stderr)
    if out.failed and out.ok:
        return EXIT_PARTIAL
    if out.failed:
        return EXIT_TOTAL_FAILURE
    return EXIT_OK


def cmd_template(args):
    import polymorphize_template as tmpl
    tmpl.write_template(args.path)
    print("Wrote %s" % args.path)
    return EXIT_OK


def build_parser():
    p = argparse.ArgumentParser(
        prog="polymorphize",
        description="Generate one OpenAPI specification per operation sheet of "
                    "an SOR mapping workbook.")
    p.add_argument("--version", action="version", version="SOR Polymorphizer " + __version__)
    sub = p.add_subparsers(dest="command", required=True)

    c = sub.add_parser("check", help="validate a workbook without writing anything")
    c.add_argument("workbook")
    c.add_argument("--level", choices=("lenient", "strict"), default="lenient",
                   help="lenient reports only what stops an endpoint generating; "
                        "strict adds the full template contract")
    c.add_argument("--report", help="write the validation report to this path")
    c.add_argument("--sor", action="append", metavar="SPEC", default=None,
                   help="a System of Record specification, or a folder of "
                        "them. Repeatable. Verifies every SOR field name in "
                        "the workbook against the SOR endpoint its sheet names")
    c.add_argument("-v", "--verbose", action="store_true",
                   help="list every finding on stdout")
    c.set_defaults(func=cmd_check)

    g = sub.add_parser("generate", help="write one specification per operation sheet")
    g.add_argument("workbook")
    g.add_argument("out_dir")
    g.add_argument("--format", choices=("yaml", "json"), default="yaml")
    g.add_argument("--api-version", default="1.0.0",
                   help="the version stamped into info.version")
    g.add_argument("--strict", action="store_true",
                   help="treat every warning as a failure")
    g.add_argument("--split", action="store_true",
                   help="write one specification per sheet instead of one for "
                        "the whole workbook. Polymorphism across operations "
                        "cannot apply in this mode, because each sheet gets its "
                        "own schema namespace. Intended for debugging a single "
                        "endpoint")
    g.add_argument("--name", default="openapi",
                   help="base file name of the merged specification "
                        "(default: openapi)")
    g.add_argument("--title", help="info.title of the merged specification")
    g.add_argument("--sor", action="append", metavar="SPEC", default=None,
                   help="a System of Record specification, or a folder of "
                        "them. Repeatable. Every SOR field name in the workbook "
                        "is looked up in the SOR endpoint its sheet names, and "
                        "an element whose field is not there is excluded from "
                        "the interface. Without this the workbook is taken on "
                        "trust")
    g.add_argument("--no-showcase", action="store_true",
                   help="skip the HTML showcase report")
    g.set_defaults(func=cmd_generate)

    t = sub.add_parser("template", help="write the mapping workbook template")
    t.add_argument("path", nargs="?", default="SOR_mapping_template.xlsx")
    t.set_defaults(func=cmd_template)
    return p


def main(argv=None):
    args = build_parser().parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
