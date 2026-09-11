"""
polymorphize_validate — workbook validation against the Level-format contract.

Version 6.8.

Validation runs the real reader and the real generator rather than a parallel
set of rules, so a workbook that validates clean is a workbook that generates,
and every message an analyst sees names the sheet, the cell, what was found,
what was expected and what to do about it.

Two levels
==========

``lenient``
    Only what stops an endpoint being generated. This is what the run itself
    enforces, and a lenient pass means every endpoint produced a
    specification.

``strict``
    The full template contract, including the documentation an analyst is
    expected to supply: a use case, a service domain, a behaviour qualifier,
    a description on every attribute and an example on every mandatory one.
    Strict findings never block generation on their own; they are the house
    standard for a workbook that is going to be handed to a developer.

Code families
=============

===== ==========================================================
L     layout: the header row and the Level columns
B     banner: the labelled cells beneath the header row
S     sections: Request Body and Response Body
A     attribute rows: names, nesting, duplicates
T     Data Type cells
U     Usage cells
M     SOR columns
V     variants: ``requestVariant`` against the SOR columns
E     emission: a section that would publish an empty object
G     generation: the method and the body
D     documentation, strict only
X     faults in the tool itself
===== ==========================================================
"""

from __future__ import annotations

import os
from collections import Counter, OrderedDict

import polymorphize_generate as gen
import polymorphize_workbook as wbk
from polymorphize_workbook import Finding, cell_ref, text

__version__ = "6.8"

#: Banner labels a strict workbook is expected to fill in.
REQUIRED_BANNER = OrderedDict([
    ("use_case", "Use Case"),
    ("service_domain", "Service Domain"),
    ("bian_endpoint", "Equivalent BIAN API Endpoint"),
    ("sor_endpoint", "SOR API Endpoint"),
])

RECOMMENDED_BANNER = OrderedDict([
    ("behaviour_qualifier", "BQ, alongside the Service Domain"),
    ("business_endpoint", "Proposed Business API Endpoint"),
])


# --------------------------------------------------------------------------- #
# Strict-only checks
# --------------------------------------------------------------------------- #


def _banner_checks(result, findings):
    layout = result.layout
    label_col = layout.label_cols[0] if layout.label_cols else 0
    for field_name, label in REQUIRED_BANNER.items():
        if not text(result.banner.get(field_name, "")):
            findings.append(Finding(
                "B001", "warning", result.sheet,
                cell_ref(layout.header_row + 2, label_col),
                "%s is missing or empty" % label,
                "%s stated in the banner beneath the header row" % label,
                "add a row in column %s labelled %r with the value in the next "
                "column" % (wbk.col_letter(label_col), label + ":"),
            ))
    for field_name, label in RECOMMENDED_BANNER.items():
        if not text(result.banner.get(field_name, "")):
            findings.append(Finding(
                "B002", "warning", result.sheet,
                cell_ref(layout.header_row + 2, label_col),
                "%s is not stated" % label,
                "%s stated in the banner" % label,
                "add %s so the generated specification can carry it" % label,
            ))
    extra = result.banner.get("_extra") or []
    for label, _value in extra:
        findings.append(Finding(
            "B003", "warning", result.sheet,
            cell_ref(layout.header_row + 2, label_col),
            "%r is not a recognised banner label, so its value was ignored" % label,
            "one of " + ", ".join(sorted(set(
                list(REQUIRED_BANNER.values()) +
                ["Proposed Business API Endpoint", "Method"]))),
            "correct the spelling of %r, or move the value out of the banner" % label,
        ))


def _documentation_checks(result, findings):
    layout = result.layout
    desc_col = layout.roles.get("description")
    ex_col = layout.roles.get("example")
    for section in (result.request, result.response):
        if section is None or section.root is None:
            continue

        def walk(node, is_root=False):
            if not is_root and not node.description:
                findings.append(Finding(
                    "D001", "warning", result.sheet,
                    cell_ref(node.row, desc_col if desc_col is not None else 0),
                    "%r has no description" % node.name,
                    "a description on every attribute",
                    "describe what %r means to a consumer of the API" % node.name,
                ))
            if (not node.children and node.usage == wbk.USAGE_REQUIRED
                    and not node.example):
                findings.append(Finding(
                    "D002", "warning", result.sheet,
                    cell_ref(node.row, ex_col if ex_col is not None else 0),
                    "%r is Mandatory but has no example value" % node.name,
                    "an example for every mandatory attribute",
                    "put a realistic value in the Dummy Value column for %r"
                    % node.name,
                ))
            for c in node.children:
                walk(c)
        walk(section.root, is_root=True)


def _column_checks(result, findings):
    layout = result.layout
    for role, label in (("type", "Data Type"), ("usage", "Usage"),
                        ("description", "Descriptions"), ("example", "Dummy Value")):
        if role not in layout.roles:
            findings.append(Finding(
                "L003", "warning", result.sheet,
                cell_ref(layout.header_row + 1, layout.level_cols[-1] + 1),
                "there is no %s column" % label,
                "a %s column on the header row" % label,
                "add a %s header to the right of the Level columns" % label,
            ))
    seen = Counter(c.endpoint for c in layout.sor_cols)
    for endpoint, n in seen.items():
        if n > 1:
            findings.append(Finding(
                "M002", "error", result.sheet,
                cell_ref(layout.header_row + 1,
                         next(c.index for c in layout.sor_cols
                              if c.endpoint == endpoint)),
                "%d SOR columns name the same endpoint %r" % (n, endpoint),
                "one column per SOR endpoint",
                "give each SOR column its own endpoint in the header, for example "
                "SOR /v1/card/details",
            ))
    if len(layout.sor_cols) > 1:
        for c in layout.sor_cols:
            if not c.endpoint:
                findings.append(Finding(
                    "M003", "error", result.sheet,
                    cell_ref(layout.header_row + 1, c.index),
                    "the header reads %r, which does not name an endpoint, and this "
                    "sheet has more than one SOR column" % c.header,
                    "SOR followed by the endpoint path",
                    "rename it to SOR followed by the endpoint, for example "
                    "SOR /v1/account/info",
                ))


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #


class SheetReport:
    """Everything known about one worksheet after validation."""

    __slots__ = ("sheet", "status", "pattern", "findings", "result", "spec")

    def __init__(self, sheet):
        self.sheet = sheet
        self.status = "ok"
        self.pattern = ""
        self.findings = []
        self.result = None
        self.spec = None

    @property
    def errors(self):
        return [f for f in self.findings if f.severity == "error"]

    @property
    def warnings(self):
        return [f for f in self.findings if f.severity == "warning"]


def validate_workbook(path, strict=False, sor=None):
    """Validate every operation sheet. Returns a list of :class:`SheetReport`.

    ``sor`` is one or more System of Record specifications, or a folder of
    them. When given, every SOR field name is verified against the SOR
    endpoint its sheet names, so the report says which elements would be
    excluded for want of a real SOR field rather than for want of a workbook
    entry.
    """
    reports = []
    import polymorphize_classify as cls
    verdict = cls.classify(path)
    if verdict.fmt == cls.MAPPING:
        import polymorphize_mapping as mapfmt
        results = mapfmt.read_workbook(path)
    else:
        results = wbk.read_workbook(path)
    index = None
    if sor:
        import polymorphize_sor as sormod
        index = sormod.SorIndex.from_paths(sor)
        if not index.empty:
            sormod.verify_workbook(results, index)
    for result in results:
        rep = SheetReport(result.sheet)
        rep.result = result
        if result.status == "skipped":
            rep.status = "skipped"
            reports.append(rep)
            continue

        rep.findings.extend(result.findings)

        # The generator is the authority on whether a sheet produces a
        # specification, so run it rather than guessing.
        spec = gen.generate_sheet(result)
        rep.spec = spec
        rep.pattern = spec.pattern
        known = {(f.code, f.where, f.what) for f in rep.findings}
        for f in spec.findings:
            if (f.code, f.where, f.what) not in known:
                rep.findings.append(f)
        if spec.status == "failed":
            rep.status = "failed"

        if strict and result.layout is not None:
            _banner_checks(result, rep.findings)
            # The column checks name Level headings, so they apply to the
            # Level format only. The field mapping reader raises its own
            # column findings, F002 and F004, at read time.
            if verdict.fmt != cls.MAPPING:
                _column_checks(result, rep.findings)
            _documentation_checks(result, rep.findings)

        reports.append(rep)
    return reports


def has_blocking_errors(reports):
    """True when at least one endpoint failed."""
    return any(r.status == "failed" for r in reports)


def summarise(reports):
    """Counts for a status line."""
    ok = [r for r in reports if r.status == "ok"]
    failed = [r for r in reports if r.status == "failed"]
    skipped = [r for r in reports if r.status == "skipped"]
    return {
        "sheets": len(reports),
        "operations": len(ok) + len(failed),
        "ok": len(ok),
        "failed": len(failed),
        "skipped": len(skipped),
        "errors": sum(len(r.errors) for r in reports),
        "warnings": sum(len(r.warnings) for r in reports),
    }


def findings_rows(reports):
    """Flat rows for a table or a Treeview: severity, code, sheet, cell, text."""
    rows = []
    for r in reports:
        for f in sorted(r.findings, key=lambda x: (x.severity != "error", x.code)):
            rows.append((f.severity, f.code, f.sheet, f.cell, f.what, f.expected, f.fix))
    return rows


def format_report(path, reports, strict, sor=None, input_format=None):
    """The validation report, failures first."""
    s = summarise(reports)
    L = [x for x in ["# Workbook validation report", "",
         "Workbook: `%s`" % os.path.basename(path),
         "",
         "Level: **%s**" % ("strict" if strict else "lenient"),
         "",
         ("Input format: **%s**" % input_format) if input_format else "",
         "" if input_format else None,
         ("SOR specification: **%s**" % ", ".join(
             os.path.basename(x) for x in sor)) if sor else
         "SOR specification: **none supplied**, so every SOR field name in the "
         "workbook was taken on trust.",
         "",
         "%d operation sheet%s, **%d would generate**, **%d would fail**. "
         "%d support sheet%s skipped."
         % (s["operations"], "" if s["operations"] == 1 else "s",
            s["ok"], s["failed"], s["skipped"],
            "" if s["skipped"] == 1 else "s"),
         "",
         "%d error%s, %d warning%s."
         % (s["errors"], "" if s["errors"] == 1 else "s",
            s["warnings"], "" if s["warnings"] == 1 else "s"),
         ""] if x is not None]

    failed = [r for r in reports if r.status == "failed"]
    if failed:
        L += ["## Endpoints that would fail", "",
              "Each of these needs a correction in the workbook before a "
              "specification can be generated for it.", ""]
        for r in failed:
            L += ["### %s" % r.sheet, ""]
            for f in r.errors:
                L += ["- **%s at %s**" % (f.code, f.where),
                      "  - Found: %s" % f.what,
                      "  - Expected: %s" % f.expected,
                      "  - Fix: %s" % f.fix]
            L.append("")

    ok = [r for r in reports if r.status == "ok"]
    if ok:
        L += ["## Endpoints that would generate", "",
              "| Sheet | Pattern | Warnings |", "|---|---|---|"]
        for r in ok:
            L.append("| %s | %s | %d |" % (r.sheet, r.pattern or "-", len(r.warnings)))
        L.append("")

    warn = [(r, f) for r in reports for f in r.warnings]
    if warn:
        L += ["## Warnings", "",
              "| Code | Where | Found | Fix |", "|---|---|---|---|"]
        for r, f in warn:
            L.append("| %s | `%s` | %s | %s |"
                     % (f.code, f.where, f.what.replace("|", "/").replace("\n", " "),
                        f.fix.replace("|", "/").replace("\n", " ")))
        L.append("")

    counts = Counter(f.code for r in reports for f in r.findings)
    if counts:
        L += ["## Findings by code", "", "| Code | Count |", "|---|---|"]
        for code, n in sorted(counts.items()):
            L.append("| %s | %d |" % (code, n))
        L.append("")
    return "\n".join(L)


__all__ = ["SheetReport", "findings_rows", "format_report", "has_blocking_errors",
           "summarise", "validate_workbook"]
