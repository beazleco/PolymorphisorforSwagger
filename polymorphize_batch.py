#!/usr/bin/env python3
"""
polymorphize_batch — several mapping workbooks in one run.

Version 6.4.

Under the workbook-authoritative contract one workbook already fans out to many
endpoints, so a batch is a list of workbooks rather than a list of jobs with
options. The manifest is correspondingly small.

Manifest columns, header row, case-insensitive, unknown columns ignored
======================================================================

===============  =========================================================
workbook         the mapping workbook                            [required]
out_dir          where its specifications go   (default: <workbook>_openapi)
name             a label for the roll-up      (default: the workbook's name)
level            lenient | strict                          (default: lenient)
format           yaml | json                                  (default: yaml)
api_version      stamped into info.version                   (default: 1.0.0)
title            info.title of the merged specification
split            true writes one file per sheet             (default: false)
sor              SOR specification(s) or folder, ';'-separated
===============  =========================================================

Relative paths resolve against the manifest's own folder unless ``--base-dir``
says otherwise.

Outputs
=======

Each workbook's specifications go to its own folder, alongside that run's
``generation_report.md`` and a ``.FAILED.md`` for every endpoint that failed.
The batch itself writes ``batch_summary.csv`` and ``batch_index.html``.

Exit codes match the single-workbook CLI: 0 when every endpoint of every
workbook generated, 6 when some generated and some failed, 7 when none did.
"""

from __future__ import annotations

import argparse
import concurrent.futures as cf
import csv
import html as _html
import os
import sys
import traceback

import polymorphize_generate as gen

__version__ = "6.4"

EXIT_OK, EXIT_UNREADABLE, EXIT_PARTIAL, EXIT_TOTAL_FAILURE = 0, 2, 6, 7

DEFAULTS = {"level": "lenient", "format": "yaml", "api_version": "1.0.0"}

TRUE = {"1", "true", "yes", "y", "on", "t"}


def _norm(text):
    return "".join(ch for ch in str(text or "").lower() if ch.isalnum())


def read_manifest(path):
    """Rows of the manifest as dictionaries keyed by normalised column name."""
    rows = []
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        header = None
        for raw in ws.iter_rows(values_only=True):
            values = ["" if c is None else str(c).strip() for c in raw]
            if header is None:
                if any(values):
                    header = [_norm(v) for v in values]
                continue
            if any(values):
                rows.append(dict(zip(header, values)))
        wb.close()
    else:
        with open(path, newline="", encoding="utf-8-sig") as fh:
            reader = csv.reader(fh)
            header = None
            for values in reader:
                if header is None:
                    header = [_norm(v) for v in values]
                    continue
                if any(v.strip() for v in values):
                    rows.append(dict(zip(header, [v.strip() for v in values])))
    return rows


def _resolve(base_dir, value):
    if not value:
        return ""
    return value if os.path.isabs(value) else os.path.join(base_dir, value)


def plan(manifest_path, base_dir=None):
    """The manifest as a list of jobs, with defaults filled in."""
    base = base_dir or os.path.dirname(os.path.abspath(manifest_path))
    jobs = []
    for i, row in enumerate(read_manifest(manifest_path), start=2):
        workbook = _resolve(base, row.get("workbook", ""))
        if not workbook:
            raise ValueError("manifest row %d has no workbook" % i)
        out_dir = _resolve(base, row.get("outdir", "")) or \
            os.path.splitext(workbook)[0] + "_openapi"
        jobs.append({
            "row": i,
            "name": row.get("name") or os.path.basename(workbook),
            "workbook": workbook,
            "out_dir": out_dir,
            "level": (row.get("level") or DEFAULTS["level"]).lower(),
            "format": (row.get("format") or DEFAULTS["format"]).lower(),
            "api_version": row.get("apiversion") or DEFAULTS["api_version"],
            "title": row.get("title") or None,
            "split": (row.get("split") or "").strip().lower() in TRUE,
            # Resolved against the manifest's folder, like every other path
            # in the row. Left unresolved, a relative SOR path silently fails
            # to load and the run quietly stops verifying.
            "sor": [_resolve(base, x.strip())
                    for x in (row.get("sor") or "").split(";")
                    if x.strip()] or None,
        })
    return jobs


def run_job(job):
    """One workbook. Never raises: a broken job is reported, not fatal."""
    result = dict(job)
    result.update({"status": "ok", "ok": 0, "failed": 0, "skipped": 0,
                   "failed_sheets": [], "error": "", "merged": {},
                   "sor_verified": False, "showcase": ""})
    try:
        out = gen.run(job["workbook"], job["out_dir"],
                      strict=(job["level"] == "strict"),
                      version=job["api_version"], fmt=job["format"],
                      split=job["split"], title=job["title"],
                      sor=job["sor"], log=lambda *_a: None)
    except Exception as exc:                                   # noqa: BLE001
        result["status"] = "unreadable"
        result["error"] = "%s: %s" % (type(exc).__name__, exc)
        result["traceback"] = traceback.format_exc()
        return result
    result["ok"] = len(out.ok)
    result["failed"] = len(out.failed)
    result["skipped"] = len(out.skipped)
    if out.merged is not None and out.merged.errors:
        result["status"] = "merge-failed"
        result["error"] = "; ".join(f.what for f in out.merged.errors)[:400]
    result["sor_verified"] = bool(out.sor_verified)
    result["showcase"] = out.showcase_path
    result["merged"] = (out.merged.stats
                        if out.merged is not None and out.merged.document
                        else {})
    result["failed_sheets"] = [
        (s.sheet, next((f.code for f in s.findings if f.severity == "error"), ""),
         next((f.what for f in s.findings if f.severity == "error"), ""))
        for s in out.failed]
    if out.failed and out.ok:
        result["status"] = "partial"
    elif out.failed:
        result["status"] = "failed"
    return result


def write_summary(results, out_path):
    cols = ["name", "workbook", "out_dir", "level", "status", "ok", "failed",
            "skipped", "sor_verified", "showcase", "error"]
    with open(out_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore")
        writer.writeheader()
        for r in results:
            writer.writerow(r)


BADGE = {"ok": ("#e8f4ea", "#1e6b34", "all generated"),
         "merge-failed": ("#fdeaea", "#a11b1b", "merge failed"),
         "partial": ("#fdf6e3", "#8a6100", "some failed"),
         "failed": ("#fdeaea", "#a11b1b", "none generated"),
         "unreadable": ("#fdeaea", "#a11b1b", "could not be read")}


def write_index(results, out_path):
    e = _html.escape
    total_ok = sum(r["ok"] for r in results)
    total_failed = sum(r["failed"] for r in results)
    parts = ["<!doctype html><meta charset='utf-8'>",
             "<title>SOR Polymorphizer batch</title>",
             "<style>",
             "body{font:14px/1.55 -apple-system,Segoe UI,Roboto,sans-serif;",
             "margin:0;padding:32px;background:#f4f6fa;color:#1a1a1a}",
             "h1{font-size:21px;color:#1f3864;margin:0 0 4px}",
             "p.sub{color:#5a6472;margin:0 0 24px}",
             "table{border-collapse:collapse;width:100%;background:#fff;",
             "box-shadow:0 1px 2px rgba(0,0,0,.06)}",
             "th,td{padding:9px 12px;border-bottom:1px solid #e4e9f2;",
             "text-align:left;vertical-align:top}",
             "th{background:#1f3864;color:#fff;font-weight:600;font-size:12px}",
             "td.n{text-align:right;font-variant-numeric:tabular-nums}",
             "span.b{padding:2px 9px;border-radius:10px;font-size:12px;",
             "font-weight:600;white-space:nowrap}",
             "details{margin-top:6px}summary{cursor:pointer;color:#a11b1b}",
             "code{font:12px ui-monospace,Consolas,monospace;color:#5a6472}",
             "</style>",
             "<h1>SOR Polymorphizer batch</h1>",
             "<p class='sub'>%d workbook%s. %d endpoint%s generated, "
             "<strong>%d failed</strong>.</p>"
             % (len(results), "" if len(results) == 1 else "s",
                total_ok, "" if total_ok == 1 else "s", total_failed),
             "<table><tr><th>Workbook</th><th>Status</th><th>Generated</th>",
             "<th>Failed</th><th>Support sheets</th><th>Output</th></tr>"]
    for r in results:
        bg, fg, label = BADGE.get(r["status"], ("#eee", "#333", r["status"]))
        cell = ["<td><strong>%s</strong><br><code>%s</code>"
                % (e(r["name"]), e(r["workbook"]))]
        if r["failed_sheets"]:
            cell.append("<details><summary>%d endpoint%s failed</summary><ul>"
                        % (len(r["failed_sheets"]),
                           "" if len(r["failed_sheets"]) == 1 else "s"))
            for sheet, code, what in r["failed_sheets"]:
                cell.append("<li><strong>%s</strong> &mdash; %s %s</li>"
                            % (e(sheet), e(code), e(what)))
            cell.append("</ul></details>")
        if r["error"]:
            cell.append("<br><code>%s</code>" % e(r["error"]))
        cell.append("</td>")
        parts.append("<tr>%s<td><span class='b' style='background:%s;color:%s'>"
                     "%s</span></td><td class='n'>%d</td><td class='n'>%d</td>"
                     "<td class='n'>%d</td><td><code>%s</code></td></tr>"
                     % ("".join(cell), bg, fg, e(label), r["ok"], r["failed"],
                        r["skipped"], e(r["out_dir"])))
    parts.append("</table>")
    with open(out_path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(parts))


def main(argv=None):
    p = argparse.ArgumentParser(
        prog="polymorphize_batch",
        description="Generate specifications from several mapping workbooks.")
    p.add_argument("--manifest", required=True, help="CSV or Excel manifest")
    p.add_argument("--base-dir", help="resolve relative paths against this "
                                      "folder instead of the manifest's")
    p.add_argument("--out-dir", default=".",
                   help="where batch_summary.csv and batch_index.html go")
    p.add_argument("--jobs", type=int, default=1,
                   help="how many workbooks to process at once")
    p.add_argument("--sor", action="append", metavar="SPEC", default=None,
                   help="SOR specification(s) for every workbook, overriding "
                        "the manifest")
    p.add_argument("--split", action="store_true",
                   help="write one specification per sheet for every workbook, "
                        "overriding the manifest")
    p.add_argument("--version", action="version",
                   version="SOR Polymorphizer batch " + __version__)
    args = p.parse_args(argv)

    try:
        jobs = plan(args.manifest, args.base_dir)
    except Exception as exc:                                   # noqa: BLE001
        print("Could not read the manifest: %s: %s"
              % (type(exc).__name__, exc), file=sys.stderr)
        return EXIT_UNREADABLE
    if not jobs:
        print("The manifest lists no workbooks.", file=sys.stderr)
        return EXIT_UNREADABLE

    if args.split:
        for job in jobs:
            job["split"] = True
    if args.sor:
        for job in jobs:
            job["sor"] = args.sor
    print("%d workbook(s) to process" % len(jobs))
    results = []
    if args.jobs > 1:
        with cf.ThreadPoolExecutor(max_workers=args.jobs) as pool:
            for r in pool.map(run_job, jobs):
                results.append(r)
                print("  %-9s %-40s %d ok, %d failed"
                      % (r["status"], r["name"][:40], r["ok"], r["failed"]))
    else:
        for job in jobs:
            r = run_job(job)
            results.append(r)
            print("  %-9s %-40s %d ok, %d failed"
                  % (r["status"], r["name"][:40], r["ok"], r["failed"]))

    os.makedirs(args.out_dir, exist_ok=True)
    write_summary(results, os.path.join(args.out_dir, "batch_summary.csv"))
    write_index(results, os.path.join(args.out_dir, "batch_index.html"))
    print("")
    print("Wrote batch_summary.csv and batch_index.html to %s"
          % os.path.abspath(args.out_dir))

    total_ok = sum(r["ok"] for r in results)
    total_failed = sum(r["failed"] for r in results) + \
        sum(1 for r in results
            if r["status"] in ("unreadable", "merge-failed"))
    if total_failed:
        print("")
        print("=" * 72)
        print("%d ENDPOINT%s FAILED ACROSS THE BATCH"
              % (total_failed, "" if total_failed == 1 else "S"))
        for r in results:
            for sheet, code, what in r["failed_sheets"]:
                print("  %-30s %-30s %s %s"
                      % (r["name"][:30], sheet[:30], code, what[:80]))
            if r["status"] in ("unreadable", "merge-failed"):
                print("  %-30s %s: %s"
                      % (r["name"][:30], r["status"], r["error"][:100]))
        if not r["failed_sheets"] and r["status"] in ("unreadable",
                                                      "merge-failed"):
            print("  %-30s %s: %s"
                  % (r["name"][:30], r["status"], r["error"][:100]))
        print("=" * 72)
    if total_failed and total_ok:
        return EXIT_PARTIAL
    if total_failed:
        return EXIT_TOTAL_FAILURE
    return EXIT_OK


if __name__ == "__main__":
    raise SystemExit(main())
