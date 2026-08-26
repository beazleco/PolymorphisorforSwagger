#!/usr/bin/env python3
"""
polymorphize_batch.py  (v5.1)
=============================

Run the SOR refactor over MANY swaggers in one go, driven by a manifest (CSV or
Excel) that lists one job per row.

Each row needs at least ``swagger`` and ``mapping``; everything else is
optional and falls back to a sensible default. Relative paths in the manifest
are resolved against the manifest's own folder (override with --base-dir).

Manifest columns (header row, case-insensitive; unknown columns ignored)
------------------------------------------------------------------------
    name              friendly label for the report (default: swagger filename)
    swagger           input OpenAPI file                             [required]
    mapping           SOR mapping spreadsheet (.xlsx)                [required]
    out               output file (default: <swagger>_polymorphic.yaml)
    report            optional Markdown report path
    html_report       optional HTML change report (default: <out>_changes.html)
    sor_swagger       SOR API YAML(s) to validate mapped fields, ';'-separated
    require_sor_field eliminate a mapped attribute absent from the SOR YAML
    on_unmatched      keep | drop  — what to do with attributes the mapping
                      does not mention                        (default: keep)
    discriminator     add discriminator to base schemas         (default: true)
    auto_poly         auto-detect polymorphism families         (default: true)
    auto_dedupe       auto de-duplicate inline objects          (default: true)
    dedupe_strictness deep | shape                              (default: deep)
    disc_prop         discriminator property name      (default: resultContext)
    protected         schemas never pruned            (default: ErrorResponse)
    poly              explicit group(s), ';'-separated (overrides auto_poly)
    dedupe            explicit de-dupe(s), ';'-separated
    min_match_rate    abort below this mapping match rate       (default: 0.25)
    max_total_loss    abort above this attribute loss           (default: 0.75)
    guards            enforce the two thresholds above          (default: true)
    assertions        enforce post-run output assertions        (default: true)
    auto_layout       detect the spreadsheet layout             (default: true)
    level_first_col / level_last_col / dtype_col / sor_col / not_used_text
                      fixed spreadsheet layout, used only when auto_layout is
                      false

Outputs
-------
    <each job's out + html>   the transformed swaggers
    batch_summary.csv         one row per job, including the mapping match rate
    batch_index.html          linked roll-up you can open in a browser

A job whose mapping barely matched, or which would strip most of the
specification, is reported as ABORTED and writes nothing. That is deliberate:
a silently emptied specification is worse than a failed job.

Usage
-----
    python polymorphize_batch.py --manifest jobs.csv --jobs 4
"""
import argparse
import concurrent.futures as cf
import csv
import html as _html
import os
import sys

import openpyxl

import polymorphize_core as core

BOOL_TRUE = {"1", "true", "yes", "y", "on", "t"}


# --------------------------------------------------------------------------- #
# Manifest loading
# --------------------------------------------------------------------------- #
def load_manifest(path):
    rows = []
    if path.lower().endswith((".xlsx", ".xlsm")):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        it = ws.iter_rows(values_only=True)
        headers = [str(h).strip().lower() if h is not None else "" for h in next(it)]
        for r in it:
            if not any(c not in (None, "") for c in r):
                continue
            rows.append({h: v for h, v in zip(headers, r) if h})
    else:
        import io
        text, _enc = core.read_text(path)
        for row in csv.DictReader(io.StringIO(text, newline="")):
            rows.append({(k or "").strip().lower(): v for k, v in row.items()})
    return rows


def _split_multi(text):
    if not text:
        return []
    parts = []
    for chunk in str(text).replace("\n", ";").split(";"):
        chunk = chunk.strip()
        if chunk:
            parts.append(chunk)
    return parts


def build_job(row, base_dir, idx):
    def g(*names, default=None):
        for n in names:
            if n in row and row[n] is not None and str(row[n]).strip() != "":
                return str(row[n]).strip()
        return default

    def gb(*names, default=False):
        v = g(*names)
        return default if v is None else (v.strip().lower() in BOOL_TRUE)

    def gi(name, default):
        v = g(name)
        try:
            return int(float(v)) if v is not None else default
        except ValueError:
            return default

    def gf(name, default):
        v = g(name)
        try:
            return float(v) if v is not None else default
        except ValueError:
            return default

    def resolve(p):
        if not p:
            return None
        return p if os.path.isabs(p) else os.path.normpath(os.path.join(base_dir, p))

    swagger = resolve(g("swagger", "input"))
    mapping = resolve(g("mapping", "spreadsheet"))
    sor_swaggers = [resolve(p) for p in _split_multi(g("sor_swagger", "sor_yaml", "sor"))]
    out = resolve(g("out", "output"))
    if not out and swagger:
        out = os.path.splitext(swagger)[0] + "_polymorphic.yaml"
    report = resolve(g("report", "md_report"))
    html = resolve(g("html_report", "html"))
    if not html and out:
        html = os.path.splitext(out)[0] + "_changes.html"

    sheet_cfg = {
        "level_first_col": gi("level_first_col", core.DEFAULT_SHEET_CFG["level_first_col"]),
        "level_last_col": gi("level_last_col", core.DEFAULT_SHEET_CFG["level_last_col"]),
        "dtype_col": gi("dtype_col", core.DEFAULT_SHEET_CFG["dtype_col"]),
        "sor_col": gi("sor_col", core.DEFAULT_SHEET_CFG["sor_col"]),
        "not_used_text": g("not_used_text", default=core.DEFAULT_SHEET_CFG["not_used_text"]),
        "auto_layout": gb("auto_layout", default=True),
        "header_scan_rows": gi("header_scan_rows",
                               core.DEFAULT_SHEET_CFG["header_scan_rows"]),
    }
    poly = [core.parse_group_spec(s) for s in _split_multi(g("poly", "poly_groups"))]
    dedupe = [core.parse_dedupe_spec(s) for s in _split_multi(g("dedupe"))]
    on_unmatched = (g("on_unmatched", default="keep") or "keep").lower()
    if on_unmatched not in ("keep", "drop"):
        on_unmatched = "keep"

    return {
        "name": g("name") or (os.path.basename(swagger) if swagger else f"job{idx + 1}"),
        "swagger": swagger, "mapping": mapping, "out": out, "report": report, "html": html,
        "sor_swaggers": sor_swaggers, "require_sor_field": gb("require_sor_field"),
        "on_unmatched": on_unmatched,
        "add_discriminator": gb("discriminator", default=True),
        "disc_prop": g("disc_prop", default="resultContext"),
        "poly": poly, "dedupe": dedupe,
        "auto_poly": gb("auto_poly", default=not poly),
        "auto_dedupe": gb("auto_dedupe", default=not dedupe),
        "dedupe_strictness": (g("dedupe_strictness", default="deep") or "deep").lower(),
        "protected": tuple(s.strip() for s in g("protected", default="ErrorResponse").split(",")
                           if s.strip()),
        "min_match_rate": gf("min_match_rate", 0.25),
        "max_total_loss": gf("max_total_loss", 0.75),
        "guards": gb("guards", default=True),
        "assertions": gb("assertions", default=True),
        "sheet_cfg": sheet_cfg,
    }


# --------------------------------------------------------------------------- #
# Worker (module-level so it is picklable for ProcessPoolExecutor)
# --------------------------------------------------------------------------- #
def _blank(name, out, status, error):
    return {"name": name, "status": status, "out": out, "html": None,
            "total": 0, "matched": 0, "match_rate": "", "eliminated": 0, "flagged": 0,
            "cascade_objects": 0, "cascade_refs": 0, "poly": 0, "dedup": 0,
            "assertions": 0, "error": error}


def run_job(job):
    name = job["name"]
    try:
        if not job["swagger"] or not os.path.isfile(job["swagger"]):
            raise FileNotFoundError(f"swagger not found: {job['swagger']}")
        if not job["mapping"] or not os.path.isfile(job["mapping"]):
            raise FileNotFoundError(f"mapping not found: {job['mapping']}")
        for y in job["sor_swaggers"]:
            if not y or not os.path.isfile(y):
                raise FileNotFoundError(f"SOR YAML not found: {y}")
        for p in (job["out"], job["report"], job["html"]):
            if p:
                os.makedirs(os.path.dirname(os.path.abspath(p)), exist_ok=True)
        res = core.run(
            job["swagger"], job["mapping"], job["out"], job["report"],
            html_report=job["html"], sor_swaggers=job["sor_swaggers"],
            require_sor_field=job["require_sor_field"], on_unmatched=job["on_unmatched"],
            add_discriminator=job["add_discriminator"], discriminator_prop=job["disc_prop"],
            poly_groups=job["poly"], dedupe_pairs=job["dedupe"],
            auto_poly=job["auto_poly"], auto_dedupe=job["auto_dedupe"],
            dedupe_strictness=job["dedupe_strictness"],
            protected=job["protected"], sheet_cfg=job["sheet_cfg"],
            min_match_rate=job["min_match_rate"], max_total_loss=job["max_total_loss"],
            enforce_guards=job["guards"], enforce_assertions=job["assertions"],
            log=lambda m: None,
        )
        d = res["diagnostics"]
        return {"name": name, "status": "ok", "out": job["out"], "html": job["html"],
                "total": d["attributes_total"], "matched": d["attributes_matched"],
                "match_rate": f"{d['match_rate'] * 100:.1f}%",
                "eliminated": len(res["dropped"]), "flagged": d["attributes_unmatched"],
                "cascade_objects": len(res["cascade"]["removed_schemas"]),
                "cascade_refs": len(res["cascade"]["removed_properties"]),
                "poly": len(res["poly_log"]), "dedup": len(res["dedupe_log"]),
                "assertions": len(res["assertions"]), "error": ""}
    except core.MappingCoverageError as e:
        return _blank(name, job["out"], "ABORTED", f"mapping coverage: {e}")
    except core.ExcessiveLossError as e:
        return _blank(name, job["out"], "ABORTED", f"excessive loss: {e}")
    except core.OutputAssertionError as e:
        r = _blank(name, job["out"], "ASSERT", str(e))
        r["html"] = job["html"]
        return r
    except Exception as e:  # noqa: BLE001 - report, never abort the batch
        return _blank(name, job["out"], "ERROR", f"{type(e).__name__}: {e}")


# --------------------------------------------------------------------------- #
# Consolidated reports
# --------------------------------------------------------------------------- #
COLS = ["name", "status", "total", "matched", "match_rate", "eliminated", "flagged",
        "cascade_objects", "cascade_refs", "poly", "dedup", "assertions", "out", "html", "error"]


def write_summary_csv(path, results):
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=COLS)
        w.writeheader()
        for r in results:
            w.writerow({c: r.get(c, "") for c in COLS})


def write_index_html(path, results):
    base = os.path.dirname(os.path.abspath(path))

    def rel(p):
        try:
            return os.path.relpath(p, base) if p else ""
        except ValueError:
            return p or ""

    ok = [r for r in results if r["status"] == "ok"]
    failed = [r for r in results if r["status"] != "ok"]
    tot_elim = sum(r["eliminated"] for r in ok)
    tot_flag = sum(r["flagged"] for r in ok)
    matched = sum(r["matched"] for r in ok)
    total = sum(r["total"] for r in ok) or 1

    trs = []
    for r in results:
        cls = {"ok": "ok", "ABORTED": "bad", "ERROR": "bad", "ASSERT": "warn"}.get(r["status"], "bad")
        out_link = (f'<a href="{_html.escape(rel(r["out"]))}">yaml</a>'
                    if r.get("out") and r["status"] in ("ok", "ASSERT") else "")
        html_link = f'<a href="{_html.escape(rel(r["html"]))}">diff</a>' if r.get("html") else ""
        trs.append(
            f'<tr><td>{_html.escape(str(r["name"]))}</td>'
            f'<td class="{cls}">{_html.escape(r["status"])}</td>'
            f'<td class="num">{r["matched"]}/{r["total"]}</td>'
            f'<td class="num">{_html.escape(str(r["match_rate"]))}</td>'
            f'<td class="num">{r["eliminated"]}</td><td class="num">{r["flagged"]}</td>'
            f'<td class="num">{r["cascade_objects"]}/{r["cascade_refs"]}</td>'
            f'<td class="num">{r["poly"]}</td><td class="num">{r["dedup"]}</td>'
            f'<td class="num">{r["assertions"]}</td>'
            f'<td>{out_link}</td><td>{html_link}</td>'
            f'<td class="err">{_html.escape(r["error"])}</td></tr>')

    core._write(path, f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<title>SOR Polymorphizer - batch index</title><style>
 body{{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f4f6f8;color:#1f2a33}}
 header{{background:#21295C;color:#fff;padding:20px 28px}} header h1{{margin:0;font-size:20px}}
 .wrap{{padding:22px 28px;max-width:1350px;margin:0 auto}}
 .cards{{display:flex;gap:14px;margin-bottom:20px;flex-wrap:wrap}}
 .card{{background:#fff;border-radius:10px;padding:16px 20px;box-shadow:0 1px 4px rgba(0,0,0,.1);min-width:130px}}
 .card .n{{font-size:30px;font-weight:700;color:#065A82}} .card .l{{font-size:13px;color:#5A6B78}}
 table{{border-collapse:collapse;width:100%;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.08)}}
 td,th{{padding:8px 12px;border-bottom:1px solid #eef2f5;text-align:left;font-size:13px}}
 th{{background:#eef3f7}} .num{{text-align:right;font-variant-numeric:tabular-nums}}
 .ok{{color:#1E7A34;font-weight:600}} .bad{{color:#B3261E;font-weight:700}}
 .warn{{color:#9A6A00;font-weight:700}} .err{{color:#B3261E;font-size:12px}}
 a{{color:#065A82}}
</style></head><body>
<header><h1>SOR Polymorphizer — batch results (engine v{core.__version__})</h1></header>
<div class="wrap">
 <div class="cards">
  <div class="card"><div class="n">{len(results)}</div><div class="l">swaggers processed</div></div>
  <div class="card"><div class="n">{len(ok)}</div><div class="l">succeeded</div></div>
  <div class="card"><div class="n">{len(failed)}</div><div class="l">aborted / failed</div></div>
  <div class="card"><div class="n">{matched / total * 100:.0f}%</div><div class="l">overall mapping<br>match rate</div></div>
  <div class="card"><div class="n">{tot_elim}</div><div class="l">attributes eliminated</div></div>
  <div class="card"><div class="n">{tot_flag}</div><div class="l">kept &amp; flagged</div></div>
 </div>
 <table><tr><th>Job</th><th>Status</th><th>Matched</th><th>Rate</th><th>Elim.</th>
  <th>Flagged</th><th>Cascade obj/ref</th><th>Poly</th><th>De-dup</th><th>Assert</th>
  <th>Output</th><th>Report</th><th>Error</th></tr>
 {''.join(trs)}
 </table>
</div></body></html>""")


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main(argv=None):
    ap = argparse.ArgumentParser(description="Batch-run the SOR refactor from a manifest.")
    ap.add_argument("--manifest", required=True, help="CSV or Excel manifest, one job per row")
    ap.add_argument("--base-dir", default=None,
                    help="resolve relative paths against this folder (default: manifest folder)")
    ap.add_argument("--jobs", type=int, default=1, help="parallel workers (default 1)")
    ap.add_argument("--summary-csv", default=None,
                    help="summary CSV path (default: batch_summary.csv next to manifest)")
    ap.add_argument("--index", default=None,
                    help="HTML index path (default: batch_index.html next to manifest)")
    args = ap.parse_args(argv)

    if not os.path.isfile(args.manifest):
        print(f"Manifest not found: {args.manifest}", file=sys.stderr)
        return 2
    base_dir = args.base_dir or os.path.dirname(os.path.abspath(args.manifest))
    mdir = os.path.dirname(os.path.abspath(args.manifest))
    summary_csv = args.summary_csv or os.path.join(mdir, "batch_summary.csv")
    index_html = args.index or os.path.join(mdir, "batch_index.html")

    rows = load_manifest(args.manifest)
    jobs = [build_job(r, base_dir, i) for i, r in enumerate(rows)]
    if not jobs:
        print("Manifest has no job rows.", file=sys.stderr)
        return 2

    print(f"Processing {len(jobs)} swagger(s) with {max(1, args.jobs)} worker(s)…\n")
    results = []
    if args.jobs and args.jobs > 1:
        with cf.ProcessPoolExecutor(max_workers=args.jobs) as ex:
            for res in ex.map(run_job, jobs):
                results.append(res)
                _print_line(res)
    else:
        for job in jobs:
            res = run_job(job)
            results.append(res)
            _print_line(res)

    write_summary_csv(summary_csv, results)
    write_index_html(index_html, results)

    ok = sum(1 for r in results if r["status"] == "ok")
    failed = len(results) - ok
    tot = sum(r["eliminated"] for r in results if r["status"] == "ok")
    print(f"\nDone. {ok} ok, {failed} aborted/failed, {tot} attributes eliminated in total.")
    print(f"Summary : {summary_csv}")
    print(f"Index   : {index_html}")
    return 1 if failed else 0


def _print_line(res):
    tag = {"ok": "OK ", "ABORTED": "ABT", "ASSERT": "ASR"}.get(res["status"], "ERR")
    if res["status"] == "ok":
        extra = (f"match={res['match_rate']} elim={res['eliminated']} "
                 f"flagged={res['flagged']} cascade={res['cascade_objects']}/"
                 f"{res['cascade_refs']} poly={res['poly']} dedup={res['dedup']}")
    else:
        extra = res["error"]
    print(f"  [{tag}] {res['name']:<40} {extra}")


if __name__ == "__main__":
    sys.exit(main())
