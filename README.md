# SOR Polymorphizer

Tools that inspect an OpenAPI/Swagger file and produce a new one in which
attributes **with no System-of-Record (SOR) connection are eliminated**, using
OpenAPI polymorphism (`allOf` sub-typing, optionally with a `discriminator`) so
that schemas shared across operations no longer carry unused attributes.

The SOR connection for each attribute is read from the data-model **mapping
spreadsheet**. Optionally, one or more of the System-of-Record's own **SOR API
YAML files** can be supplied to validate that the mapped fields actually exist
in the SOR.

An attribute is **kept** when the mapping gives it a concrete SOR field.
It is **eliminated** only when the mapping positively says there is none, that
is when the SOR cell is blank, `N/A`, `-`, `TBD` or `Not Used In SOR`.
An attribute the mapping does not mention at all is **kept and flagged**,
because a lookup that found nothing is not evidence that the SOR lacks the
field. Classes (`object` / `array`) are structural and are never removed for
lacking an SOR field, but an object that loses **all** of its attributes is
removed along with the references to it.

---

## What changed in v5, and why

Version 4 emptied the CASA v1.4.0 specification: 986 attributes were logged as
having no SOR connection, 945 of them with the reason `absent-from-mapping`.
Of 427 scalar attributes in the input, 39 survived, and 62 of 156 component
schemas were left holding a description and nothing else, so payload examples
rendered as `{}`. The cause was not the object generator. It was the mapping
lookup, plus four design choices that turned a lookup failure into silent data
loss. A seventh defect, an encoding fault reported after the first release,
is corrected in v5.1. All seven are listed below.

| # | Defect | Correction |
|---|---|---|
| **D1** | The spreadsheet layout was **assumed** (levels in columns C..G, data type I, SOR field J). Real mapping workbooks put the header row part-way down the sheet and carry the field name as one dotted path. On the CASA mapping, v4 read *Usage*..*Description* as the attribute name and an empty column as the SOR field, so nothing matched. | The header row is **located** and columns are resolved **by label**. Both the dotted-path layout and the legacy `Level 1..n` layout are supported, and the detected layout is printed, shown in the GUI and written into every report. |
| **D2** | Matching used a **flat, global leaf namespace**. The same leaf name in two operations collided, and qualified mapping rows could never be used. | Every schema's **usage paths** are computed by following `$ref` from the operation bodies, and attributes resolve **most specific first**: exact qualified path, then path suffixes, then bare leaf. The strategy used is recorded per attribute. |
| **D3** | **Not found was treated as delete.** A join failure was indistinguishable from a genuine "not used in SOR". | Elimination now requires a **positive** no-SOR signal. Unmatched attributes are kept and listed under *Kept but not found in the mapping*, with the lookup keys that were tried. `--on-unmatched drop` restores the old behaviour explicitly. |
| **D4** | Emptied objects were **left as shells**, with every `$ref` still pointing at them. | **Cascade pruning.** An object that loses all its attributes is removed, together with the properties that referenced it, the array properties whose item type died, and the matching `required` entries. Operation body schemas are never deleted; they are reported instead. |
| **D5** | De-duplication ran **after** elimination and compared **property-key sets**. Two schemas that differed in the input became "identical" once both were gutted and were merged. On CASA this collapsed `AccountLimitResponse.PaymentTransaction` into `PaymentTransaction` and lost `paymentTransactionStatus`. | Equivalence is computed on the **original** document with a deep (or shape) comparison, and **re-verified** at apply time. A pair that diverged is skipped and logged, never merged. |
| **D6** | **Nothing checked the output.** An emptied specification was written and reported as a success. | **Circuit breakers** (minimum mapping match rate, maximum total attribute loss) abort before anything is written, and **post-run assertions** check for empty object schemas, dangling `$ref`s and `required` entries with no property. |
| **D7** *(v5.1)* | **Text encoding was left to the platform.** The output was written with the platform default encoding and the HTML report then read it back as hard-coded UTF-8. On Windows that raised `UnicodeDecodeError: 'utf-8' codec can't decode byte 0x97` on the em dash in the CASA description, and a specification saved as ANSI could not be read at all. A report failure also discarded the whole run. | Input encoding is **detected** (BOM, UTF-8, then cp1252 and latin-1); the output and both reports are **always written as UTF-8**; and report generation is **non-fatal**, so a cosmetic failure never discards a document that is already on disk. |

### Measured effect on the CASA v1.4.0 specification

| | v4 | v5 |
|---|---|---|
| Mapping rows indexed | 0 usable | 2,371 across 26 sheets |
| Attributes resolved against the mapping | 0 of 986 | **986 of 986 (100%)**, 979 by exact qualified path |
| Attributes eliminated | 986 (945 of them merely "not found") | **491**, every one on a positive no-SOR signal |
| Scalar attributes surviving (input has 427) | 39 | **279** |
| Component schemas left with `properties: {}` | 62 | **0** |
| `paymentTransactionStatus` | silently lost to a bad merge | preserved |
| Post-run assertion failures | not checked | 0 |

---

## Files

| File | Purpose |
|------|---------|
| `polymorphize_core.py` | The engine. Fully parameterised, no hard-coded names. |
| `polymorphize_cli.py`  | Command-line front-end, including `--diagnose`. |
| `polymorphize_gui.py`  | Windows desktop app (Tkinter) with a mapping-coverage panel. |
| `polymorphize_batch.py`| Batch runner — process many swaggers from a manifest. |
| `build_windows_exe.bat`| One-click PyInstaller build → `dist\PolymorphizeSwagger.exe`. |
| `run_batch.bat`        | Windows launcher for the batch runner. |
| `manifest_template.csv` / `.xlsx` | Starter manifest to copy and fill in. |
| `tests/test_regressions.py` | Regression tests, one group per defect above. |

---

## Start with a coverage check

Before transforming anything, confirm that the mapping actually joined to the
specification. This writes nothing:

```bash
python polymorphize_cli.py --diagnose \
    --swagger casa.yaml --mapping casa_mapping.xlsx
```

It prints the detected layout for every sheet, how many mapping rows were
indexed, the **match rate**, the resolution strategies used, and for anything
it could not find, the exact keys it tried alongside the nearest keys in the
mapping. A low match rate means the lookup failed. It does not mean the SOR is
missing the fields.

---

## Windows application (recommended for non-technical users)

1. Install Python 3.9+ from python.org (tick "Add python.exe to PATH").
2. Double-click **`build_windows_exe.bat`**. It installs the dependencies and
   produces **`dist\PolymorphizeSwagger.exe`**, a single self-contained file.
3. Share or run that `.exe`; end users need **no** Python install.

The three actions are numbered on screen — **STEP 1 choose files → STEP 2
Analyze → STEP 3 Run (approve)** — and the Run button stays disabled until you
have analyzed:

1. **STEP 1** — choose the input Swagger and **one** SOR mapping spreadsheet.
   Optionally attach **one or more SOR API YAML files** with **Add…** or by
   dropping them onto the list; these validate that every mapped attribute's
   SOR field really exists in the SOR.
2. **STEP 2 — Analyze.** The **Mapping coverage** box fills in first: match
   rate, rows indexed, resolution strategies and the detected layout for each
   sheet. Read it before anything else. The log then lists what would be
   eliminated, what would be kept and flagged, what would be cascade-removed,
   the proposed de-duplications and the polymorphism groups. Groups and
   de-duplications are detected from the files you entered and can be edited.
3. **STEP 3 — Run transform**, then **approve**. Nothing is written until you
   approve, and the approval dialog repeats the match rate with a warning when
   it is low. Changing any input invalidates the analysis.

### Seeing what changed
Every change is visible four ways:

* **The coverage box and the log** — colour-coded, with counts.
* **A one-click HTML change report** — summary cards including the match rate,
  the detected layout per sheet, the eliminated attributes with the strategy
  that matched them, the attributes kept but not found, the cascade removals,
  the assertion results and a full red/green diff.
* **A Markdown report** — the same content in text form.
* **Inside the output swagger itself** — a banner comment, an
  `info.x-sor-polymorphizer` record carrying `mappingCoverage`,
  `eliminatedAttributes`, `keptButNotInMapping`, `cascadeRemovedObjects` and
  `cascadeRemovedReferences`, and `>>> CHANGE` comments at each edited schema.

---

## Command line

```bash
pip install ruamel.yaml openpyxl

python polymorphize_cli.py \
    --swagger casa.yaml --mapping casa_mapping.xlsx \
    --out casa_polymorphic.yaml --report report.md \
    --html-report changes.html --auto-dedupe
```

### Parameters

| Parameter | Meaning |
|-----------|---------|
| `--swagger` / `--out` / `--report` / `--html-report` | input, output and report paths |
| `--mapping` | SOR mapping spreadsheet, exactly one |
| `--diagnose` | report mapping coverage and exit without writing |
| `--sor-swagger` | SOR API YAML to validate mapped fields, repeatable or `;`-separated |
| `--require-sor-field` | also eliminate a mapped attribute if its SOR field is absent from the SOR YAML(s) |
| `--on-unmatched keep\|drop` | what to do with an attribute the mapping does not mention (default `keep`) |
| `--poly "Base = A:v1, B:v2"` | factor schemas A,B into a shared base (repeatable) |
| `--dedupe "Owner.prop = Target"` | replace an inline object with a `$ref` (repeatable) |
| `--auto-poly` / `--auto-dedupe` | auto-detect families / inline duplicates |
| `--dedupe-strictness deep\|shape` | `deep` requires full equality, `shape` ignores prose and examples |
| `--discriminator` / `--disc-prop` | add a discriminator and name it |
| `--protected` | schema names never pruned (default `ErrorResponse`) |
| `--min-match-rate` | abort below this mapping match rate (default `0.25`) |
| `--max-total-loss` | abort above this share of attributes eliminated (default `0.75`) |
| `--no-guards` / `--no-assert` | downgrade the guards or the assertions to warnings |
| `--no-auto-layout` | use the fixed columns below instead of detecting the layout |
| `--level-first-col` / `--level-last-col` / `--dtype-col` / `--sor-col` / `--not-used-text` | fixed spreadsheet layout, used only with `--no-auto-layout` |
| `--header-scan-rows` | how many rows to scan for the header (default `40`) |

Exit codes: `0` success, `3` mapping coverage abort, `4` excessive loss abort,
`5` post-run assertion failure.

---

## Batch — many swaggers at once

Drive the run from a **manifest**, one job per row. Start from
`manifest_template.csv` or `.xlsx`. Each row lists at least a `swagger` and its
`mapping`; everything else has a sensible default. Relative paths resolve
against the manifest's own folder.

```bash
python polymorphize_batch.py --manifest jobs.csv --jobs 4
```

On Windows, drag your manifest onto `run_batch.bat`. It writes each transformed
swagger with its HTML change report, plus:

* `batch_summary.csv` — one row per job with the **match rate**, attributes
  matched, eliminated, flagged, cascade counts and any error,
* `batch_index.html` — a linked roll-up with the overall match rate and
  per-job status.

A job whose mapping barely matched, or which would strip most of the
specification, is reported as **ABORTED** and writes nothing. That is
deliberate: a silently emptied specification is worse than a failed job.
Override per row with `guards`, `min_match_rate` and `max_total_loss`.

As before, groups and de-duplications are detected **per file** at run time, so
leave the `poly` and `dedupe` columns blank unless you want to override the
automatic proposal for that one file.

---

## Tests

```bash
python tests/test_regressions.py
```

Forty-nine checks grouped by defect: layout detection on both spreadsheet
shapes, the not-used token set, qualified-path resolution, keep-and-flag,
cascade pruning of an emptied object and of the array that referenced it,
de-duplication that must not merge and de-duplication that must, both circuit
breakers, the post-run assertions, and the encoding cases (UTF-8, BOM and
ANSI input, UTF-8 output, and a failing report that must not lose the run).

---

## Notes

* **Mapping coverage is the first thing to check** when output looks thin. The
  report's *Detected spreadsheet layout* section names the header row and the
  column chosen for the field name and the SOR field. If those are wrong, fix
  them there rather than in the specification.
* **Discriminator trade-off:** with `--discriminator` a required selector
  property is added to each generated base. Omit it for pure `allOf`
  inheritance; the elimination of non-SOR attributes is identical either way.
* **`required` semantics:** eliminating an attribute removes it from its
  schema's `required` list. A post-run assertion confirms no `required` entry
  survives without its property.
* **Encoding:** the CASA specification contains one em dash (U+2014), three
  right single quotation marks and 56 non-breaking hyphens (U+2011). The
  non-breaking hyphen has **no cp1252 equivalent**, so re-saving the YAML as
  ANSI in a Windows editor will corrupt it. Keep these files as UTF-8. The
  tool reads ANSI if it has to and always writes UTF-8, and it logs a note
  when the input was not UTF-8.
