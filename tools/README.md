# SOR Polymorphizer 6.3

Turns a System of Record mapping workbook into **one** OpenAPI 3.0.3
specification, publishing only the attributes the SOR actually supports and
specialising shared schemas so a consumer is never shown a field that is not
wired to anything.

> Full documentation has not yet been updated for this release. See
> `DOCUMENTATION_NOTES.md` for what changed and what the documentation pass
> must cover.

## Quick start

```bash
# Write a template to work from
python polymorphize_cli.py template SOR_mapping_template.xlsx

# Check a workbook without writing anything
python polymorphize_cli.py check my_workbook.xlsx
python polymorphize_cli.py check my_workbook.xlsx --level strict --report check.md
python polymorphize_cli.py check my_workbook.xlsx --sor sor_api.yaml

# Generate one specification, verified against the System of Record
python polymorphize_cli.py generate my_workbook.xlsx ./openapi \
    --sor sor_api.yaml --sor more_sor_apis/

# Without --sor every SOR field name is taken on trust
python polymorphize_cli.py generate my_workbook.xlsx ./openapi

# One file per sheet instead, for debugging a single endpoint
python polymorphize_cli.py generate my_workbook.xlsx ./openapi --split

# Several workbooks
python polymorphize_batch.py --manifest jobs.csv --out-dir ./batch

# The desktop front end
python polymorphize_gui.py     # drag and drop needs: pip install tkinterdnd2

# Tests
python tests/run_all.py
```

## The workbook contract

One worksheet describes one endpoint. Every worksheet of the workbook is
merged into a single specification.

- **Row 1** is the header row. It carries `Level 1` to `Level N`, then
  `Dummy Value`, `Data Type`, `Usage` and `Descriptions`, then one column per
  SOR endpoint headed `SOR <endpoint>`.
- **Rows 2 to 6** are the banner: `Use Case:`, `Service Domain:` with
  `BQ: <name>` beside it, `Equivalent BIAN API Endpoint:`,
  `Proposed Business API Endpoint:` and `SOR API Endpoint:`.
- **Nesting** is the Level columns. A name in `Level 3` belongs to the nearest
  row above it in `Level 2`. Never skip a Level.
- **`Method:`** is optional. Without it the HTTP method comes from the BIAN
  action term, so a Retrieve becomes a GET, which cannot carry a body. State
  POST there for an operation that reads with a request body.
- **Only `Request Body` and `Response Body` are read.** Header and parameter
  sections are skipped. Anything below three or more blank rows is ignored, so
  pasted payload samples and working notes are safe to keep there.
- **An empty SOR cell means that endpoint does not supply the attribute**, and
  the attribute is removed from that variant. That is how the interface
  shrinks.

A sheet with no `Level` columns is a support sheet and is skipped silently.

## Polymorphism: two axes

Polymorphism only means anything inside a single namespace, which is why the
default output is one specification. Merging the sheets makes two kinds of
variance visible, and they are resolved differently because they are settled
at different times.

| Axis | Question | Settled | Mechanism |
|---|---|---|---|
| Which operation | Which endpoint am I calling? | Design time: the path decides | `allOf` specialisation, **no discriminator** |
| Which variant | Which SOR endpoint serves this call? | Run time: the client sets `requestVariant` | `oneOf` with `discriminator: requestVariant` |

The operation axis takes no discriminator because there is no runtime
ambiguity to resolve, and one would demand a tag property that no SOR supplies.
The two compose: a shared group can sit inside a variant of a variant-driven
operation.

The variant axis follows the shape of the sheet:

| Pattern | Condition | Output |
|---|---|---|
| P1 | One SOR column | One schema per message, pruned |
| P2 | Several SOR columns and a `requestVariant` enumeration | Base plus `allOf` variants, wrapped in `oneOf` with a discriminator on the request |
| P3 | Several SOR columns, no tag | `oneOf` without a discriminator |

In P2 the request carries `discriminator: requestVariant` and each variant
narrows its `enum` to one code. The response is a `oneOf` without a
discriminator, annotated `x-selected-by: requestVariant`, because OpenAPI
cannot express a response subtype that depends on a request property.

Every group with children is published as a named component, so a property is
always either a scalar or a `$ref`. Shapes are compared on the contract, type,
format, length, obligation and enum, and never on prose, so differing
descriptions cannot fragment the namespace. An SOR field that differs between
operations is recorded as `x-sor-field-by-operation`.

### On the reference workbook

| | Attributes a consumer of one operation is shown |
|---|---|
| Merged, with cross-operation specialisation | 137 |
| Merged, without it | 188 |

Specialising prevents 51 attribute occurrences, 37% more, that the SOR behind
the operation does not supply.

## How an element is known to be backed by the SOR

The workbook naming an SOR field is an assertion. Supply the System of Record
specification with `--sor` and it becomes evidence.

For every row the tool takes the SOR endpoint from the sheet's
`SOR API Endpoint` banner cell, which is the definitive statement of it, and
looks the field up: a Request Body row in that endpoint's request, meaning its
body for POST and PUT or its parameters for GET, and a Response Body row in
its response. One banner endpoint serves every SOR column; several are paired
with the columns in order.

| Rule | |
|---|---|
| A field that is not there | The element is **excluded** from the interface, and the removal is reported against its cell |
| Matching | On the **last segment** of the field name, so `accountInfo.currency` resolves to any element named `currency`. A name found in more than one place is kept and reported as ambiguous |
| A cell naming several fields | Resolves if any of them does, which suits both the comma form for alternatives and the space form for several fields feeding one element |
| Which endpoint | The `SOR API Endpoint` **banner cell only**. An endpoint on a column header is a label: where the two differ the banner wins and the difference is reported |
| Nothing on a sheet resolving | The operation **fails**: that means the wrong endpoint or the wrong SOR file, not an empty interface |

Without `--sor` none of this happens, and every report says so prominently.

## The showcase report

Alongside the specification the tool writes an HTML report named after the
service domain, so the reference workbook produces `credit-card.html`. It is
organised by the **integration API** a consumer calls, not by the SOR
endpoints behind it.

It has two halves. For every class of every integration API it lists every
element the workbook declared and, for each one, either that it is published
or why it is not:

**Published** &middot; **Variant only** &middot; **Not in the SOR** &middot;
**No SOR field** &middot; **Container**

And a section of its own for **the elements eliminated from the interface**,
one table per integration API, because those are the point of the tool: every
one is a field a consumer would otherwise have seen, and could have populated,
with nothing behind it in the System of Record. Elements *narrowed* to some
variants rather than removed outright are listed separately.

Nothing in it is read back out of the generated specification. Every row is
the workbook's own tree annotated with the decision the tool made, which is
what makes it usable as evidence rather than as documentation. A sixth
disposition, **LOST**, appears only if an element was mapped, verified and yet
did not reach the interface; that should always be zero.

## Empty messages

A `Request Body` or `Response Body` label must be present, but the section may
be empty. An empty request emits no `requestBody` and keeps its natural verb;
an empty response emits `204 No Content`. A *missing* label is a hard failure.

## Failure handling

A sheet that cannot be generated fails alone; the rest of the workbook
generates. For a failed sheet the tool writes no specification, removes any
left from an earlier run, and writes `<sheet>.FAILED.md` naming the cell and
the correction.

Exit codes: `0` all generated, `2` workbook unreadable, `6` some failed,
`7` none generated. Code 6 is not zero on purpose.

## Checking levels

- **Lenient** reports only what stops an endpoint generating. This is what a
  run enforces.
- **Strict** adds the template contract: a use case, a service domain, a
  behaviour qualifier, a description on every attribute and an example on
  every mandatory one.

Every finding names the sheet, the cell, what was found, what was expected and
how to fix it.

## Modules

| File | Purpose |
|---|---|
| `polymorphize_workbook.py` | Reads the Level format. Layout, banner, sections, tree, findings |
| `polymorphize_generate.py` | One document per sheet, the three variant patterns, and the run |
| `polymorphize_merge.py` | The single specification: hoisting, cross-operation specialisation, collision checks |
| `polymorphize_sor.py` | Verification against the System of Record: the index, the resolver, the endpoint decision |
| `polymorphize_showcase.py` | The HTML showcase: every element and why it is or is not published |
| `polymorphize_validate.py` | Validation over the real reader and generator |
| `polymorphize_template.py` | Writes the template workbook |
| `polymorphize_gui.py` | Desktop front end |
| `polymorphize_cli.py` | `check`, `generate` (with `--split` and `--sor`), `template` |
| `polymorphize_batch.py` | Several workbooks, with a roll-up |
| `polymorphize_core.py` | The original trim engine: an existing specification plus a **dotted-path** workbook. Secondary path |

## Samples

| File | What it is |
|---|---|
| `samples/SOR_mapping_template.xlsx` | The template. Validates clean, generates cleanly |
| `samples/creditcard_v2.0.0.xlsx` | The reference workbook, 12 endpoints, one with a real defect |
| `samples/creditcard_out/` | The merged specification, the report and the failure explanation |
| `samples/creditcard_split/` | The same workbook with `--split`, one file per sheet |
| `samples/sor_creditcard_fixture.yaml` | A **synthetic** SOR specification for the reference workbook. Not SOR documentation: see the note in its own `info.description` |
| `samples/sor_template_fixture.yaml` | The same for the template's worked examples |
| `samples/manifest_template.csv` | A batch manifest, including the `sor` column |
| `tools/build_sor_fixture.py` | Builds those fixtures from a workbook's SOR columns |
