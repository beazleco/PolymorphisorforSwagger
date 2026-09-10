# SOR Polymorphizer 6.6

Turns a System of Record mapping workbook into **one** OpenAPI 3.0.3
specification, publishing only the attributes the SOR actually supports and
specialising shared schemas so a consumer is never shown a field that is not
wired to anything.

Two workbook formats are accepted and the tool decides which one it has been
given by looking at the file, so nothing needs to be declared on the command
line. Every operation it publishes carries the full Apigee error set and
declares its lifecycle stage, and every schema states its class and its
context separately, so a downstream process can parse them apart.

It also writes the field mapping document back out, from the same run that
produced the specification, with three columns saying what reached the
interface and what did not. That document is itself a valid input, so an
analyst can correct the rows it flags and feed it straight back in.

> Full documentation for this release is in `SOR_Polymorphizer_User_Manual.docx`
> (for analysts building the workbook, including an error guide organised by
> root cause) and `SOR_Polymorphizer_Technical_Guide.docx` (for maintainers,
> including a prompt guide for working with an AI coding assistant).
> `DOCUMENTATION_NOTES.md` records what changed at each release, the rulings
> taken on every open question, and what remains open. Both Word documents
> describe 6.3 behaviour and predate the second input format, so read the notes
> for anything added since.

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

# What specialisation prevents, on any workbook
python tools/measure_specialisation.py my_workbook.xlsx

# Tests
python tests/run_all.py
```

Nothing above names the input format. Both formats go through the same
commands and the tool works out which it has been handed.

## Which format has been supplied

The tool recognises two workbook formats and classifies the file before it
reads it. Classification says what the workbook is; it never says whether the
workbook is any good, which is what `check` is for. A file it cannot place is
refused with exit code `3` rather than treated as an empty run.

| Format | How it is recognised | What it gives |
|---|---|---|
| **Level format** | `Level 1` in the header row | Several SOR columns per sheet, so both axes of variance |
| **Field mapping document** | a `Parameter Type` column beneath the banner block | One SOR column per sheet, so the operation axis only |

Recognition is a bounded peek: fourteen rows and thirty columns of at most six
worksheets, opened read only. On the reference field mapping workbook that
peek takes 284 ms where an ordinary open of the same file takes 75.7 seconds,
a factor of 266, which is why the desktop front end can classify a file the
moment its name is entered and label it before anything else happens.

A workbook carrying sheets of both kinds resolves to the Level format and the
mixture is reported.

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

## The field mapping contract

The second format states nesting as a dotted path in one column rather than as
a ladder of Level columns.

- **The banner occupies rows 1 to 9** and the header row is found beneath it,
  normally row 10. The tool searches the first sixteen rows for it rather than
  assuming a fixed position.
- **`Parameter Type` is authoritative.** It decides whether a row belongs to
  the request or the response, and only body rows are read. Header and
  parameter rows are skipped, exactly as in the Level format. Where the column
  and the section heading disagree, the column wins and the disagreement is
  reported as `F001`.
- **`Reusable API Field Name`** carries the dotted path, so
  `cardDetails.expiry.month` nests three deep. The container rows may be
  present or absent; either way the tree is the same.
- **One `SOR API Field Name` column**, so there is a single SOR per sheet and
  no variant axis. Every sheet is therefore pattern P1.
- **`N/A` and a blank cell both mean unmapped.** A blank is additionally
  reported as `F004` at the strict level, because a deliberate `N/A` and an
  unfinished cell look identical otherwise.
- **`Required in Swagger` is advisory.** Obligation comes from `Usage`, which
  is the column analysts actually maintain.

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

## Class and context

Where a group splits because the operations using it publish different
attributes, the family shares one class and its members differ only in
context. The two are separated by a double underscore, so a following process
can parse them apart without a heuristic:

| Schema | Class | Context |
|---|---|---|
| `CardDetails` | `CardDetails` | the plain name, held by the shape most references point at |
| `CardDetails__Base` | `CardDetails` | the shared core |
| `CardDetails__CardDetailsRetrieve` | `CardDetails` | the shape that one operation needs |

The plain name goes to the shape most references point at, because a code
generator names its classes after component schemas and the plain name is what
a consumer meets. The base class is always present: a family of *n* shapes
publishes the core as `<Class>__Base` whether or not any operation uses it
alone.

Every schema also states the pair explicitly, so nothing has to be recovered
from the key at all:

```yaml
CardDetails__CardDetailsRetrieve:
  x-class: CardDetails
  x-context: CardDetailsRetrieve
  x-qualified-name: CardDetails~CardDetailsRetrieve
```

`__` is the separator in the key and `~` appears only inside
`x-qualified-name`, for a consumer that prefers a single readable string. The
reason is narrow and worth recording: the OpenAPI component key regex is
`^[a-zA-Z0-9\.\-_]+$`, which forbids `~`, and RFC 6901 makes `~` the escape
character in a JSON Pointer, so a `~` in a key has to be written `~0` in every
`$ref` that reaches it. Some validators accept the unescaped form, which is
worse than a rejection, because the defect ships silently.

None of this reaches the wire. A property key comes from the workbook, and the
class and context govern only the names of the components a generated client
declares.

### What specialisation prevents

| Reference workbook, 11 operations | Attribute occurrences a consumer is shown |
|---|---|
| With cross-operation specialisation | 228 |
| Without it, one shared shape per group | 275 |

Specialising prevents 47 occurrences, 21% more, that the SOR behind the
operation does not supply. On the field mapping sample, 22 operations, the
figures are 354 against 372.

Both numbers come from `tools/measure_specialisation.py`, which counts from the
workbook tree so that the pair differs in one respect only, and which runs on
any workbook:

```bash
python tools/measure_specialisation.py samples/creditcard_v2.0.0.xlsx
```

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

## The field mapping document

Alongside the specification the desktop window and the batch runner write a
field mapping document in the API COE's own format, produced by the same run,
so the two agree by construction rather than by somebody keeping them in step.

**It is not on the command line.** By instruction: the export belongs where a
whole set of documents is refreshed, which is the batch runner, and where an
analyst is working a single workbook, which is the window. `generate --help`
says so rather than leaving the absence to be discovered.

```bash
# The batch runner writes one for every workbook. Set export to false in a
# manifest row to suppress it for that workbook.
python polymorphize_batch.py --manifest jobs.csv --out-dir ./batch
```

### What is in it

The format's own nine columns, unchanged, then three more:

| Column | What it holds |
|---|---|
| **Publication Status** | Published, Variant only, Not in the SOR, No SOR field, Container, or Not generated |
| **Why** | The reason, naming the SOR field or the variants where that is what decided it |
| **Analyst Action** | What to change to publish the element, blank where nothing needs doing |

A row that did not reach the interface is shaded, so the rows worth attention
are the ones the eye lands on. A `_Round Trip` sheet at the front records the
round number, the source workbook, the specification the document matches and
that specification's fingerprint, a per sheet count of what was published, and
anything the conversion could not carry.

### One sheet per SOR endpoint

The format has a single SOR API Field Name column, so a Level sheet naming
several SOR endpoints becomes several sheets, one per endpoint, each complete.
The reference workbook's twelve operations become seventeen sheets on that
rule. A field mapping sheet already names one endpoint, so its sheet name is
preserved.

### Iteration

The exported workbook is a valid input, so the loop closes:

1. Generate. The document names the elements that did not reach the interface.
2. Correct those rows in the document itself.
3. Feed it back in. The next specification carries the corrections, and the
   next document is written from that run.

For that to be safe the export has to carry the whole of the original, not
only the parts the tool reads. Where the source was a field mapping document
every cell is written back **verbatim** from the row it came from, including
the header rows, the parameter rows, and any column beyond the format's nine.
An operation that failed to generate is carried too, marked Not generated with
the correction against it, because losing an operation on the round trip would
be worse than any finding it carries. The export is a fixed point after the
first round: exporting an export changes nothing.

Where the source was a Level workbook there is nothing to carry and the sheet
is composed from the tree, which is a conversion rather than a round trip. The
`_Round Trip` sheet states the three losses: Remarks and Required in Swagger
have no source in the Level format, the header block comes from the constant
rather than the workbook, and the Schema and Usage wording is the tool's
spelling.

### The standard header block

The Level format carries no header rows. The reference field mapping document
carries 220 of them across its 22 sheets, and they are seven distinct headers
repeated. So the block is a constant, taken from that document, and supplied
where the source had none.

Writing it from a constant also settles an inconsistency in the source, where
`x-BDO-Application-Id` is declared both `string` and `string(10)`, three
headers appear once with a length and a Usage and once with neither, and three
rows carry `CVV`, `auxiliaryPan` and `auxiliaryExpiry` in the Required in
Swagger column, which is paste drift from the neighbouring column.

## The API lifecycle stage

Every operation declares the stage it has reached:

```yaml
description: '**<u>API Lifecycle Status - Design</u>**'
x-api-lifecycle-status: Design
x-use-case: Use this API to retrieve card available funds details
```

The description is the statement and nothing else, so it renders as one line
in Swagger UI and Redoc. The use case is not displaced: it remains the
operation summary and `x-use-case` carries it untruncated. `Design` is fixed
rather than an option, because the tool generates a design and an operation
that has reached a later stage is no longer something it produced.

## The error set

Every operation carries the same twelve error responses, referenced from
`components/responses` rather than inlined:

`400` `401` `403` `404` `405` `409` `422` `429` `500` `502` `503` `504`

The set is not a judgement call. It is taken from the API COE's own
specification, `apicoeissueddeviceadministrationswaggerv1.0.8.yaml`, where all
twelve are structurally identical, share the `ErrorResponse` schema, and each
carries the disclaimer stating that the text is an Apigee message and is not
mapped to the System API error. That disclaimer is reproduced verbatim.

Two departures from that sample are deliberate. The three
`x-BDO-Client-Request-Id`, `-Trace-Id` and `-Span-Id` headers appear on its
success responses as well as its errors, so they are a house convention rather
than an error convention and this tool attaches them to every response. And the
sample's `ErrorResponse` example does not conform to its own schema: the schema
declares `errors` as an array while the example supplies `error` as a single
object, `status` and `realm` are absent, and the title reads "The request was
successful, but there is no content in the response", which is a 204 message in
an error example. It is not copied. The example this tool writes conforms, and
the defect has been raised as a note for the API COE rather than propagated
into every specification generated from here.

## Empty messages

A `Request Body` or `Response Body` label must be present, but the section may
be empty. An empty request emits no `requestBody` and keeps its natural verb;
an empty response emits `204 No Content`. A *missing* label is a hard failure.

## Failure handling

A sheet that cannot be generated fails alone; the rest of the workbook
generates. For a failed sheet the tool writes no specification, removes any
left from an earlier run, and writes `<sheet>.FAILED.md` naming the cell and
the correction.

A `--split` run writes no field mapping document, because the document states
which specification it matches and a split run produces several.

Exit codes: `0` all generated, `2` workbook unreadable, `3` format not
recognised, `6` some failed, `7` none generated. Codes 3 and 6 are not zero on
purpose: a file the tool cannot place used to produce a clean run with no
output, which reads as success.

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
| `polymorphize_classify.py` | Decides which format a file is, on a bounded read-only peek |
| `polymorphize_workbook.py` | Reads the Level format. Layout, banner, sections, tree, findings |
| `polymorphize_mapping.py` | Reads the field mapping format into the same tree |
| `polymorphize_errors.py` | The Apigee error set, the shared schema and the standard headers |
| `polymorphize_export.py` | The field mapping document, written back out with the difference columns |
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
| `samples/creditcard_validation.md` | The lenient check report for that workbook |
| `samples/creditcard_validation_strict.md` | The strict one, which adds the template contract |
| `samples/issueddevice_fieldmapping_v1.0.7.xlsx` | The reference **field mapping** workbook, 22 endpoints |
| `samples/issueddevice_out/` | Its specification, report, showcase and **field mapping document**. All 22 generate |
| `samples/creditcard_out/openapi_field_mapping.xlsx` | The same document converted from the Level format, 17 sheets for 11 operations |
| `samples/issueddevice_validation.md` | The check report for it |
| `samples/issueddevice_out/issued-device_field_mapping.xlsx` | The round trip case: the document written back in its own format |
| `samples/issueddevice_reference_swagger_v1.0.8.yaml` | The API COE specification the error set was taken from, and the source of the standard header block |
| `samples/manifest_template.csv` | A batch manifest, including the `sor` column |
| `samples/archive_v5.3/` | Output from 5.3, kept for reference only. Not a baseline: see its own note |
| `tools/build_sor_fixture.py` | Builds those fixtures from a workbook's SOR columns |
| `tools/measure_specialisation.py` | Reproduces the specialisation figures above on any workbook |
