# Notes for the deferred documentation pass

> **Status: the documentation pass was carried out at release 6.3 and has not
> been repeated since.** `SOR_Polymorphizer_User_Manual.docx` and
> `SOR_Polymorphizer_Technical_Guide.docx` were rewritten from nothing against
> items 1 to 37 below. See "What the pass covered, and what it did not" in the
> middle of this file, and "Still open after 6.5" at the end. Items 38 to 47
> came in after that pass, so the two documents do not yet describe the second
> input format, classification, the error set, class and context, the field
> mapping export or the lifecycle stage. The items are kept because they are
> the record of why each decision was made.

This file records everything
the pass must cover. Items 1 to 7 were carried over from v5.3; items 8 to 18
came in with the Level-format contract in v6.0; items 19 to 25 came with the
single merged specification in v6.1; items 26 to 32 came with verification
against the System of Record and the HTML showcase in v6.2; items 33 to 37 came
with v6.3; item 38 with v6.4; items 39 to 47 with the two-format build in 6.5;
items 48 to 55 with the field mapping export in 6.6.

The two Word documents (`SOR_Polymorphizer_Technical_Guide.docx`,
`SOR_Polymorphizer_User_Manual.docx`) and the presentation
(`Polymorphism_OpenAPI_for_BAs_v2_corrected.pptx`) all describe a tool that no
longer works the way they say. Treat them as needing a rewrite rather than an
edit.

---

## Carried over from v5.3

### 1. The purpose statement changed
The tool is not a general OpenAPI trimmer. Its purpose is to stop a shared
class being published in full at every point of use, which exposes attributes
no System of Record supports and misleads a consumer about what a request
filter actually does. Both documents open with the old framing.

### 2. The discriminator recommendation was reversed, then partly restored
v5.2 and the presentation recommended a discriminator everywhere. v5.3
removed it in favour of plain `allOf`, on the evidence that in the CASA
specification every divergence was settled at design time by which operation
was called. v6.0 restores it **only** for the case the credit card workbook
introduced: several SOR endpoints behind one operation, selected at run time
by a `requestVariant` value. See item 10.

Slides 12 and 13 of the presentation still teach the v5.2 position.

### 3. Test counts
39 → 49 → 62 → 93 → 271 → 343 → 486 → 499 → 501 → 736 → **787** across
three suites (556 + 156 + 75). `tests/run_all.py` runs all three.

### 4. The workbook read is bounded
Sheets declare enormous dimensions because of stray formatting far below the
data. `core.load_grids` stops after a run of blank rows. This took a
validation pass from about 40 seconds to about half a second.

### 5. Encoding
Every read sniffs BOM, UTF-8, cp1252 and latin-1; every write is explicit
UTF-8. Report generation failures are non-fatal and never discard a good run.

### 6. Screenshots
Every screenshot in the user manual is obsolete. The GUI has been rebuilt
twice since they were taken, and v6.1 added the output-mode radio and the
drag-and-drop hints. Current captures are in `shots6/`.

### 7. Slide 11 has never been verified
The cheque SOR field names on slide 11 (`instrumentId`, `view3`,
`rejectReasonCode`) have not been checked against a cheque mapping workbook,
because none has been supplied.

---

## New in v6.0

### 8. The input contract is entirely different
The dotted-path workbook format the documentation describes is not the format
in use. The authoritative format, taken from
`samples/creditcard_v2.0.0.xlsx`:

| | Documented (v5.3) | Actual (v6.0) |
|---|---|---|
| Header row | Row 11, banner above it | Row 1, banner in rows 2 to 6 **below** it |
| Hierarchy | Dotted path in one Field Name column | `Level 1` to `Level N` columns; depth varies 3 to 6 by sheet |
| SOR mapping | One `SOR Field` column | One to five columns headed `SOR <endpoint>` |
| Example column | `Example` | `Dummy Value` |
| Usage column | `Usage` | `Usage` or `Required` |
| Description column | `Description` | `Description` or `Descriptions` |
| Not supported | Explicit vocabulary | An empty SOR cell |
| Scope | One workbook, one specification | One tab, one operation. See item 19: v6.1 merges the tabs into one specification |

Every worked example, screenshot and field reference in both documents needs
replacing.

### 9. Only Request Body and Response Body are read
This is a rule, not an implementation detail, and belongs in the user manual
prominently. Request Header, Response Header and any parameter section are
skipped. So is everything below the mapping grid: three or more blank rows
end the grid, and analysts may keep pasted SOR payload samples, notes or
working below that line. `SOR Payload` in the reference workbook is an entire
sheet of such content and is skipped because it has no Level columns.

### 10. Three polymorphism patterns, chosen by the sheet
Document the decision rule, not just the output.

| Pattern | Condition | Output |
|---|---|---|
| **P1** | One SOR column | One schema per message, pruned to what SOR supplies. No `allOf`, no `oneOf`. |
| **P2** | Several SOR columns and a `requestVariant` enumeration | Base holding the intersection; each variant `allOf` base plus its delta; request wrapped in `oneOf` with `discriminator: requestVariant`. |
| **P3** | Several SOR columns, no tag | `oneOf` without a discriminator over `allOf` derived schemas. |

Points that need stating explicitly:

- The discriminator property is retained and made required whatever the SOR
  columns say, because the API layer consumes it rather than SOR. Each variant
  narrows its `enum` to the single code that selects it.
- The **response** in P2 gets `oneOf` **without** a discriminator, plus
  `x-selected-by: requestVariant`. OpenAPI cannot express a response
  discriminator that depends on a request property, and inventing a response
  tag would be a lie about the payload.
- P3 arises when the SOR endpoint is chosen by which identifier the caller
  supplies. `Card Issued Device_Retrieve` is the only example, and it fails
  for an unrelated reason, so P3 has no working sample yet.
- These three are the **variant** axis only. v6.1 added the operation axis on
  top of them; see item 20.

### 11. Per-endpoint failure isolation
A sheet that cannot be generated fails alone. The rest of the workbook
generates. For a failed sheet the tool writes no specification, deletes any
specification left from an earlier run, and writes `<sheet>.FAILED.md`
instead. Document why: a stale specification that looks current is worse than
a missing one.

The failure is surfaced in four places, and the manual should show each: the
red banner in the GUI, the red row in the sheet table, the Findings tab with
Found and Fix, and the failure block at the end of the CLI output.

### 12. Two real defects in the reference workbook
Both make good worked examples for the manual's troubleshooting section.

- `Card Issued Device_Retrieve` (**S001**): the request schema sits under a
  section labelled `Request Header` and there is no `Request Body`. The
  analyst mislabelled it. The message names cell A7 and the replacement text.
- `Card RSAEncrypted` (**E001**): the Response Body declares three attributes
  and none names an SOR field, so the response would have been published as
  `properties: {}`. This is precisely the defect the developer originally
  reported on `CASAYAML.yaml`. **Superseded in v6.1**: an empty message is now
  permitted and the endpoint generates with `204 No Content`, so `E001` is a
  warning. See item 24. The guard against an empty object reaching the output
  still stands, as `X002` and `X003`.

### 13. Exit codes
`0` every endpoint generated, `2` the workbook could not be opened, `6` some
generated and some failed, `7` none generated. **6 is deliberately not zero**,
so a partial run cannot pass a build step quietly. Same codes in
`polymorphize_batch`.

### 14. The finding catalogue
The codes are stable and belong in an appendix. Families: `L` layout, `B`
banner, `S` sections, `A` attribute rows, `T` Data Type, `U` Usage, `M` SOR
columns, `V` variants, `E` emission, `G` generation, `D` documentation
(strict only), `X` faults in the tool itself.

Lenient reports only what stops an endpoint generating, which is what a run
enforces. Strict adds the house standard: a use case, a service domain, a
behaviour qualifier, a description on every attribute and an example on every
mandatory one. On the reference workbook as of v6.1: lenient 1 error and 20
warnings, strict 1 error and 134 warnings. Item 23 adds the merge codes.

### 15. The template is the contract
`samples/SOR_mapping_template.xlsx`, written by `polymorphize_template.py` or
`python polymorphize_cli.py template <path>`. Five sheets: **How to use**,
**Vocabulary**, **Example_SingleSor** (P1), **Example_MultiSor** (P2) and
**Operation_Template**. Dropdowns on Data Type and Usage are bound to the
Vocabulary sheet. It validates with zero findings at both levels and all
three operation sheets generate, which is a test (`W13`).

### 16. The measured result
The figures the documentation should quote, from the reference workbook:

| | Attributes a consumer could populate |
|---|---|
| As declared in the workbook | 287 |
| After removing what no SOR endpoint supplies | 158 (45% fewer) |
| After specialising per variant | 136 (53% fewer overall) |

The strongest single case is the `Card Details_Retrieve` response: 62
attributes declared, 30 supported by at least one SOR endpoint, at most 16 in
any one variant, and 6 in the leanest. A consumer of variant `02` sees six
attributes where the flat schema would have shown sixty-two.

Every generated specification passes `openapi-spec-validator` against OpenAPI
3.0.3.

### 17. Trim mode is now the secondary path
`polymorphize_core.py` still holds the original engine that takes an existing
specification plus a mapping workbook, and its tests still pass (D1 to D7, D9,
D11 in `tests/test_regressions.py`). It reads the older dotted-path workbook
format, not the Level format. The documentation must be clear about which
mode reads which format, or analysts will feed the wrong workbook to the wrong
entry point. D8 and D10 of that suite were retired because the dotted-path
generator and the old validator have been replaced.

### 18. Open question for Colin
`Card Issued Device_Retrieve` selects its SOR endpoint by which identifier the
caller supplies, and the sheet does not state the rule. `oneOf` without a
discriminator over mutually exclusive identifier groups is the current
reading. It needs confirming before that endpoint is published, and the answer
should go into the manual as the P3 worked example.


---

## New in v6.1

### 19. One specification, not one per sheet
This is the headline change and it inverts what items 8 and 11 describe.
Polymorphism only means anything inside a single namespace, so the default
output is now **one specification for the whole workbook**. One file per sheet
survives as `--split` on the CLI, a radio button in the GUI and a `split`
column in the batch manifest, and is for debugging a single endpoint. Say
plainly in the manual that cross-operation specialisation cannot apply in
split mode, because each file has its own namespace.

### 20. Two axes of variance, and why only one takes a discriminator
This is the conceptual heart of the documentation and the presentation.

| Axis | Question | Settled | Mechanism |
|---|---|---|---|
| Which operation | Which endpoint am I calling? | Design time: the path decides before a byte is sent | `allOf` specialisation of the shared group, **no discriminator** |
| Which variant | Which SOR endpoint serves this call? | Run time: the client sets `requestVariant` | `oneOf` with `discriminator: requestVariant` |

The operation axis takes no discriminator because there is no runtime
ambiguity to resolve, and a discriminator would demand a tag property that no
SOR supplies and no client can meaningfully set. The variant axis takes one
because the client genuinely chooses. The two compose: 14 of the shared groups
in the reference workbook sit inside a variant-driven operation.

Slides 12 and 13 need rewriting around this table rather than around either of
the previous positions.

### 21. Hoisting, and why it makes both axes one calculation
Every group with children is published as a named component, so a property's
value in any schema is either a scalar fragment or a `$ref`. Both axes then
reduce to the same operation: the base declares a property when its value is
identical everywhere, and everything else goes into each derived schema's
delta. This is sound because `allOf` conjoins rather than overrides.

Two consequences worth stating:

- **No component is ever an array.** An array publishes as
  `{type: array, items: {$ref: ...}}` pointing at its item, because there is
  no way to specialise an array through `allOf`. An `Array` row with a single
  `object` row beneath it takes that row's name; an array with several rows
  beneath it gets `<Name>Item`.
- **Shapes are compared structurally.** Type, format, length, obligation and
  enum count; descriptions, examples and SOR field names do not, because
  fragmenting the namespace on prose would defeat the reuse. The
  documentation is folded back in at emission, and an SOR field that differs
  between operations is recorded as `x-sor-field-by-operation`.

### 22. The measured result on the reference workbook
Quote these figures.

| | |
|---|---|
| Attributes a consumer of one operation is shown, per-sheet output | 137 |
| The same, merged with cross-operation specialisation | 137 |
| The same, merged **without** specialisation | 188 |

Merging is the requirement; specialising is what stops the merge costing
anything. A naive merge that mapped each group name to one schema holding the
union of its shapes would show consumers **51 extra attribute occurrences**,
37% more, that the SOR behind their operation does not supply. The worst single
case is the `Card Details_Retrieve` response: 16 attributes specialised
against 30 naive.

Structurally: 11 operations, 122 schemas, 50 groups hoisted, 13 specialised
across operations. Per-sheet output for comparison is 41 schemas across 11
files. The schema count rises because groups become named components instead
of being inlined; that is the price of reuse and traceability.

### 23. New findings, and the fact that they are actionable
The merge added five codes, and they are the most useful diagnostics the tool
produces because each one names a workbook change that yields more reuse.

| Code | What it says |
|---|---|
| `N001` | Two names differ only by case, so in one namespace they become separate schemas one keystroke apart. Reported, never normalised: renaming would silently change published field names |
| `N002` | The same service domain is spelled several ways. The tag is folded to the most-used spelling, since a tag is not part of the payload |
| `P001` | A shared group could not get a base. Says whether the shapes share no attribute at all, which is legitimate, or share one declared inconsistently, which is fixable |
| `P002` | One attribute of a shared group is declared two or more different ways, with both declarations spelled out |
| `Z002`, `Z003` | Two sheets claim the same message schema name, or the same path and method. Errors: the merge is refused rather than one operation silently overwriting another |

On the reference workbook: 1 `N001`, 1 `N002`, 6 `P001`, 5 `P002`, no errors.
The `N001` case, `AccountIdentifier` against `Accountidentifier`, is the single
correction that would collapse the largest family.

### 24. An empty message is permitted when the label is present
Changed from v6.0 on Colin's instruction. A `Request Body` or `Response Body`
label must be there, but the section may be empty. An empty request emits no
`requestBody` and keeps the natural verb, since there is no body to carry; an
empty response emits `204 No Content`. `E001` became a warning. A *missing*
label is still a hard failure, which is `S001` and `S003`.

The practical effect: `Card RSAEncrypted` now generates instead of failing, so
the reference workbook is 11 of 12 rather than 10 of 12, and only the
mislabelled sheet fails. The manual's troubleshooting section needs
rewriting: `E001` is no longer an example of a failure.

### 25. Drag and drop
Both file fields accept a dropped file or folder, using `tkinterdnd2` when it
is installed. It is optional: without it the window still opens and the fields
still work through Browse, and the hint text drops the drag wording. Document
the install line, and that a workbook dropped on the output field or a folder
dropped on the workbook field is corrected rather than refused.

Also new in the template: a `Method:` banner row. The HTTP method otherwise
comes from the BIAN action term, so a Retrieve becomes a GET and cannot carry
a body; stating POST in the Method cell is the supported way to fix that, and
it is why the template no longer raises `G001`.


---

## New in v6.2

### 26. The missing half of the argument: verification against the SOR
Until now the workbook's assertion that an element is backed by an SOR field
was accepted without evidence. A typo, a renamed field or a stale mapping all
looked exactly like a real mapping, and the element was published as though
the SOR supported it. That is the defect the tool exists to remove, so the
assertion is now checked against the SOR's own specification.

The four rules, settled with Colin and not to be changed without asking again:

1. **An unresolved field is excluded from the interface**, exactly as an empty
   SOR cell is, and every exclusion is reported with the sheet, the cell, the
   field name and the endpoint that was searched.
2. **Matching is on the leaf name alone.** `accountInfo.currency` and
   `currency` both resolve to any element named `currency` anywhere in the
   message. The most forgiving rule, chosen deliberately; the cost is that a
   name occurring in more than one place is ambiguous, and those are resolved
   and reported rather than silently taken.
3. ~~A disagreement about which endpoint applies is not guessed.~~
   **Superseded in v6.3, item 33: the banner is definitive.**
4. **Request Body is checked against the SOR request**, meaning its body for
   POST and PUT or its parameters for GET, and **Response Body against the SOR
   response**.

Verification is optional. Without SOR files the tool behaves as before and
every report says, prominently, that nothing was verified. Document that
difference as a difference in the strength of the guarantee, not a
configuration detail.

### 27. Why the banner needed a parser
The reference workbook writes the same endpoint six ways: with a method and a
colon, with a method and no colon, with neither, with and without a leading
slash, with a path parameter, and with several endpoints run together
separated by spaces or newlines. All of that is folded away, and `{cardNumber}`
becomes `{}` so a template matches whatever the SOR calls its parameter. Worth
a short section in the manual, because it tells an analyst they do not have to
be neat, only consistent between the banner and the column headers.

### 28. What verification found in the reference workbook
Against `samples/sor_creditcard_fixture.yaml`:

| | |
|---|---|
| Endpoints generated | 10 of 12 |
| Failures | `Card Issued Device_Retrieve` (S001, mislabelled section) and `Card RSAEncrypted` (**R001, the banner and the column header name different endpoints**) |
| SOR field names checked | 296 |
| Removed because the named field does not exist | 2 |
| Ambiguous, matched in more than one place | 12 |

`Card RSAEncrypted` is the new failure and it is a genuine workbook defect: the
banner says `GET /v1/card/{cardNumber}/cardRSAEncrypted`, the column header
says `SOR /v1/card/cancel`. One of them is wrong.

The twelve ambiguities are the honest cost of leaf-name matching. `value` and
`description` each occur in three places in one SOR response, so the tool says
which element it matched and lists the alternatives rather than pretending to
know.

### 29. The `R` finding family
| Code | What it says |
|---|---|
| `R001` | The banner and a column header name different endpoints. Error, the operation fails |
| `R002` | The banner names several endpoints and a column names none, so the pairing is undecidable. Error |
| `R003` | No endpoint is named anywhere. Error |
| `R004` | The endpoint is not in any supplied SOR specification. Error |
| `R005` | The workbook states a method the SOR does not define. Warning, the defined one is used |
| `R010` | A named SOR field does not exist, so the element was removed. Warning, with the cell |
| `R011` | A field name matched in more than one place. Warning, the element is kept |
| `R020` | Not one field name on the sheet resolved, which means the wrong endpoint or the wrong SOR file. Error |
| `R021` | The per-sheet total of removals. Warning |

### 30. The showcase, and why it is the deliverable
`<name>_showcase.html` is written beside the specification and is the
validation instrument: for every class and every endpoint it lists every
element the workbook declared, and for each one either that it is published or
why it is not. Five dispositions: **Published**, **Variant only**, **Not in
the SOR**, **No SOR field**, **Container**.

Two things about it are deliberate and should be said in the manual:

- Nothing in it is read back out of the generated specification. Every row is
  the workbook's own tree annotated with the decision the tool made, so it is
  an account of the reasoning rather than a description of the output. That is
  what makes it usable as evidence.
- It carries a sixth disposition, **LOST**, for an element that is mapped,
  verified and yet absent from the interface. It should always be zero and the
  report is loud when it is not, because a silent loss is the one failure mode
  this tool must not have.

The headline numbers it reports for the reference workbook: 287 elements
declared, 154 published, 133 removed, of which 131 have no SOR field in the
workbook and 2 name a field the SOR does not have.

### 31. New inputs everywhere
`--sor` on both `check` and `generate`, repeatable, taking a file or a folder.
A third field in the GUI with drag and drop, where a drop adds to the list
rather than replacing it, since several SOR files are normal. A `sor` column in
the batch manifest, resolved against the manifest's folder like every other
path, and a `--sor` override.

### 32. The SOR fixtures are fixtures
`samples/sor_creditcard_fixture.yaml` and `samples/sor_template_fixture.yaml`
are generated by `tools/build_sor_fixture.py` from the workbooks' own SOR
column values, so most of them resolve by construction. Three defects are
deliberate, to exercise the interesting paths: five leaf names omitted, one
duplicated, and one endpoint declaring only `PUT`. **They are not SOR
documentation and the manual must say so**, or someone will draw a conclusion
about the real system from them. Replace them with the real specifications
before the tool is used in anger.


---

## New in v6.3

### 33. The banner is definitive for the SOR endpoint
Reverses rule 3 of item 26 on Colin's instruction. The `SOR API Endpoint`
banner cell is the declaration of record and is the only place the SOR
endpoint is read from. An endpoint written on an SOR column header is a label:
it is compared with the banner only so that a difference can be reported as
`R001`, now a **warning**, and the banner is used regardless.

One banner endpoint serves every SOR column. Several are paired with the
columns in left-to-right order, the same positional convention the
`requestVariant` values already use, and a count mismatch is `R002`, an error,
rather than a guess. An empty banner is `R003`, an error, even where a column
header names an endpoint, because the header is not a fallback.

Effect on the reference workbook: `Card RSAEncrypted` now **generates**. Its
banner says `GET /v1/card/{cardNumber}/cardRSAEncrypted` and its column header
says `SOR /v1/card/cancel`; the banner wins, an `R001` warning records the
difference, and the sheet publishes. The reference workbook is therefore 11 of
12 again, with only the mislabelled `Card Issued Device_Retrieve` failing.

The manual's troubleshooting section needs rewriting once more: `R001` is no
longer an example of a failure.

### 34. The showcase is named after the service domain
`<service-domain>.html`, from the first service domain encountered in sheet
order, so the reference workbook produces `credit-card.html` rather than
`openapi_showcase.html`. A workbook spanning several domains is still one
interface and gets one name, hence "the first one". With no service domain
anywhere it falls back to `<merged name>_showcase.html`.

### 35. The showcase leads with the integration API
The report is about the APIs a consumer calls, so those are now its identity
rather than the sheet names:

- Each section is headed with the method and the integration path, for example
  **POST** `/CreditCard/Account/Retrieve`, with the sheet name as metadata.
- The navigation lists the integration APIs under that heading.
- The SOR endpoints have moved to a subordinate "Backed by, downstream" line.
  They are still there, because an analyst validating a mapping needs them,
  but they no longer read as the subject of the report.

### 36. Elements eliminated has a section of its own
Previously the eliminated elements were only visible greyed out among the
published ones. They now have a top-level section, because they are the point
of the tool: every one is a field a consumer would otherwise have seen, and
could have populated, with nothing behind it in the System of Record.

The section carries four counts, then one table per integration API listing
only the eliminated elements with their class, their path within the message,
the reason, the SOR field the workbook named and the detail. A second part
lists the elements **narrowed** rather than removed: backed by the SOR on some
variants of their operation and not others, so each variant publishes only its
own.

For the reference workbook: 296 elements declared, 155 published, 141
eliminated, of which 139 have no SOR field in the workbook and 2 name a field
absent from the SOR, plus 60 narrowed to some variants only.

### 37. The Windows build script
`build_windows_exe.bat` was still telling the user to look for `5.3` in the
title bar and still listed only four modules in its `.py.txt` rescue loop, so
a partial download could leave a stale module beside fresh ones. Rewritten: a
single `VERSION` variable at the top, all eleven modules in the rescue loop and
the presence check, and a `--hidden-import` for each module the GUI imports
lazily. Keep `VERSION` in step with `__version__` when the toolkit is next
bumped.


---

## What the pass covered, and what it did not

Carried out at release 6.3. Both Word documents were rewritten rather than
edited, because items 8 and 19 changed the input contract and the output shape
so far that no paragraph of the 5.x editions survived intact.

### `SOR_Polymorphizer_User_Manual.docx`

Rewritten for a reasonably technical reader. Covers items 1, 8, 9, 10, 11, 13,
14, 15, 16, 19, 20, 21, 22, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35 and
36. The centrepiece is chapter 9, an error guide organised by **what went wrong
in the workbook** rather than by finding code, on the reasoning that a message
describes a symptom while a correction depends on the cause, and that two
sheets can raise the same code for opposite reasons. Fifteen root causes, each
with what happened, how to confirm it, the correction and the consequence of
leaving it. Appendix A maps all 50 codes to the section that explains them.

Item 12 was rewritten rather than reproduced: `E001` and `R001` are no longer
examples of failure (items 24 and 33), so the only worked failure in the manual
is the `S001` mislabelled section. Item 6, the obsolete screenshots, was
resolved by removing every screenshot: the window is described in prose and by
its three tabs, so the manual does not go stale the next time the layout moves.

### `SOR_Polymorphizer_Technical_Guide.docx`

Written for maintainers. Covers items 2, 3, 4, 5, 11, 13, 14, 17, 20, 21, 22,
23, 29, 32 and 37, plus the module dependency order, the ten invariants with
the check group that guards each, the finding contract and the test strategy.

Chapter 8 is the prompt guide the request asked for: a standing preamble to
paste before every task, a table of which files to hand the assistant for which
kind of change and what the acceptance gate is in each case, a complete
paste-ready specification for recreating the toolkit from nothing, task prompts
for adding a finding, changing emission and diagnosing a failing check, a table
of the thirteen traps this codebase has actually produced with the instruction
that prevents each, a seven-point review checklist, and four things not to
delegate. Appendix B carries the measured baselines as regression anchors.

### Still open

1. **Item 18**, how `Card Issued Device_Retrieve` chooses between its three SOR
   endpoints. Recorded in Appendix C of the technical guide as an open
   question. Until it is answered, pattern P3 has no worked example in the
   manual and the only P3 sheet fails for an unrelated reason.
2. **Item 7**, the cheque field names on slide 11. Still unverified, still
   waiting on a cheque mapping workbook. Recorded in Appendix C.
3. **Items 2 and 20**, slides 12 and 13 of
   `Polymorphism_OpenAPI_for_BAs_v2_corrected.pptx`. Still teaching the 5.2
   discriminator position. The two-axis table they need is now written in
   chapter 3 of the technical guide and chapter 4 of the manual, so the slide
   rewrite is a transcription job rather than a research one.

Two newer decks, `SOR_Polymorphizer_Showcase.pptx` and
`BIAN_Interface_Independence.pptx`, were built at 6.3 and are current.


---

## New in 6.4

### 38. The plain concept name goes to the shape consumers reference
Raised in review as observation three, and correctly. Up to 6.3, when a group
split because the operations using it published different attributes, the plain
family name was reserved for the intersection **base** and every actual shape
took a suffix. Measured on the reference workbook, not one of the ten bases was
referenced by any property: the clean name belonged exclusively to the schema
nobody pointed at, and all 36 of the consumer-visible suffixed references
carried `For<Operation>` or `Profile<n>`.

That is the wrong way round. A code generator names its classes after component
schema names, and Swagger UI and Redoc list them, so the suffix reaches a
consuming developer even though the contract does not change. `assign_names`
now gives the plain name to the shape seen in the most places and names the
intersection base `<Name>Base`.

Effect on the reference workbook: suffixed property references fall from **36
of 68 to 23 of 68**, 53% to 34%. The example raised in the review reads as
`IssuedDeviceStatus` on `AvailableFundsRetrieve` where it previously read
`IssuedDeviceStatusForAvailableFundsRetrieve`; only the adjustment operation,
whose shape genuinely differs, still carries a suffix.

**The wire is unchanged, and this needs saying whenever the change is
described.** A property key comes from the workbook's Level column and the
schema name from the registry, so `"IssuedDeviceStatus": {"$ref": ...}` sends
the same JSON whatever the referenced schema is called. What does change is the
class names a generator produces, which is why this is 6.4 and not 6.3.1:
anyone who has generated a client from 6.3 output will see renamed classes.

An alternative tie-break, giving the plain name to the shape used by the most
distinct operations rather than the shape seen in the most places, was measured
and rejected: it improved the figure by one reference out of sixty-eight and
needed a longer rule to explain.

Three checks in `W16` guard this and should not be relaxed: that most property
references carry no tool-generated suffix, that every intersection base is
named with the `Base` suffix, and that the plain name each base gave up is a
published shape. Test count 499 to **501**.

### Still open from earlier releases
Items 7, 18 and the presentation slides are unaffected by this change and
remain as recorded above.


---

## Rules for the next build, set by Colin. Implemented in 6.5.

Recorded ahead of the work so they survive the conversation. The rules stand as
written below; what each one became in the code, and how the eight outstanding
decisions were settled, is recorded under **New in 6.5** at the end of this
file.

### R1. Build in support for the field mapping format
The dotted-path field mapping document, evidenced by
`apicoeissueddeviceadministrationfieldmappingv1.0.7.xlsx`: 22 sheets, one per
operation, banner in rows 1 to 9, header on row 10, columns Parameter Type,
Reusable API Field Name, Usage, Schema, SOR API Field Name (with the endpoint
appended to the header), Example, Description, Remarks, Required in Swagger.
Hierarchy is a dotted path up to eight segments deep. `core.load_mapping`
already reads it with no configuration: 22 layouts, 898 rows, 761 attributes,
137 section rows, 278 mapped, 483 unmapped.

### R2. Open in the most efficient way to peek
Measured, and this is an acceptance criterion rather than a preference. A
bounded read-only peek at the top rows of a few sheets costs 88 ms on the
field mapping workbook, 179 ms on the credit card workbook and 284 ms on the
8.7 MB CASA mapping. An ordinary `load_workbook` on that last file takes
**75.7 seconds**, a factor of 266, and would present as the window hanging on
file drop. So: `read_only=True`, bounded rows, bounded sheets, and a test that
holds the peek under a stated budget.

### R3. Classify the content, do not validate it
The check says what the file is and how much of it, never whether it is good.
No pass indicator, because a correctly classified workbook can still be full
of defects and a green tick would be read as a promise the check cannot keep.
Validation stays where it is. Classification must not block a run on content
grounds, and where it routes a run it says so rather than acting silently.

### R4. Drive the interface where required
Permission rather than obligation. On recognising a field mapping document the
window may reveal the field for the specification to trim, and hide it for a
Level workbook, so the detection lands while the user can still act on it. The
command line equivalent is a pre-flight that names the format and exits
non-zero when it cannot proceed, which also closes the defect below.

### The defect this closes
Today the tool run against a field mapping workbook classifies all 22 sheets as
support sheets, because the Level reader searches only the first six rows for a
header and this format's header is on row 10. It reports nought operation
sheets, nought errors, nought warnings, and **exits zero**. A build step would
go green having produced nothing. That is the worst available failure mode and
it argues for R4 independently of R1.

### Decisions still outstanding before this can be built
*All eight were settled before the 6.5 build. The rulings are recorded under
**New in 6.5**, item 40.*
1. **Which codebase.** Polymorphizer 6.4 or Substantiate 7.0. Building a second
   reader twice is waste, and the two lines have diverged.
2. **Depth of support.** Route the format to the trim engine only, or generate a
   new specification from it as the Level format does, or both.
3. **The variant axis.** This format carries one SOR field column on all 22
   sheets, so everything generated from it is pattern P1: no `requestVariant`,
   no `oneOf`, no discriminator. Accept that, or define how the format would
   express several SOR endpoints behind one operation.
4. **Parameter Type against the section banner.** The format carries both a
   per-row Parameter Type and section banner rows. Which is authoritative.
5. **`N/A` against blank** in the SOR field column, 216 and 289 rows
   respectively. Identical meaning, or `N/A` as decided and blank as unfinished.
6. **`Required in Swagger`**, blank on 607 of 783 rows, `Yes` on 132, `No` on 22.
   An instruction that overrides the mapping, or advisory.
7. **Findings.** Roughly half the codes are phrased in Level terms. Same codes
   with format-aware fix text, or a parallel family.
8. **Arrays.** Ten rows only, typed `array` in the Schema column with no item
   structure visible in the path. Needs one confirmed example.

### Working context, set by Colin
Unless Substantiate is named explicitly, every question, answer and build in
this workstream concerns the **SOR Polymorphizer** (6.4 line). Substantiate 7.0
remains a separate tree and is only in scope when called out by name. Recorded
here because this session has already crossed one context boundary and the
instruction must survive the next one.

### R5. Emit the full Apigee error set
Set by Colin against `apicoeissueddeviceadministrationswaggerv1.0.8.yaml`
(OpenAPI 3.0.3, 22 paths, 79 schemas). This closes review observation four,
and it supersedes the earlier proposal to adopt BIAN's six-response set: the
house standard is twelve, and it is evidenced.

**All twelve are Apigee errors.** Each response description carries the same
sentence, so no judgement is needed about which layer raises them: *"Disclaimer:
The error message is Apigee error only and not mapped to the System API error.
Consumers should not rely on the exact error text as it may change depends on
the format of the System API its calling."*

| Component | Code | Reason phrase |
|---|---|---|
| `BadRequest` | 400 | Bad Request |
| `Unauthorized` | 401 | Unauthorized |
| `Forbidden` | 403 | Forbidden |
| `NotFound` | 404 | Not Found |
| `MethodNotAllowed` | 405 | Method Not Allowed |
| `Conflict` | 409 | Conflict |
| `UnprocessableEntity` | 422 | Unprocessable Entity |
| `TooManyRequests` | 429 | Too Many Requests |
| `InternalServerError` | 500 | Internal Server Error |
| `BadGateway` | 502 | Bad Gateway |
| `ServiceUnavailable` | 503 | Service Unavailable |
| `GatewayTimeout` | 504 | Gateway Timeout |

Structurally all twelve are identical, verified: the same three response
headers, the same `ErrorResponse` schema, the same disclaimer. Every one of the
264 error entries across the 22 operations is a `$ref` into
`components/responses`, none inlined, which is the shape to reproduce.

**`ErrorResponse`**, to sit in `components/schemas`: `status` (string),
`title` (string), `timestamp` (string, date-time) and `errors`, an **array** of
objects carrying `realm`, `code`, `errordesc`, `detail` and `instance`.

**The three headers are not error-specific.** `x-BDO-Client-Request-Id`,
`x-BDO-Client-Request-Trace-Id` and `x-BDO-Client-Request-Span-Id` appear on the
200 responses too, so they belong in `components/headers` and on every response
the tool emits, not only on the failures.

**Do not copy the sample's example verbatim.** It does not conform to its own
schema: the schema declares `errors` as an array and the example provides
`error` as a single object, it omits `status` and `realm`, and its `title` reads
"The request was successful, but there is no content in the response", which is
a 204 message pasted into an error example. Write a conforming example, or emit
none. Raise this with the API COE as a defect in their sample rather than
silently correcting it.

**Open, and needs Colin's ruling.** The sample also declares four
`securitySchemes` (BasicAuth, BearerAuth, ApiKeyAuth, OAuth2) applied at root
level with no per-operation override. Whether the tool should emit those as well
is a separate question from the errors and has not been asked for.

---

## New in 6.5

The build that carried out R1 to R6. Test count 501 to **736** across the three
suites (516 + 156 + 64), with four new groups: W18 classification, W19 the
field mapping format, W20 the Apigee error set and W21 class and context
naming.

### 39. Two input formats, and the file decides which

`polymorphize_classify.py` is new and it runs before any reader. It answers one
question, which format is this, and it answers it on a bounded peek:
`read_only=True`, fourteen rows, thirty columns, at most six worksheets, under
a budget a test enforces. Measured on the reference workbooks that peek costs
284 ms where an ordinary open of the largest of them costs 75.7 seconds, so
classification is cheap enough to run on every entry point including the moment
a filename is typed into the desktop window.

A sheet showing `Level 1` in its header row is the Level format; a sheet showing
a `Parameter Type` column beneath the banner block is a field mapping document.
A workbook holding both resolves to the Level format and reports the mixture
rather than choosing quietly.

Three properties of the classifier are deliberate and each answers a rule:

- **It classifies, it does not validate** (R3). Its report says what the file is
  and how much of it, never whether it is any good. There is no pass indicator,
  because a correctly classified workbook can still be full of defects and a
  green tick would be read as a promise the check cannot keep.
- **It never raises.** An unreadable file, a corrupt archive or a password
  produce a classification carrying `readable=False` and the reason. Failing
  soft matters because this runs on the typing path in the window.
- **It drives the interface where it is useful** (R4). The desktop window shows
  the verdict under the workbook field and, on recognising a field mapping
  document, asks for the specification to trim. The command line equivalent is
  a pre-flight in `polymorphize_cli._preflight`, shared by `check` and
  `generate`.

### 40. The eight outstanding decisions, as settled

| | Question | Ruling |
|---|---|---|
| 1 | Which codebase | The Polymorphizer line, per the working context recorded above. Substantiate 7.0 is untouched |
| 2 | Depth of support | Full generation, not trim only. The format produces a specification exactly as the Level format does |
| 3 | The variant axis | Accepted as absent. One SOR column means every sheet is P1, and a sheet that looks as though it wants variants is told so once, as `F005` |
| 4 | Parameter Type against the banner | Parameter Type is authoritative and the banner corroborates it. A disagreement is reported as `F001`, never resolved silently |
| 5 | `N/A` against blank | Both unmapped, but not merged. `N/A` is a decision recorded and passes silently; a blank is a decision not yet stated and is reported at strict level as `F004` |
| 6 | `Required in Swagger` | Advisory. Blank on 607 of 783 rows, so it cannot carry an include or exclude instruction. Recorded as `x-required-in-swagger` and nothing else |
| 7 | Findings | One catalogue with format-aware fix text, plus a new `F` family for what only this format can get wrong. Not a parallel set: an analyst should not have to learn two vocabularies |
| 8 | Arrays | A container mapped but holding no published member is removed and reported as `A006`, rather than emitted as `properties: {}`. This was found by running the format, not by inspection: three sheets produced an empty schema until `supported()` was corrected to judge a container by its descendants |

### 41. The full Apigee error set (R5)

`polymorphize_errors.py` is new and holds the twelve responses, the
`ErrorResponse` schema, the three standard headers and the disclaimer, quoted
verbatim from the API COE specification. `attach()` is idempotent, adds the
twelve to every operation that does not already declare that status, puts the
three headers on success responses as well as errors, and creates the shared
components once.

Two departures from the sample were taken deliberately and both are argued in
the module docstring: the headers go on every response because the sample
carries them on its 200s, and the sample's `ErrorResponse` example is not
copied because it does not conform to its own schema. A conforming example is
written instead. The defect stands as a note for the API COE.

One ordering defect surfaced here and is worth recording, because it would
recur in any future addition to `components/schemas`: `attach()` has to run
**before** the `identify()` loop that stamps `x-class` and `x-context`, in both
`generate` and `merge`, or `ErrorResponse` ships without them. It did, in the
first cut.

### 42. Class and context are separate, and parseable

Colin asked for the class to be separated from its context by a `~` so that a
following process could parse the two apart. The separator shipped is `__` and
the reason is narrow rather than a matter of taste:

- The OpenAPI component key regex is `^[a-zA-Z0-9\.\-_]+$`, which does not
  admit `~`.
- RFC 6901 makes `~` the escape character in a JSON Pointer, so a `~` in a key
  must be written `~0` in every `$ref` that reaches it. `#/components/schemas/
  Account~CardDetailsRetrieve` does not point at `Account~CardDetailsRetrieve`.
- `openapi-spec-validator` accepts the unescaped form regardless, which is
  worse than a rejection, because the defect would ship silently and surface
  in whichever generator handles pointers correctly.

So `__` is the separator in the key, and the requirement, that a following
process can parse the two apart without a heuristic, is met by three extensions
on every schema: `x-class`, `x-context` and `x-qualified-name`, the last
carrying the `~` form for a consumer that prefers a single readable string.
Colin accepted this in preference to a literal `~` in the key.

### 43. The base class is present, verified rather than asserted

Colin asked for confirmation that the base class is emitted and not only the
subsidiary shapes, earlier versions having dropped it. Checked on the reference
workbook: 10 of 10 families that split carry a `<Class>__Base`. Seven families
carry no base, and the split is 4 legitimate, being families of one shape where
a base would be a duplicate, against 3 reported through `P001`.

Two of my own measurements in this area were wrong and were corrected to Colin
in the same session, which is recorded here so the corrected figures are the
ones that survive: 23 of 68 schemas carry a context suffix, 34%, not the 36 of
68 and 53% first stated. The first error assumed the plain name went to the most
*referenced* shape where the code orders by places seen; the second counted
array item schemas such as `AccountDetialsItem` as suffixed.

### 44. Exit code 3, and the failure mode it closes

A field mapping workbook handed to 6.4 classified all 22 sheets as support
sheets, because the Level reader searches only the first six rows for a header
and this format's header is on row 10. It reported nought operations, nought
errors, nought warnings, and **exited zero**: a build step would have gone
green having produced nothing. `EXIT_UNRECOGNISED = 3` closes it. A file the
tool cannot place is now refused, by name, with a non-zero code, in `check`,
`generate` and the batch runner, which gained an `unrecognised` status.

### 45. What the build produces on the field mapping reference

All 22 operations generate: 99 schemas, 35 groups hoisted, 13 specialised
across operations, valid OpenAPI 3.0.3. Every component key sits inside
`^[a-zA-Z0-9.\-_]+$` with no tilde anywhere, every schema carries `x-class`,
and all twelve error codes are present on all 22 operations. The credit card
baseline is unchanged at 11 of 12 with 123 schemas, which is the point of
running it: the second format was added without moving the first.

### 46. The specialisation figures are now reproducible

The figures the README quoted for what specialisation prevents were measured
once by a script that was not kept, so they could not be checked against this
build. `tools/measure_specialisation.py` is new, it counts from the workbook
tree so that the two figures differ in one respect only, and it runs on any
workbook. Current figures, from that script: the reference workbook 228 against
275, 47 occurrences prevented, 21% more; the field mapping reference 354 against
372.

The earlier 137 against 188 is not carried forward, because it cannot be
reproduced and a number that cannot be reproduced should not be quoted.

### 47. Output from 5.3 has been archived rather than left in place

`samples/casa_*_v53.*` predated the naming convention, the plain-name rule and
the error set, so read as a baseline they contradict 6.5. They now sit in
`samples/archive_v5.3/` behind a note saying what they are and what they
predate. Nothing in the tool or the tests reads them, and the workbook they
came from is not part of the sample set, so they cannot be regenerated.

### Still open after 6.5
*Items 3 and 5 were closed in 6.6. See item 55.*

1. **Item 7**, the cheque field names on slide 11, still unverified because no
   cheque mapping workbook has been supplied.
2. **Item 18**, unchanged.
3. **The two Word documents** describe 6.3 behaviour. They do not cover the
   second input format, classification, the error set or class and context, so
   they need a pass rather than an edit.
4. **The presentations** still show the pre-composition mechanism.
5. **Security schemes.** The API COE sample declares four `securitySchemes`
   applied at root with no per-operation override. Whether the tool should emit
   them is a separate question from the errors and has not been asked.

---

## Rules for 6.6, set by Colin

Four instructions, each recorded with what it became.

**Batch and window only, not the command line.** The export is written by
`polymorphize_batch` and by the desktop window. `gen.run` takes an `export`
flag that defaults off, the command line never sets it, and
`generate --help` carries an epilog saying where the export lives rather than
leaving its absence to be discovered.

**The version to create the spreadsheet from is determined, so the
spreadsheet matches the swagger.** The document is written by the run that
wrote the specification, from the same trees, and its provenance sheet names
the specification file and carries its SHA-256 fingerprint. A later reader can
tell whether the specification has been regenerated since.

**Show the differences between what was requested and what the swagger
contains.** Three columns beyond the format's nine.

**Round tripping, bringing all of the original spreadsheet along.** The export
is a valid input and carries every row of the source, including the rows the
tool does not read.

**Each SOR endpoint is a separate sheet.**

**The description of each operation should read
`**<u>API Lifecycle Status - Design</u>**`.**

---

## New in 6.6

Test count 736 to **787** across the three suites (556 + 156 + 75), with three
new groups and three new regression sections: W22 the lifecycle status, W23
the export, W24 the round trip, D12 the showcase reader, D13 a failed sheet
carried, D14 a correction reaching the next specification.

### 48. The field mapping document, written back out

`polymorphize_export.py` is new. It writes the API COE's own format from the
run that produced the specification, so the two cannot drift: there is no
second read of the workbook and no second interpretation of it. The verdict in
the Publication Status column is the same verdict the showcase draws, taken
from the same function, because computing it twice by two routes is how a
report and a spreadsheet come to tell an analyst different things about the
same cell.

The nine columns of the format are unchanged and the header stays on row 10,
so the document opens as the format the API COE circulates. The three added
columns are Publication Status, Why and Analyst Action, and a row that did not
reach the interface is shaded so the rows worth attention are the ones the eye
lands on.

### 49. Why the round trip needed the whole document, not the parts we read

The tool reads Request Body and Response Body and nothing else. An export
composed only from what it reads would drop every header row, every parameter
row and every column beyond the nine, so an analyst who corrected a flagged
row and fed the workbook back would lose a third of their document on each
pass. That is a worse failure than the one the export exists to fix.

So the field mapping reader now keeps the rows it read, on
`SheetResult.source_rows`, and the exporter writes them back **verbatim**. The
same list, not a copy, so it costs nothing. Verbatim also means the tool does
not rewrite an analyst's `String (10)` into its own `string(10)`: a comparison
between two rounds shows what actually changed and nothing else.

Two consequences worth stating:

* **A failed sheet is still exported**, marked `Not generated` with the
  finding and its correction against it. Losing an operation on the round trip
  would be worse than any finding it carries. `D13` holds this.
* **The export is a fixed point.** Exporting an export changes not one cell.
  `W24` holds this, and it caught the defect it was written for: the three
  added columns were being carried as unknown extra columns on the way back
  in, so a document on its fifth pass would have been fifteen columns wider
  than the format allows.

### 50. Where the format cannot carry what the Level format holds

A Level workbook has no Remarks column, no Required in Swagger column, no API
Name and no SOR Name, and no header rows at all. Exporting one is a
**conversion**, not a round trip, and the provenance sheet says so in those
words and lists the three losses. API Name is derived from the sheet name
rather than left blank, and recorded as derived. SOR Name is left empty
because there is nothing to derive it from.

A Level sheet that failed to read is not exported at all, because there is no
tree to compose from and nothing to carry. The provenance sheet names it.

### 51. The standard header block, and an inconsistency it settles

The 220 header rows across the reference document's 22 sheets are seven
distinct headers repeated, so the block is a constant taken from that
document and supplied where the source had none.

Those seven are not declared consistently in the source, which is worth
raising with the API COE: `x-BDO-Application-Id` appears as both `string` and
`string(10)`, three of them appear once with a length and a Usage and once
with neither, and three rows carry `CVV`, `auxiliaryPan` and
`auxiliaryExpiry` in the Required in Swagger column, which is paste drift from
the neighbouring column. The constant takes the most complete declaration of
each, so an exported document is tidier than the hand maintained one.

### 52. One sheet per SOR endpoint, and what that exposed

The format has one SOR API Field Name column, so a Level sheet naming several
SOR endpoints becomes several sheets. The reference workbook's eleven
generated operations become seventeen sheets, `Card Details_Retrieve` alone
becoming five.

Writing the banner for each of those sheets exposed a rule worth recording.
The first cut took the endpoint from the SOR column header, and on one sheet
of the reference field mapping document the column header and the banner name
different endpoints. The banner is the definitive statement and the column
header is a label, a rule 6.3 already established for reading, so rewriting
the banner from the column would have changed which SOR the next round
verified against. A mapping sheet's banner is now reproduced exactly as
written. A Level banner naming several endpoints pairs them with the columns
in order, so sheet *n* takes fragment *n*, keeping the analyst's own method
prefix and spelling.

### 53. The API lifecycle stage

Every operation now carries:

```yaml
description: '**<u>API Lifecycle Status - Design</u>**'
x-api-lifecycle-status: Design
```

The description is the statement and nothing else, as instructed, so it
renders as one line. The use case is not displaced: it was already the
summary, and `x-use-case` now carries it untruncated, so nothing is lost to
make room. `Design` is a module constant rather than an option, because the
tool generates a design and an operation at a later stage is not something it
produced. The bold-and-underline is markdown emphasis around raw HTML, which
CommonMark permits and which Swagger UI and Redoc both render.

This closes review observation five.

### 54. A defect in 6.5 that the export uncovered

The showcase re-read the workbook rather than using the run's own results, and
it re-read it with the **Level** reader whatever the format was. For a field
mapping document every sheet came back as a support sheet, so the showcase
listed no elements at all: the 6.5 sample report ran to 40 KB against the
credit card report's 323 KB and that went unnoticed, because nothing checked
that a report had content.

The run now carries the results it read, on `RunResult.results`, and the
showcase and the export both use them. The re-read survives only for a caller
that builds a `RunResult` by hand, and it dispatches on the format. `D12`
holds it, and asserts a floor on the number of elements the report sees rather
than merely that it was written.

### 55. What this closes from the client review

Observation two, generated field mapping documentation, and observation five,
the API Lifecycle Design status. Both were listed as not implemented in the
response sent to the client.

### Still open after 6.6

1. **Item 7**, the cheque field names on slide 11, still unverified because no
   cheque mapping workbook has been supplied.
2. **Item 18**, unchanged.
3. **The two Word documents** describe 6.3 behaviour and now trail by three
   releases. They cover neither input-format classification, nor the error
   set, nor class and context, nor the export and its round trip. They need a
   pass rather than an edit.
4. **The presentations** still show the pre-composition mechanism.
5. **Security schemes.** The API COE sample declares four `securitySchemes`
   applied at root with no per-operation override. Still not asked for.
6. **A later lifecycle stage.** `Design` is fixed. If a specification ever has
   to be published at Build or Live, the constant becomes an option and the
   question of who is entitled to set it has to be answered first.
