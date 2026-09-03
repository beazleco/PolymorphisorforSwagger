# Notes for the deferred documentation pass

Documentation has deliberately not been updated. This file records everything
the pass must cover. Items 1 to 7 were carried over from v5.3; items 8 to 18
came in with the Level-format contract in v6.0; items 19 to 25 came with the
single merged specification in v6.1; items 26 to 32 came with verification
against the System of Record and the HTML showcase in v6.2; items 33 to 36 are
new in v6.3.

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
39 → 49 → 62 → 93 → 271 → 343 → 486 → **499** across three suites
(279 + 156 + 64). `tests/run_all.py` runs all three.

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
