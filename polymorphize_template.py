"""
polymorphize_template — writes the SOR mapping workbook template.

Version 6.6.

The template is the contract. It carries the header row the reader looks for,
the banner labels it recognises, the two section labels it consumes, and
dropdowns on every column with a controlled vocabulary, so that a workbook
built from it validates clean before anyone runs a generation.

Sheets written
==============

How to use
    The contract in prose, including what is read and what is ignored.
Vocabulary
    The lists behind the dropdowns. Do not rename or reorder.
Example_SingleSor
    A worked operation with one SOR endpoint. Pattern P1.
Example_MultiSor
    A worked operation with three SOR endpoints selected by ``requestVariant``.
    Pattern P2, the case that produces the largest reduction.
Operation_Template
    A minimal valid skeleton to copy for each new endpoint.
"""

from __future__ import annotations

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import polymorphize_workbook as wbk

__version__ = "6.6"

LEVELS = 6

#: Row 1, exactly as the reader expects to find it.
HEADER = (["ValueType"] + ["Level %d" % i for i in range(1, LEVELS + 1)] +
          ["Dummy Value", "Data Type", "Usage", "Descriptions"])

#: Banner labels, rows 2 to 6, in the label column.
BANNER = [
    ("Use Case:", "What this endpoint is for, in one sentence."),
    ("Service Domain:", "BIAN service domain"),
    ("Equivalent BIAN API Endpoint:", "/ServiceDomain/BehaviourQualifier/Action"),
    ("Proposed Business API Endpoint:", "/service-domain/resource/action"),
    ("SOR API Endpoint:", "GET: /v1/example/resource"),
    ("Method:", "POST"),
]

TYPES = ["object", "Array", "String (35)", "String (3)", "Number(12)",
         "integer", "boolean", "Date", "DateTime"]
USAGES = ["Mandatory", "ConditionalMandatory", ""]

FONT = "Arial"
FILL_HEADER = PatternFill("solid", fgColor="FF1F3864")
FILL_BANNER = PatternFill("solid", fgColor="FFE7EEF7")
FILL_SECTION = PatternFill("solid", fgColor="FFD6E4F0")
FILL_ANALYST = PatternFill("solid", fgColor="FFFFF6D6")
FILL_SOR = PatternFill("solid", fgColor="FFEAF4E6")
THIN = Side(style="thin", color="FFB8C4D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _sheet(wb, title):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    return ws


def _put(ws, row, col, value, *, bold=False, fill=None, wrap=False, size=10,
         colour=None, border=True):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name=FONT, size=size, bold=bold,
                  color=colour or ("FFFFFFFF" if fill is FILL_HEADER else "FF1A1A1A"))
    if fill is not None:
        c.fill = fill
    if border:
        c.border = BORDER
    c.alignment = Alignment(vertical="top", wrap_text=wrap)
    return c


def _write_header(ws, sor_endpoints):
    """Row 1: the fixed columns then one column per SOR endpoint."""
    for i, label in enumerate(HEADER, start=1):
        _put(ws, 1, i, label, bold=True, fill=FILL_HEADER, wrap=True)
    first_sor = len(HEADER) + 1
    for j, endpoint in enumerate(sor_endpoints):
        _put(ws, 1, first_sor + j, "SOR %s" % endpoint, bold=True,
             fill=FILL_HEADER, wrap=True)
    ws.column_dimensions["A"].width = 30
    for i in range(2, 2 + LEVELS):
        ws.column_dimensions[get_column_letter(i)].width = 26
    widths = {"Dummy Value": 20, "Data Type": 15, "Usage": 21, "Descriptions": 46}
    for name, width in widths.items():
        ws.column_dimensions[get_column_letter(HEADER.index(name) + 1)].width = width
    for j in range(len(sor_endpoints)):
        ws.column_dimensions[get_column_letter(first_sor + j)].width = 26
    ws.freeze_panes = "B2"
    return first_sor


def _write_banner(ws, values=None):
    """Rows 2 to 6: label in column A, value in column B."""
    for i, (label, hint) in enumerate(BANNER, start=2):
        _put(ws, i, 1, label, bold=True, fill=FILL_BANNER)
        _put(ws, i, 2, (values or {}).get(label, hint), fill=FILL_ANALYST, wrap=True)
    # The behaviour qualifier sits beside the service domain.
    _put(ws, 3, 3, (values or {}).get("BQ:", "BQ: Behaviour Qualifier"),
         fill=FILL_ANALYST)
    return len(BANNER) + 2


def _write_rows(ws, start_row, section, rows, first_sor, n_sor):
    """A section label then its attribute rows.

    ``rows`` are ``(depth, name, example, type, usage, description, [sor...])``
    with ``depth`` 1-based against the Level columns.
    """
    r = start_row
    _put(ws, r, 1, section, bold=True, fill=FILL_SECTION)
    for i in range(2, len(HEADER) + n_sor + 1):
        _put(ws, r, i, None, fill=FILL_SECTION)
    r += 1
    for depth, name, example, dtype, usage, desc, *sor in rows:
        _put(ws, r, 1 + depth, name, bold=(depth == 1))
        _put(ws, r, HEADER.index("Dummy Value") + 1, example or None, fill=FILL_ANALYST)
        _put(ws, r, HEADER.index("Data Type") + 1, dtype or None, fill=FILL_ANALYST)
        _put(ws, r, HEADER.index("Usage") + 1, usage or None, fill=FILL_ANALYST)
        _put(ws, r, HEADER.index("Descriptions") + 1, desc or None,
             fill=FILL_ANALYST, wrap=True)
        for j in range(n_sor):
            _put(ws, r, first_sor + j, (sor[j] if j < len(sor) else None) or None,
                 fill=FILL_SOR)
        r += 1
    return r


def _validations(ws, last_row, n_sor):
    """Dropdowns on Data Type and Usage, sourced from the Vocabulary sheet."""
    type_col = get_column_letter(HEADER.index("Data Type") + 1)
    usage_col = get_column_letter(HEADER.index("Usage") + 1)
    dv_type = DataValidation(
        type="list", formula1="=Vocabulary!$A$2:$A$%d" % (len(TYPES) + 1),
        allow_blank=True, showDropDown=False)
    dv_type.error = ("Use a type from the list, optionally with a length in "
                     "brackets, for example String (19).")
    dv_type.errorTitle = "Data Type not recognised"
    dv_type.promptTitle = "Data Type"
    dv_type.prompt = ("object for a nested group, Array for a repeating group, "
                      "or a scalar with an optional length.")
    dv_usage = DataValidation(
        type="list", formula1="=Vocabulary!$B$2:$B$%d" % (len(USAGES) + 1),
        allow_blank=True, showDropDown=False)
    dv_usage.error = ("Mandatory, ConditionalMandatory, or leave the cell blank "
                      "for optional.")
    dv_usage.errorTitle = "Usage not recognised"
    dv_usage.promptTitle = "Usage"
    dv_usage.prompt = ("Mandatory lands in the required list. "
                       "ConditionalMandatory is documented but not required. "
                       "Blank means optional.")
    ws.add_data_validation(dv_type)
    ws.add_data_validation(dv_usage)
    dv_type.add("%s2:%s%d" % (type_col, type_col, last_row + 200))
    dv_usage.add("%s2:%s%d" % (usage_col, usage_col, last_row + 200))


# --------------------------------------------------------------------------- #
# Worked examples
# --------------------------------------------------------------------------- #

SINGLE_SOR = "/v1/deposit/account/balance"

SINGLE_REQUEST = [
    (1, "DepositAccountRetrieveRequest", "", "object", "", "Request wrapper."),
    (2, "AccountIdentifier", "", "object", "", "Identifies the account."),
    (3, "accountIdentificationType", "IBAN", "String (12)", "",
     "Type of identifier supplied. Not held in SOR, so it is dropped."),
    (3, "accountIdentification", "GB29NWBK60161331926819", "String (34)",
     "Mandatory", "The account number.", "accountNo"),
    # AccountIdentifier is declared identically on the other example sheet, so
    # the two operations share one component instead of forcing a base and two
    # derived schemas. That is the reuse the P001 and P002 warnings ask for.
    (2, "accountCurrencyCode", "GBP", "String (3)", "", "ISO currency code.",
     "ccy"),
]

SINGLE_RESPONSE = [
    (1, "DepositAccountRetrieveResponse", "", "object", "", "Response wrapper."),
    (2, "AccountBalance", "", "object", "", "Balance group."),
    (3, "balanceType", "Available", "String (20)", "", "Kind of balance.",
     "balType"),
    (3, "balanceAmount", "1250.75", "Number(18)", "Mandatory", "Balance value.",
     "balAmt"),
    (3, "balanceCurrencyCode", "GBP", "String (3)", "", "Currency of the balance.",
     "balCcy"),
    (2, "balanceAsOfDateTime", "2026-08-31T23:59:59Z", "DateTime", "",
     "When the balance was struck.", "asOf"),
]

MULTI_SOR = ["/v1/account/contact", "/v1/account/info", "/v1/card/account/list"]

MULTI_VARIANT_DESC = ("01. Account Contact Details\n"
                      "02. Account Information\n"
                      "03. Account Card List")

MULTI_REQUEST = [
    (1, "AccountRetrieveRequest", "", "object", "", "Request wrapper."),
    (2, wbk.VARIANT_PROPERTY, "01", "String (2)", "Mandatory", MULTI_VARIANT_DESC),
    (2, "AccountIdentifier", "", "object", "", "Identifies the account."),
    (3, "accountIdentification", "00143085930", "String (34)", "Mandatory",
     "The account number.", "account", "account", "account"),
    (2, "accountCurrencyCode", "PHP", "String (3)", "", "ISO currency code.",
     "currencyCode", "currencyCode", "currencyCode"),
    (2, "includePhoneNumbers", "true", "boolean", "ConditionalMandatory",
     "Only the contact endpoint honours this.", "includePhone"),
    (2, "includeCardList", "true", "boolean", "ConditionalMandatory",
     "Only the card list endpoint honours this.", "", "", "includeCards"),
]

MULTI_RESPONSE = [
    (1, "AccountRetrieveResponse", "", "object", "", "Response wrapper."),
    (2, "accountIdentification", "00143085930", "String (11)", "Mandatory",
     "Echoed account number. Common to every variant.",
     "account", "account", "accountNumber"),
    (2, "accountName", "J DOE", "String (70)", "", "Only the info endpoint.",
     "", "accountName"),
    (2, "PhoneNumber", "", "object", "", "Only the contact endpoint."),
    (3, "phoneNumberType", "Mobile", "String (12)", "", "Kind of number.",
     "phoneType"),
    (3, "phoneNumber", "+639171234567", "String (20)", "", "The number itself.",
     "phoneNo"),
    (2, "CardList", "", "Array", "", "Only the card list endpoint.",
     "", "", "cards"),
    (3, "cardTokenNumber", "15555000000002588", "String (19)", "",
     "Tokenised card identifier.", "", "", "cards.cardToken"),
    (3, "cardExpiryDate", "2311", "String (4)", "", "Expiry, YYMM.",
     "", "", "cards.expiryDate"),
]

BLANK_REQUEST = [
    (1, "MyOperationRequest", "", "object", "", "Rename this to your request schema."),
    (2, "MyIdentifier", "", "object", "", "Delete or rename. Groups are objects."),
    (3, "myIdentification", "12345", "String (20)", "Mandatory",
     "One mapped attribute is the minimum for a valid endpoint.", "sorFieldName"),
]

BLANK_RESPONSE = [
    (1, "MyOperationResponse", "", "object", "", "Rename this to your response schema."),
    (2, "myReturnedValue", "example", "String (35)", "",
     "One mapped attribute is the minimum for a valid endpoint.", "sorFieldName"),
]


# --------------------------------------------------------------------------- #
# Prose sheets
# --------------------------------------------------------------------------- #

HOW_TO_USE = [
    ("The SOR mapping workbook", 16, True),
    ("", 10, False),
    ("One worksheet describes one endpoint. This workbook produces one "
     "OpenAPI specification per operation sheet, so add a sheet per endpoint "
     "rather than adding rows to an existing one.", 10, False),
    ("", 10, False),
    ("Layout", 13, True),
    ("Row 1 is the header row. Do not move it, rename its columns or reorder "
     "them. The Level columns express nesting: put a name in Level 1 for the "
     "schema, Level 2 for its attributes, Level 3 for the attributes of a "
     "Level 2 object, and so on. Never skip a Level.", 10, False),
    ("", 10, False),
    ("Rows 2 to 6 are the banner. The label goes in column A and the value in "
     "column B. The behaviour qualifier goes in column C beside the service "
     "domain, written as BQ: followed by the name.", 10, False),
    ("", 10, False),
    ("What is read", 13, True),
    ("Only two sections are read: Request Body and Response Body. Put the "
     "label in column A on the row above the schema.", 10, False),
    ("", 10, False),
    ("What is ignored", 13, True),
    ("Request Header, Response Header and any parameter section are skipped. "
     "So is anything below the mapping grid: leave three or more blank rows "
     "and you can paste sample payloads, notes or working there without "
     "affecting the output.", 10, False),
    ("", 10, False),
    ("If a sheet has no Request Body section the endpoint fails and no "
     "specification is written for it. The rest of the workbook still "
     "generates. The most common cause is a Request Body section labelled "
     "Request Header by mistake.", 10, False),
    ("", 10, False),
    ("SOR columns", 13, True),
    ("Add one column per SOR endpoint, headed SOR followed by the endpoint "
     "path. Put the SOR field name in the cell when that endpoint supplies "
     "the attribute, and leave it empty when it does not. An empty cell is "
     "how an attribute gets removed from the published interface, which is "
     "the entire point of the exercise.", 10, False),
    ("", 10, False),
    ("Several SOR endpoints on one sheet", 13, True),
    ("When one operation is served by several SOR endpoints, add a "
     "requestVariant attribute to the request and list the variants in its "
     "Descriptions cell, one per line, as 01. Name. The nth variant must "
     "describe the nth SOR column. Each variant then publishes only the "
     "attributes its endpoint supplies, inheriting the common ones through "
     "allOf, and the request carries a discriminator so a consumer knows "
     "which shape applies.", 10, False),
    ("", 10, False),
    ("If the endpoint is chosen by which identifier the caller supplies "
     "rather than by a tag, leave requestVariant out. The variants are then "
     "published as mutually exclusive alternatives without a discriminator.",
     10, False),
    ("", 10, False),
    ("The method", 13, True),
    ("The HTTP method comes from the BIAN action term at the end of the "
     "equivalent BIAN endpoint, so a Retrieve becomes a GET. A GET cannot "
     "carry a request body, so where an operation reads with a body, state "
     "POST in the Method cell. Leave Method blank to take the BIAN action "
     "term.", 10, False),
    ("", 10, False),
    ("Vocabularies", 13, True),
    ("Data Type: object, Array, or a scalar with an optional length in "
     "brackets. Usage: Mandatory, ConditionalMandatory, or blank for "
     "optional. Both have dropdowns.", 10, False),
    ("", 10, False),
    ("Before you hand the workbook over", 13, True),
    ("Run the Workbook check tab. Lenient reports only what would stop an "
     "endpoint generating. Strict adds the house standard: a use case, a "
     "service domain, a behaviour qualifier, a description on every "
     "attribute and an example on every mandatory one.", 10, False),
]


def _write_how_to(ws):
    ws.column_dimensions["A"].width = 112
    r = 1
    for line, size, bold in HOW_TO_USE:
        c = ws.cell(row=r, column=1, value=line or None)
        c.font = Font(name=FONT, size=size, bold=bold,
                      color="FF1F3864" if bold and size >= 13 else "FF1A1A1A")
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = None if not line else max(15, 13 * (len(line) // 95 + 1))
        r += 1


def _write_vocabulary(ws):
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 60
    _put(ws, 1, 1, "Data Type", bold=True, fill=FILL_HEADER)
    _put(ws, 1, 2, "Usage", bold=True, fill=FILL_HEADER)
    _put(ws, 1, 3, "Notes", bold=True, fill=FILL_HEADER)
    for i, t in enumerate(TYPES, start=2):
        _put(ws, i, 1, t)
    for i, u in enumerate(USAGES, start=2):
        _put(ws, i, 2, u if u else "(leave blank for optional)")
    notes = [
        "Do not rename or reorder these columns: the dropdowns point at them.",
        "A length in brackets becomes maxLength on a string.",
        "object means a nested group. Array means a repeating group.",
        "Mandatory lands in the required list of the generated schema.",
        "ConditionalMandatory is recorded as x-conditional, not as required.",
        "A blank Usage cell means optional.",
    ]
    for i, n in enumerate(notes, start=2):
        _put(ws, i, 3, n, wrap=True)


def _write_operation(wb, title, sor_endpoints, request_rows, response_rows,
                     banner=None):
    ws = _sheet(wb, title)
    first_sor = _write_header(ws, sor_endpoints)
    n_sor = len(sor_endpoints)
    values = {"SOR API Endpoint:": ", ".join(sor_endpoints)}
    values.update(banner or {})
    r = _write_banner(ws, values)
    r = _write_rows(ws, r, "Request Body", request_rows, first_sor, n_sor)
    r += 1
    r = _write_rows(ws, r, "Response Body", response_rows, first_sor, n_sor)
    _validations(ws, r, n_sor)
    return ws


def write_template(path):
    """Write the template workbook to ``path``."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _write_how_to(_sheet(wb, "How to use"))
    _write_vocabulary(_sheet(wb, "Vocabulary"))
    # Each example needs its own path and service domain, because the sheets
    # of one workbook are merged into one specification and two operations
    # cannot claim the same path and method.
    _write_operation(
        wb, "Example_SingleSor", [SINGLE_SOR], SINGLE_REQUEST, SINGLE_RESPONSE,
        banner={
            "Use Case:": "Retrieve the balances held against a deposit account.",
            "Service Domain:": "Deposit Account",
            "BQ:": "BQ: Account Balance",
            "Equivalent BIAN API Endpoint:":
                "/DepositAccount/AccountBalance/Retrieve",
            "Proposed Business API Endpoint:":
                "/deposit-account/account-balance/retrieve",
        })
    _write_operation(
        wb, "Example_MultiSor", MULTI_SOR, MULTI_REQUEST, MULTI_RESPONSE,
        banner={
            "Use Case:": "Retrieve account details, in one of three variants "
                         "chosen by the caller.",
            "Service Domain:": "Deposit Account",
            "BQ:": "BQ: Account",
            "Equivalent BIAN API Endpoint:": "/DepositAccount/Account/Retrieve",
            "Proposed Business API Endpoint:":
                "/deposit-account/account/retrieve",
            "SOR API Endpoint:": ", ".join(MULTI_SOR),
            "Method:": "POST",
        })
    _write_operation(
        wb, "Operation_Template", ["/v1/replace/me"],
        BLANK_REQUEST, BLANK_RESPONSE,
        banner={
            "Use Case:": "Replace this with what your endpoint is for.",
            "Service Domain:": "Deposit Account",
            "BQ:": "BQ: Replace Me",
            "Equivalent BIAN API Endpoint:":
                "/ServiceDomain/BehaviourQualifier/Retrieve",
            "Proposed Business API Endpoint:":
                "/replace-me/resource/retrieve",
            "SOR API Endpoint:": "/v1/replace/me",
            "Method:": "POST",
        })
    wb.save(path)
    return path


def main(argv=None):
    import sys
    args = list(argv if argv is not None else sys.argv[1:])
    target = args[0] if args else "SOR_mapping_template.xlsx"
    write_template(target)
    print("Wrote %s" % target)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


__all__ = ["HEADER", "BANNER", "TYPES", "USAGES", "write_template"]
