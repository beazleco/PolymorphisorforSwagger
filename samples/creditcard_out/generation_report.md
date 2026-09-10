# SOR Polymorphizer generation report

Workbook: `creditcard_v2.0.0.xlsx`

Input format: **Level format**

Generated 11 of 12 endpoints. **1 failed.**

Output: **one specification** for the whole workbook.

## Failed endpoints

No specification was written for these sheets. Each one needs a correction in the workbook.

### Card Issued Device_Retrieve

- **S001 at Card Issued Device_Retrieve!A7**
  - Found: this sheet has no Request Body section. The only request section is 'Request Header' at row 7, and the rows beneath it carry the request schema, so the label is wrong
  - Expected: a section labelled Request Body holding the request schema
  - Fix: change 'Request Header' in cell A7 to Request Body. If the sheet genuinely describes request headers, add a separate Request Body section for the schema.

## Endpoints generated

| Sheet | Pattern | SOR endpoints | Method | Path |
|---|---|---|---|---|
| Available Funds_Retrieve | P1 | 1 | POST | `/credit-card/financial-status-tracking/retrieve` |
| Transaction Adjustment_Initiate | P1 | 1 | POST | `/credit-card/card-transaction/adjustment/initiate` |
| Recency Check_Retrieve | P1 | 1 | POST | `/credit-card/mobile-recency/retrieve` |
| Account Statement_Retrieve | P1 | 1 | POST | `/credit-card/billing/retrieve` |
| Card Transaction_Retrieve | P2 | 3 | POST | `/CreditCard/CardTransaction/Retrieve` |
| Card Details_Retrieve | P2 | 5 | POST | `/CreditCard/Account/Retrieve` |
| Card issuedDevice_Initiate | P1 | 1 | POST | `/CreditCard/IssuedDevice/Initiate` |
| Card issuedDevice_Update | P1 | 1 | PUT | `/CreditCard/IssuedDevice/Update` |
| Card Cancel | P1 | 1 | PUT | `/credit-card/issued-device/control` |
| Customer Addl Data | P1 | 1 | POST | `/credit-card/customer/additional-data/retrieve` |
| Card RSAEncrypted | P1 | 1 | POST | `/credit-card/issued-device/card-rsa-encrypted/retrieve` |

### Patterns

- **P1**, one SOR endpoint: a single schema per message, pruned to what SOR supplies.
- **P2**, several SOR endpoints with a `requestVariant` enumeration: base plus `allOf` variants, wrapped in `oneOf` with a discriminator on the request.
- **P3**, several SOR endpoints with no tag: `oneOf` without a discriminator.

## The merged specification

| | |
|---|---|
| Operations | 11 |
| Schemas | 123 |
| Groups hoisted into components | 50 |
| Of those, specialised across operations | 13 |

### Groups specialised across operations

These groups are used by more than one operation and do not agree on what they publish. Where the shapes have attributes in common, the common part is published as `<Class>__Base` and every shape inherits it through `allOf`, carrying only its own delta. Where they have nothing in common, an `allOf` against an empty base would add nothing, so each shape is published whole and the Shared base column reads none. That is not necessarily a defect: two groups can share a label and genuinely describe different things. A `P001` warning against the group is the case that is worth fixing, because there the shapes do declare the same attribute and differ only in how, so aligning the workbook produces a base. The `P002` warnings name the exact differences.

| Group | Shared base | Derived schemas |
|---|---|---|
| `Account` | `Account__Base` | 4 |
| `AccountDetialsItem` | none | 5 |
| `AccountIdentifier` | none | 2 |
| `AccountTransactionItem` | `AccountTransactionItem__Base` | 3 |
| `CardPaymentTransactionItem` | `CardPaymentTransactionItem__Base` | 3 |
| `CardTransactionAmount` | `CardTransactionAmount__Base` | 2 |
| `CardTransactionAuthorizationRecord` | `CardTransactionAuthorizationRecord__Base` | 2 |
| `CardTransactionMerchantReference` | none | 3 |
| `DelinquentAccountDetailsItem` | none | 2 |
| `IssuedDeviceIdentifier` | none | 6 |
| `IssuedDeviceStatus` | `IssuedDeviceStatus__Base` | 2 |
| `PointOfService` | `PointOfService__Base` | 2 |
| `Status` | none | 2 |

### Findings from the merge

| Code | Found | Fix |
|---|---|---|
| N001 | 'AccountIdentifier' and 'Accountidentifier' differ only by case, and in one specification they become separate schemas one keystroke apart | settle on one spelling, 'AccountIdentifier', and correct the others in the workbook. The tool does not rename them for you, because that would change the published field names without your seeing it. |
| P002 | AccountIdentifier.accountIdentification is declared 2 different ways across cardDetailsRetrieve, cardTransactionRetrieve: string (11) mandatory against string optional | make the Data Type and Usage of AccountIdentifier.accountIdentification the same on every sheet. Doing so lets the tool publish one shared schema instead of 2. |
| P002 | CardTransactionAmount.amountValue is declared 2 different ways across cardTransactionRetrieve, transactionAdjustmentInitiate: number optional against string mandatory | make the Data Type and Usage of CardTransactionAmount.amountValue the same on every sheet. Doing so lets the tool publish one shared schema instead of 2. |
| P002 | IssuedDeviceIdentifier.issuedDeviceIdentification is declared 2 different ways across accountStatementRetrieve, availableFundsRetrieve, cardTransactionRetrieve, recencyCheckRetrieve, transactionAdjustmentInitiate: string (19) optional against string (19) mandatory | make the Data Type and Usage of IssuedDeviceIdentifier.issuedDeviceIdentification the same on every sheet. Doing so lets the tool publish one shared schema instead of 6. |
| P002 | IssuedDeviceIdentifier.issueddeviceIdentification is declared 3 different ways across cardCancel, cardDetailsRetrieve, cardIssuedDeviceInitiate, cardIssuedDeviceUpdate, cardRSAEncrypted, customerAddlData: string (19) mandatory against string optional against string mandatory | make the Data Type and Usage of IssuedDeviceIdentifier.issueddeviceIdentification the same on every sheet. Doing so lets the tool publish one shared schema instead of 6. |
| P002 | Status.statusEffectiveDate is declared 2 different ways across cardCancel, cardIssuedDeviceInitiate, cardIssuedDeviceUpdate: string format date optional against string format date-time optional | make the Data Type and Usage of Status.statusEffectiveDate the same on every sheet. Doing so lets the tool publish one shared schema instead of 2. |
| P001 | 'AccountDetialsItem' is used in 5 shapes which all declare Account, but not identically, so no shared base could be published | align the Data Type and Usage of Account across every sheet. The P002 findings name the exact differences. Doing so replaces 5 schemas with a base and 5 small derived schemas. |
| P001 | 'AccountIdentifier' is used in 2 shapes which all declare accountIdentification, but not identically, so no shared base could be published | align the Data Type and Usage of accountIdentification across every sheet. The P002 findings name the exact differences. Doing so replaces 2 schemas with a base and 2 small derived schemas. |
| P001 | 'Status' is used in 2 shapes which all declare statusEffectiveDate, but not identically, so no shared base could be published | align the Data Type and Usage of statusEffectiveDate across every sheet. The P002 findings name the exact differences. Doing so replaces 2 schemas with a base and 2 small derived schemas. |
| N002 | the service domain is spelled 3 ways across the workbook: 'Credit Card', 'Credit card', 'CreditCard' | settle on 'Credit Card' in the Service Domain cell of every sheet. The generated tag uses that spelling in the meantime. |

## Warnings

These did not stop generation. Each one is worth an analyst's attention.

| Code | Where | Found | Fix |
|---|---|---|---|
| T002 | `Available Funds_Retrieve!G10` | 'issuedDeviceIdentificationType' has no Data Type and no nested rows, so it was treated as a string | set the Data Type of 'issuedDeviceIdentificationType', for example String (35) or Number(12) |
| T002 | `Available Funds_Retrieve!G19` | 'partyIdentificationType' has no Data Type and no nested rows, so it was treated as a string | set the Data Type of 'partyIdentificationType', for example String (35) or Number(12) |
| T002 | `Available Funds_Retrieve!G27` | 'primaryCardTag' has no Data Type and no nested rows, so it was treated as a string | set the Data Type of 'primaryCardTag', for example String (35) or Number(12) |
| G001 | `Available Funds_Retrieve` | the sheet name implies GET but the Request Body section declares a schema, and GET cannot carry a body, so the operation was emitted as POST | confirm POST for this operation, or move the request attributes into a Request Parameter section so they become query parameters |
| T002 | `Transaction Adjustment_Initiate!H50` | 'authenticationMethodType' has no Data Type and no nested rows, so it was treated as a string | set the Data Type of 'authenticationMethodType', for example String (35) or Number(12) |
| G001 | `Recency Check_Retrieve` | the sheet name implies GET but the Request Body section declares a schema, and GET cannot carry a body, so the operation was emitted as POST | confirm POST for this operation, or move the request attributes into a Request Parameter section so they become query parameters |
| G001 | `Account Statement_Retrieve` | the sheet name implies GET but the Request Body section declares a schema, and GET cannot carry a body, so the operation was emitted as POST | confirm POST for this operation, or move the request attributes into a Request Parameter section so they become query parameters |
| V003 | `Card Transaction_Retrieve!L1` | variant 02 'Authorization Transactions' was paired with SOR column 'SOR v1/account/search/auth' by position, but the two names have nothing in common | reorder the SOR columns to match the order of the requestVariant values, or reword one of them so the pairing is obvious |
| V003 | `Card Transaction_Retrieve!M1` | variant 03 'Installment Transactions' was paired with SOR column 'SOR /v1/account/search/mps' by position, but the two names have nothing in common | reorder the SOR columns to match the order of the requestVariant values, or reword one of them so the pairing is obvious |
| G001 | `Card Transaction_Retrieve` | the sheet name implies GET but the Request Body section declares a schema, and GET cannot carry a body, so the operation was emitted as POST | confirm POST for this operation, or move the request attributes into a Request Parameter section so they become query parameters |
| A006 | `Card Details_Retrieve!I53` | 'AvailabilityType' is declared object but has no nested rows beneath it, so there is no shape to publish and the element was removed from the interface | add the attributes of 'AvailabilityType' one Level deeper, or change its Data Type to a scalar |
| A006 | `Card Details_Retrieve!I59` | 'AddressLine' is declared object but has no nested rows beneath it, so there is no shape to publish and the element was removed from the interface | add the attributes of 'AddressLine' one Level deeper, or change its Data Type to a scalar |
| A006 | `Card Details_Retrieve!I71` | 'Country' is declared object but has no nested rows beneath it, so there is no shape to publish and the element was removed from the interface | add the attributes of 'Country' one Level deeper, or change its Data Type to a scalar |
| V003 | `Card Details_Retrieve!L1` | variant 02 'Accunt Deliquency' was paired with SOR column 'SOR /v1/account/delinquency' by position, but the two names have nothing in common | reorder the SOR columns to match the order of the requestVariant values, or reword one of them so the pairing is obvious |
| G001 | `Card Details_Retrieve` | the sheet name implies GET but the Request Body section declares a schema, and GET cannot carry a body, so the operation was emitted as POST | confirm POST for this operation, or move the request attributes into a Request Parameter section so they become query parameters |
| T002 | `Card Issued Device_Retrieve!H28` | 'partyIdentificationMasked' has no Data Type and no nested rows, so it was treated as a string | set the Data Type of 'partyIdentificationMasked', for example String (35) or Number(12) |
| V001 | `Card Issued Device_Retrieve!I1` | this sheet has 3 SOR columns but the request declares no requestVariant values | add requestVariant to the request and list one numbered variant per SOR column, or confirm that the variant is chosen by which identifier the caller supplies |
| G001 | `Customer Addl Data` | the sheet name implies GET but the Request Body section declares a schema, and GET cannot carry a body, so the operation was emitted as POST | confirm POST for this operation, or move the request attributes into a Request Parameter section so they become query parameters |
| E001 | `Card RSAEncrypted!H17` | the Response Body section declares 3 attributes and not one of them names an SOR field, so no payload was published for it | fill in the SOR column for the attributes this message carries. If it carries none, delete the rows and leave the Response Body label in place. |
| G001 | `Card RSAEncrypted` | the sheet name implies GET but the Request Body section declares a schema, and GET cannot carry a body, so the operation was emitted as POST | confirm POST for this operation, or move the request attributes into a Request Parameter section so they become query parameters |

## Content ignored by design

Only Request Body and Response Body are read. Header and parameter sections are skipped, and anything below the mapping grid, such as a pasted sample payload, is not read at all.

| Sheet | Sections skipped | Rows below the grid |
|---|---|---|
| Available Funds_Retrieve | none | 124 |
| Card Transaction_Retrieve | Request Header, ResponseHeader | 0 |
| Card Details_Retrieve | Request Header, ResponseHeader | 0 |
| Card issuedDevice_Initiate | Request Header | 0 |
| Card issuedDevice_Update | Request Header | 0 |
| Card Cancel | Request Header | 0 |
| Customer Addl Data | Request Header | 0 |
| Card RSAEncrypted | Request Header | 0 |

## Support sheets

No Level columns, so these are not operation sheets: `SOR Payload`.
