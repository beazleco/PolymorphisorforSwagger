# SOR Polymorphizer generation report

Workbook: `issueddevice_fieldmapping_v1.0.7.xlsx`

Input format: **field mapping document**

Generated 22 of 22 endpoints. **0 failed.**

Output: **one specification** for the whole workbook.

## Endpoints generated

| Sheet | Pattern | SOR endpoints | Method | Path |
|---|---|---|---|---|
| Token_Search | P1 | 1 | POST | `/issued-device-administration/token-assignment/search` |
| Token_Get Credentials | P1 | 1 | POST | `/issued-device-administration/token-assignment/credentials/retrieve` |
| Token_Details | P1 | 1 | POST | `/issued-device-administration/token-assignment/token-details/retrieve` |
| Token_Resume | P1 | 1 | POST | `/issued-device-administration/token-assignment/resume` |
| Token_Delete | P1 | 1 | POST | `/issued-device-administration/token-assignment/delete` |
| Token_Suspend | P1 | 1 | POST | `/issued-device-administration/token-assignment/suspend` |
| Token_Get Operation | P1 | 1 | POST | `/issued-device-administration/token-assignment/get-operation` |
| Token_Get All Operations | P1 | 1 | POST | `/issued-device-administration/token-assignment/get-all-operations` |
| CiamBlackList_Delete | P1 | 1 | POST | `/issued-device-administration/device-assignment/black-list/delete` |
| CiamBlackList_Add | P1 | 1 | POST | `/issued-device-administration/device-assignment/black-list/add` |
| Card Renew | P1 | 1 | POST | `/issued-device-administration/card-device-assignment/renew` |
| Card Register | P1 | 1 | POST | `/issued-device-administration/card-device-assignment/register-card` |
| Card Create | P1 | 1 | POST | `/issued-device-administration/card-device-assignment/create` |
| Card Suspend | P1 | 1 | POST | `/issued-device-administration/card-device-assignment/suspend` |
| Card Replace | P1 | 1 | POST | `/issued-device-administration/card-device-assignment/replace` |
| Card Resume | P1 | 1 | POST | `/issued-device-administration/card-device-assignment/resume` |
| Get Card Credentials | P1 | 1 | POST | `/issued-device-administration/card-device-assignment/get-card-credentials` |
| Verify Card Credentials | P1 | 1 | POST | `/issued-device-administration/card-device-assignment/verify-card-credentials` |
| Card Delete | P1 | 1 | POST | `/issued-device-administration/card-device-assignment/delete` |
| Get Card Details | P1 | 1 | POST | `/issued-device-administration/card-device-assignment/get-card-details` |
| Get Operation | P1 | 1 | POST | `/issued-device-administration/card-device-assignment/get-operation` |
| Get All Operations | P1 | 1 | POST | `/issued-device-administration/card-device-assignment/get-all-operations` |

### Patterns

- **P1**, one SOR endpoint: a single schema per message, pruned to what SOR supplies.
- **P2**, several SOR endpoints with a `requestVariant` enumeration: base plus `allOf` variants, wrapped in `oneOf` with a discriminator on the request.
- **P3**, several SOR endpoints with no tag: `oneOf` without a discriminator.

## The merged specification

| | |
|---|---|
| Operations | 22 |
| Schemas | 99 |
| Groups hoisted into components | 35 |
| Of those, specialised across operations | 13 |

### Groups specialised across operations

These groups are used by more than one operation and do not agree on what they publish. Where the shapes have attributes in common, the common part is published as `<Class>__Base` and every shape inherits it through `allOf`, carrying only its own delta. Where they have nothing in common, an `allOf` against an empty base would add nothing, so each shape is published whole and the Shared base column reads none. That is not necessarily a defect: two groups can share a label and genuinely describe different things. A `P001` warning against the group is the case that is worth fixing, because there the shapes do declare the same attribute and differ only in how, so aligning the workbook produces a base. The `P002` warnings name the exact differences.

| Group | Shared base | Derived schemas |
|---|---|---|
| `AccountIdentifier` | `AccountIdentifier__Base` | 2 |
| `AccountListItem` | `AccountListItem__Base` | 2 |
| `CardIdentifier` | none | 3 |
| `CardServicingSummary` | `CardServicingSummary__Base` | 2 |
| `DeviceIdentifier` | `DeviceIdentifier__Base` | 2 |
| `IssuedDevice` | none | 3 |
| `IssuedDeviceIdentifier` | none | 2 |
| `NewCardIdentifier` | none | 2 |
| `OffSetPagination` | `OffSetPagination__Base` | 2 |
| `PartyIdentifier` | none | 3 |
| `TokenIdentifier` | none | 2 |
| `TokenServicingEventSummaryItem` | `TokenServicingEventSummaryItem__Base` | 2 |
| `TokenServicingSummary` | none | 2 |

### Findings from the merge

| Code | Found | Fix |
|---|---|---|
| P002 | AccountIdentifier.accountIdentification is declared 2 different ways across cardCreate, getCardDetails: string (24) mandatory against string optional | make the Data Type and Usage of AccountIdentifier.accountIdentification the same on every sheet. Doing so lets the tool publish one shared schema instead of 2. |
| P002 | AccountListItem.accountCurrencyCode is declared 2 different ways across cardCreate, getCardDetails: string (3) mandatory against string optional | make the Data Type and Usage of AccountListItem.accountCurrencyCode the same on every sheet. Doing so lets the tool publish one shared schema instead of 2. |
| P002 | AccountListItem.accountDefaultFlag is declared 2 different ways across cardCreate, getCardDetails: boolean optional against boolean mandatory | make the Data Type and Usage of AccountListItem.accountDefaultFlag the same on every sheet. Doing so lets the tool publish one shared schema instead of 2. |
| P002 | CardIdentifier.issuedDeviceIdentification is declared 3 different ways across cardCreate, cardDelete, cardRegister, cardReplace, cardResume, cardSuspend, getAllOperations, getCardCredentials, getCardDetails, getOperation, tokenDetails, verifyCardCredentials: string (48) optional against string (48) mandatory against string optional | make the Data Type and Usage of CardIdentifier.issuedDeviceIdentification the same on every sheet. Doing so lets the tool publish one shared schema instead of 3. |
| P002 | DeviceIdentifier.deviceIdentification is declared 2 different ways across ciamBlackListAdd, ciamBlackListDelete, tokenDetails, tokenSearch: string optional against string mandatory | make the Data Type and Usage of DeviceIdentifier.deviceIdentification the same on every sheet. Doing so lets the tool publish one shared schema instead of 2. |
| P002 | IssuedDeviceIdentifier.issuedDeviceIdentification is declared 2 different ways across cardRenew, getCardDetails: string (48) mandatory against string optional | make the Data Type and Usage of IssuedDeviceIdentifier.issuedDeviceIdentification the same on every sheet. Doing so lets the tool publish one shared schema instead of 2. |
| P002 | NewCardIdentifier.issuedDeviceIdentification is declared 2 different ways across cardReplace: string (48) optional against string optional | make the Data Type and Usage of NewCardIdentifier.issuedDeviceIdentification the same on every sheet. Doing so lets the tool publish one shared schema instead of 2. |
| P002 | OffSetPagination.limit is declared 2 different ways across getAllOperations, tokenGetAllOperations: string (50) optional against integer optional | make the Data Type and Usage of OffSetPagination.limit the same on every sheet. Doing so lets the tool publish one shared schema instead of 2. |
| P002 | PartyIdentifier.partyIdentification is declared 3 different ways across cardCreate, cardDelete, cardRegister, cardRenew, cardReplace, cardResume, cardSuspend, getAllOperations, getCardCredentials, getCardDetails, getOperation, tokenDelete, tokenDetails, tokenGetAllOperations, tokenGetCredentials, tokenGetOperation, tokenResume, tokenSearch, tokenSuspend, verifyCardCredentials: string (10) mandatory against string (64) mandatory against string optional | make the Data Type and Usage of PartyIdentifier.partyIdentification the same on every sheet. Doing so lets the tool publish one shared schema instead of 3. |
| P002 | TokenIdentifier.tokenIdentification is declared 2 different ways across tokenDelete, tokenDetails, tokenGetAllOperations, tokenGetCredentials, tokenGetOperation, tokenResume, tokenSearch, tokenSuspend: string (64) mandatory against string optional | make the Data Type and Usage of TokenIdentifier.tokenIdentification the same on every sheet. Doing so lets the tool publish one shared schema instead of 2. |
| P001 | 'CardIdentifier' is used in 3 shapes which all declare issuedDeviceIdentification, but not identically, so no shared base could be published | align the Data Type and Usage of issuedDeviceIdentification across every sheet. The P002 findings name the exact differences. Doing so replaces 3 schemas with a base and 3 small derived schemas. |
| P001 | 'IssuedDevice' is used in 3 shapes which all declare IssuedDeviceIdentifier, but not identically, so no shared base could be published | align the Data Type and Usage of IssuedDeviceIdentifier across every sheet. The P002 findings name the exact differences. Doing so replaces 3 schemas with a base and 3 small derived schemas. |
| P001 | 'IssuedDeviceIdentifier' is used in 2 shapes which all declare issuedDeviceIdentification, but not identically, so no shared base could be published | align the Data Type and Usage of issuedDeviceIdentification across every sheet. The P002 findings name the exact differences. Doing so replaces 2 schemas with a base and 2 small derived schemas. |
| P001 | 'NewCardIdentifier' is used in 2 shapes which all declare issuedDeviceIdentification, but not identically, so no shared base could be published | align the Data Type and Usage of issuedDeviceIdentification across every sheet. The P002 findings name the exact differences. Doing so replaces 2 schemas with a base and 2 small derived schemas. |
| P001 | 'PartyIdentifier' is used in 3 shapes which all declare partyIdentification, but not identically, so no shared base could be published | align the Data Type and Usage of partyIdentification across every sheet. The P002 findings name the exact differences. Doing so replaces 3 schemas with a base and 3 small derived schemas. |
| P001 | 'TokenIdentifier' is used in 2 shapes which all declare tokenIdentification, but not identically, so no shared base could be published | align the Data Type and Usage of tokenIdentification across every sheet. The P002 findings name the exact differences. Doing so replaces 2 schemas with a base and 2 small derived schemas. |
| P001 | 'TokenServicingSummary' is used in 2 shapes which all declare TokenServicingEventSummary, but not identically, so no shared base could be published | align the Data Type and Usage of TokenServicingEventSummary across every sheet. The P002 findings name the exact differences. Doing so replaces 2 schemas with a base and 2 small derived schemas. |

## Warnings

These did not stop generation. Each one is worth an analyst's attention.

| Code | Where | Found | Fix |
|---|---|---|---|
| S002 | `Token_Get Credentials!A29` | a second request section labelled 'Body' | merge it into the first one, or move it to its own sheet |
| A006 | `Token_Details!D48` | 'deviceBindingList' is declared array but no dotted path beneath it declares a member, so there is no shape to publish and the element was removed from the interface | add the members of 'deviceBindingList' as dotted paths beneath it, for example deviceBindingList.<member>, or change its Schema to a scalar type |
| A006 | `Token_Get Operation!D63` | 'supported' is declared array but no dotted path beneath it declares a member, so there is no shape to publish and the element was removed from the interface | add the members of 'supported' as dotted paths beneath it, for example supported.<member>, or change its Schema to a scalar type |
| A006 | `Token_Get All Operations!D72` | 'supported' is declared array but no dotted path beneath it declares a member, so there is no shape to publish and the element was removed from the interface | add the members of 'supported' as dotted paths beneath it, for example supported.<member>, or change its Schema to a scalar type |
| F001 | `CiamBlackList_Delete!A21` | 'DeviceBlackListDeleteRequest' sits under no recognised section, so it was read into the request | add a Request Body or Response Body row in the Parameter Type column above this row |
| E001 | `CiamBlackList_Delete!E32` | the Response Body section declares 2 attributes and not one of them names an SOR field, so no payload was published for it | fill in the SOR column for the attributes this message carries. If it carries none, delete the rows and leave the Response Body label in place. |
| E001 | `CiamBlackList_Add!E32` | the Response Body section declares 2 attributes and not one of them names an SOR field, so no payload was published for it | fill in the SOR column for the attributes this message carries. If it carries none, delete the rows and leave the Response Body label in place. |
| F001 | `Card Replace!A21` | 'CardReplaceRequest' sits under no recognised section, so it was read into the request | add a Request Body or Response Body row in the Parameter Type column above this row |
| S002 | `Card Replace!A51` | a second request section labelled 'Body' | merge it into the first one, or move it to its own sheet |

## Content ignored by design

Only Request Body and Response Body are read. Header and parameter sections are skipped, and anything below the mapping grid, such as a pasted sample payload, is not read at all.

| Sheet | Sections skipped | Rows below the grid |
|---|---|---|
| Token_Search | Request Parameter, Response Header | 10 |
| Token_Get Credentials | Request Parameter, Response Header | 10 |
| Token_Details | Request Parameter, Response Header | 10 |
| Token_Resume | Request Parameter, Response Header | 10 |
| Token_Delete | Request Parameter, Response Header | 10 |
| Token_Suspend | Request Parameter, Response Header | 10 |
| Token_Get Operation | Request Parameter, Response Header | 10 |
| Token_Get All Operations | Request Parameter, Response Header | 10 |
| CiamBlackList_Delete | Request Parameter, Request Parameter, Response Header | 10 |
| CiamBlackList_Add | Request Parameter, Response Header | 10 |
| Card Renew | Request Parameter, Response Header | 10 |
| Card Register | Request Parameter, Response Header | 10 |
| Card Create | Request Parameter, Response Header | 10 |
| Card Suspend | Request Parameter, Response Header | 10 |
| Card Replace | Request Parameter, Request Header, Response Header | 10 |
| Card Resume | Request Parameter, Response Header | 10 |
| Get Card Credentials | Request Parameter, Response Header | 10 |
| Verify Card Credentials | Request Parameter, Response Header | 10 |
| Card Delete | Request Parameter, Response Header | 10 |
| Get Card Details | Request Parameter, Response Header | 10 |
| Get Operation | Request Parameter, Response Header | 10 |
| Get All Operations | Request Parameter, Response Header | 10 |
