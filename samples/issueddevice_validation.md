# Workbook validation report

Workbook: `issueddevice_fieldmapping_v1.0.7.xlsx`

Level: **lenient**

Input format: **Field mapping document**

SOR specification: **none supplied**, so every SOR field name in the workbook was taken on trust.

22 operation sheets, **22 would generate**, **0 would fail**. 0 support sheets skipped.

0 errors, 9 warnings.

## Endpoints that would generate

| Sheet | Pattern | Warnings |
|---|---|---|
| Token_Search | P1 | 0 |
| Token_Get Credentials | P1 | 1 |
| Token_Details | P1 | 1 |
| Token_Resume | P1 | 0 |
| Token_Delete | P1 | 0 |
| Token_Suspend | P1 | 0 |
| Token_Get Operation | P1 | 1 |
| Token_Get All Operations | P1 | 1 |
| CiamBlackList_Delete | P1 | 2 |
| CiamBlackList_Add | P1 | 1 |
| Card Renew | P1 | 0 |
| Card Register | P1 | 0 |
| Card Create | P1 | 0 |
| Card Suspend | P1 | 0 |
| Card Replace | P1 | 2 |
| Card Resume | P1 | 0 |
| Get Card Credentials | P1 | 0 |
| Verify Card Credentials | P1 | 0 |
| Card Delete | P1 | 0 |
| Get Card Details | P1 | 0 |
| Get Operation | P1 | 0 |
| Get All Operations | P1 | 0 |

## Warnings

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

## Findings by code

| Code | Count |
|---|---|
| A006 | 3 |
| E001 | 2 |
| F001 | 2 |
| S002 | 2 |
