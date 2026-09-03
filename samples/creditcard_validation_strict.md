# Workbook validation report

Workbook: `creditcard_v2.0.0.xlsx`

Level: **strict**

SOR specification: **none supplied**, so every SOR field name in the workbook was taken on trust.

12 operation sheets, **11 would generate**, **1 would fail**. 1 support sheet skipped.

1 error, 152 warnings.

## Endpoints that would fail

Each of these needs a correction in the workbook before a specification can be generated for it.

### Card Issued Device_Retrieve

- **S001 at Card Issued Device_Retrieve!A7**
  - Found: this sheet has no Request Body section. The only request section is 'Request Header' at row 7, and the rows beneath it carry the request schema, so the label is wrong
  - Expected: a section labelled Request Body holding the request schema
  - Fix: change 'Request Header' in cell A7 to Request Body. If the sheet genuinely describes request headers, add a separate Request Body section for the schema.

## Endpoints that would generate

| Sheet | Pattern | Warnings |
|---|---|---|
| Available Funds_Retrieve | P1 | 14 |
| Transaction Adjustment_Initiate | P1 | 23 |
| Recency Check_Retrieve | P1 | 4 |
| Account Statement_Retrieve | P1 | 6 |
| Card Transaction_Retrieve | P2 | 36 |
| Card Details_Retrieve | P2 | 21 |
| Card issuedDevice_Initiate | P1 | 12 |
| Card issuedDevice_Update | P1 | 13 |
| Card Cancel | P1 | 8 |
| Customer Addl Data | P1 | 3 |
| Card RSAEncrypted | P1 | 8 |

## Warnings

| Code | Where | Found | Fix |
|---|---|---|---|
| T002 | `Available Funds_Retrieve!G10` | 'issuedDeviceIdentificationType' has no Data Type and no nested rows, so it was treated as a string | set the Data Type of 'issuedDeviceIdentificationType', for example String (35) or Number(12) |
| T002 | `Available Funds_Retrieve!G19` | 'partyIdentificationType' has no Data Type and no nested rows, so it was treated as a string | set the Data Type of 'partyIdentificationType', for example String (35) or Number(12) |
| T002 | `Available Funds_Retrieve!G27` | 'primaryCardTag' has no Data Type and no nested rows, so it was treated as a string | set the Data Type of 'primaryCardTag', for example String (35) or Number(12) |
| R010 | `Available Funds_Retrieve!J25` | 'limitCurrencyCode' names the SOR field 'customerLimitCurrency', but no element called 'customerLimitCurrency' exists in the response of GET /v1/card/availableFunds, so the element was removed from the interface | correct the field name in cell J25, or clear the cell if the SOR genuinely does not supply 'limitCurrencyCode'. If the field does exist, check that the sheet names the right SOR endpoint. |
| R021 | `Available Funds_Retrieve` | 1 of 14 SOR field names on this sheet do not exist in the SOR endpoint, so 1 element was removed from the interface | the R010 findings name each one with its cell. Correct them in the workbook, or accept the removals if the SOR really has changed. |
| G001 | `Available Funds_Retrieve` | the sheet name implies GET but the Request Body section declares a schema, and GET cannot carry a body, so the operation was emitted as POST | confirm POST for this operation, or move the request attributes into a Request Parameter section so they become query parameters |
| B001 | `Available Funds_Retrieve!A2` | Equivalent BIAN API Endpoint is missing or empty | add a row in column A labelled 'Equivalent BIAN API Endpoint:' with the value in the next column |
| D001 | `Available Funds_Retrieve!I9` | 'IssuedDeviceIdentifier' has no description | describe what 'IssuedDeviceIdentifier' means to a consumer of the API |
| D001 | `Available Funds_Retrieve!I12` | 'issuedDeviceIdentificationMasked' has no description | describe what 'issuedDeviceIdentificationMasked' means to a consumer of the API |
| D001 | `Available Funds_Retrieve!I21` | 'partyIdentificationMasked' has no description | describe what 'partyIdentificationMasked' means to a consumer of the API |
| D001 | `Available Funds_Retrieve!I35` | 'CardLimit' has no description | describe what 'CardLimit' means to a consumer of the API |
| D001 | `Available Funds_Retrieve!I40` | 'amountType' has no description | describe what 'amountType' means to a consumer of the API |
| D001 | `Available Funds_Retrieve!I47` | 'LastPayment' has no description | describe what 'LastPayment' means to a consumer of the API |
| D001 | `Available Funds_Retrieve!I48` | 'paymentAmountType' has no description | describe what 'paymentAmountType' means to a consumer of the API |
| T002 | `Transaction Adjustment_Initiate!H50` | 'authenticationMethodType' has no Data Type and no nested rows, so it was treated as a string | set the Data Type of 'authenticationMethodType', for example String (35) or Number(12) |
| B001 | `Transaction Adjustment_Initiate!A2` | Equivalent BIAN API Endpoint is missing or empty | add a row in column A labelled 'Equivalent BIAN API Endpoint:' with the value in the next column |
| D001 | `Transaction Adjustment_Initiate!J19` | 'issuedDeviceIdentificationMasked' has no description | describe what 'issuedDeviceIdentificationMasked' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J25` | 'cardTransactionCode' has no description | describe what 'cardTransactionCode' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J28` | 'MajorPurchaseDetails' has no description | describe what 'MajorPurchaseDetails' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J29` | 'multiPaymentIdentification' has no description | describe what 'multiPaymentIdentification' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J30` | 'numberInstallments' has no description | describe what 'numberInstallments' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J31` | 'interestRate' has no description | describe what 'interestRate' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J32` | 'MultiPaymentStatus' has no description | describe what 'MultiPaymentStatus' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J33` | 'value' has no description | describe what 'value' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J34` | 'description' has no description | describe what 'description' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J40` | 'DebitCreditDirection' has no description | describe what 'DebitCreditDirection' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J53` | 'PointOfServiceContext' has no description | describe what 'PointOfServiceContext' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J54` | 'pANEntryMode' has no description | describe what 'pANEntryMode' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J55` | 'pOSConditionCode' has no description | describe what 'pOSConditionCode' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J57` | 'pointOfServiceLocationReference' has no description | describe what 'pointOfServiceLocationReference' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J59` | 'cardAcceptorName' has no description | describe what 'cardAcceptorName' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J60` | 'cardAcceptorClassification' has no description | describe what 'cardAcceptorClassification' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J73` | 'IssuedDeviceStatus' has no description | describe what 'IssuedDeviceStatus' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J76` | 'primaryCardTag' has no description | describe what 'primaryCardTag' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J77` | 'AccountStatus' has no description | describe what 'AccountStatus' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J78` | 'accountStatusCode' has no description | describe what 'accountStatusCode' means to a consumer of the API |
| D001 | `Transaction Adjustment_Initiate!J79` | 'accountStatusReason' has no description | describe what 'accountStatusReason' means to a consumer of the API |
| G001 | `Recency Check_Retrieve` | the sheet name implies GET but the Request Body section declares a schema, and GET cannot carry a body, so the operation was emitted as POST | confirm POST for this operation, or move the request attributes into a Request Parameter section so they become query parameters |
| B001 | `Recency Check_Retrieve!A2` | Equivalent BIAN API Endpoint is missing or empty | add a row in column A labelled 'Equivalent BIAN API Endpoint:' with the value in the next column |
| D001 | `Recency Check_Retrieve!H10` | 'IssuedDeviceIdentifier' has no description | describe what 'IssuedDeviceIdentifier' means to a consumer of the API |
| D001 | `Recency Check_Retrieve!H13` | 'issuedDeviceIdentificationMasked' has no description | describe what 'issuedDeviceIdentificationMasked' means to a consumer of the API |
| G001 | `Account Statement_Retrieve` | the sheet name implies GET but the Request Body section declares a schema, and GET cannot carry a body, so the operation was emitted as POST | confirm POST for this operation, or move the request attributes into a Request Parameter section so they become query parameters |
| B001 | `Account Statement_Retrieve!A2` | Equivalent BIAN API Endpoint is missing or empty | add a row in column A labelled 'Equivalent BIAN API Endpoint:' with the value in the next column |
| D001 | `Account Statement_Retrieve!J10` | 'Accountidentifier' has no description | describe what 'Accountidentifier' means to a consumer of the API |
| D001 | `Account Statement_Retrieve!J11` | 'accountIdentificationType' has no description | describe what 'accountIdentificationType' means to a consumer of the API |
| D001 | `Account Statement_Retrieve!J13` | 'IssuedDeviceIdentifier' has no description | describe what 'IssuedDeviceIdentifier' means to a consumer of the API |
| D001 | `Account Statement_Retrieve!J16` | 'issuedDeviceIdentificationMasked' has no description | describe what 'issuedDeviceIdentificationMasked' means to a consumer of the API |
| R011 | `Card Transaction_Retrieve!K33` | 'accountCurrencyCode' matched the SOR field 'currencyCode' in 2 different places in the response of GET /v1/account/search/trans: accountTransactions.currencyCode, accountTransactions.transactions.transactionAmount.currencyCode | write the fuller path in cell K33 so it is clear which one is meant. Matching is on the last segment, so the element was kept, but a reader cannot tell which SOR field feeds it. |
| R011 | `Card Transaction_Retrieve!L33` | 'accountCurrencyCode' matched the SOR field 'currencyCode' in 2 different places in the response of GET /v1/account/search/auth: acccountAuthorizations.authorizations.transactionAmount.currencyCode, acccountAuthorizations.currencyCode | write the fuller path in cell L33 so it is clear which one is meant. Matching is on the last segment, so the element was kept, but a reader cannot tell which SOR field feeds it. |
| R011 | `Card Transaction_Retrieve!M33` | 'accountCurrencyCode' matched the SOR field 'currencyCode' in 2 different places in the response of GET /v1/account/search/mps: accountMps.currencyCode, accountMps.mps.transactionAmount.currencyCode | write the fuller path in cell M33 so it is clear which one is meant. Matching is on the last segment, so the element was kept, but a reader cannot tell which SOR field feeds it. |
| R011 | `Card Transaction_Retrieve!K58` | 'amountCurrencyCode' matched the SOR field 'currencyCode' in 2 different places in the response of GET /v1/account/search/trans: accountTransactions.currencyCode, accountTransactions.transactions.transactionAmount.currencyCode | write the fuller path in cell K58 so it is clear which one is meant. Matching is on the last segment, so the element was kept, but a reader cannot tell which SOR field feeds it. |
| R011 | `Card Transaction_Retrieve!L58` | 'amountCurrencyCode' matched the SOR field 'currencyCode' in 2 different places in the response of GET /v1/account/search/auth: acccountAuthorizations.authorizations.transactionAmount.currencyCode, acccountAuthorizations.currencyCode | write the fuller path in cell L58 so it is clear which one is meant. Matching is on the last segment, so the element was kept, but a reader cannot tell which SOR field feeds it. |
| R011 | `Card Transaction_Retrieve!M58` | 'amountCurrencyCode' matched the SOR field 'currencyCode' in 2 different places in the response of GET /v1/account/search/mps: accountMps.currencyCode, accountMps.mps.transactionAmount.currencyCode | write the fuller path in cell M58 so it is clear which one is meant. Matching is on the last segment, so the element was kept, but a reader cannot tell which SOR field feeds it. |
| V003 | `Card Transaction_Retrieve!L1` | variant 02 'Authorization Transactions' was paired with SOR column 'SOR v1/account/search/auth' by position, but the two names have nothing in common | reorder the SOR columns to match the order of the requestVariant values, or reword one of them so the pairing is obvious |
| V003 | `Card Transaction_Retrieve!M1` | variant 03 'Installment Transactions' was paired with SOR column 'SOR /v1/account/search/mps' by position, but the two names have nothing in common | reorder the SOR columns to match the order of the requestVariant values, or reword one of them so the pairing is obvious |
| G001 | `Card Transaction_Retrieve` | the sheet name implies GET but the Request Body section declares a schema, and GET cannot carry a body, so the operation was emitted as POST | confirm POST for this operation, or move the request attributes into a Request Parameter section so they become query parameters |
| L003 | `Card Transaction_Retrieve!H1` | there is no Dummy Value column | add a Dummy Value header to the right of the Level columns |
| D001 | `Card Transaction_Retrieve!J10` | 'IssuedDeviceIdentifier' has no description | describe what 'IssuedDeviceIdentifier' means to a consumer of the API |
| D002 | `Card Transaction_Retrieve!A12` | 'issuedDeviceIdentification' is Mandatory but has no example value | put a realistic value in the Dummy Value column for 'issuedDeviceIdentification' |
| D001 | `Card Transaction_Retrieve!J13` | 'issuedDeviceIdentificationMasked' has no description | describe what 'issuedDeviceIdentificationMasked' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J16` | 'transactionCode' has no description | describe what 'transactionCode' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J17` | 'CardTransactionAmount' has no description | describe what 'CardTransactionAmount' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J18` | 'amountType' has no description | describe what 'amountType' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J30` | 'AccountIdentifier' has no description | describe what 'AccountIdentifier' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J31` | 'accountIdentificationType' has no description | describe what 'accountIdentificationType' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J39` | 'issuedDeviceIdentificationMasked' has no description | describe what 'issuedDeviceIdentificationMasked' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J42` | 'CardTransactionBillingStatementReference' has no description | describe what 'CardTransactionBillingStatementReference' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J45` | 'cardTransactionCode' has no description | describe what 'cardTransactionCode' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J48` | 'MajorPurchaseDetails' has no description | describe what 'MajorPurchaseDetails' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J49` | 'multiPaymentIdentification' has no description | describe what 'multiPaymentIdentification' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J50` | 'numberInstallments' has no description | describe what 'numberInstallments' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J51` | 'interestRate' has no description | describe what 'interestRate' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J52` | 'MultiPaymentStatus' has no description | describe what 'MultiPaymentStatus' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J53` | 'value' has no description | describe what 'value' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J54` | 'description' has no description | describe what 'description' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J55` | 'CardTransactionAmount' has no description | describe what 'CardTransactionAmount' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J57` | 'amountValue' has no description | describe what 'amountValue' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J60` | 'DebitCreditDirection' has no description | describe what 'DebitCreditDirection' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J73` | 'PointOfServiceContext' has no description | describe what 'PointOfServiceContext' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J74` | 'pANEntryMode' has no description | describe what 'pANEntryMode' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J75` | 'pOSConditionCode' has no description | describe what 'pOSConditionCode' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J77` | 'pointOfServiceLocationReference' has no description | describe what 'pointOfServiceLocationReference' means to a consumer of the API |
| D001 | `Card Transaction_Retrieve!J80` | 'cardAcceptorClassification' has no description | describe what 'cardAcceptorClassification' means to a consumer of the API |
| A006 | `Card Details_Retrieve!I53` | 'AvailabilityType' is declared object but has no nested rows beneath it | add the attributes of 'AvailabilityType' one Level deeper, or change its Data Type to a scalar |
| A006 | `Card Details_Retrieve!I59` | 'AddressLine' is declared object but has no nested rows beneath it | add the attributes of 'AddressLine' one Level deeper, or change its Data Type to a scalar |
| A006 | `Card Details_Retrieve!I71` | 'Country' is declared object but has no nested rows beneath it | add the attributes of 'Country' one Level deeper, or change its Data Type to a scalar |
| R010 | `Card Details_Retrieve!M29` | 'accountName' names the SOR field 'accounts.accountName', but no element called 'accountName' exists in the response of GET /v1/account/info, so the element was removed from the interface | correct the field name in cell M29, or clear the cell if the SOR genuinely does not supply 'accountName'. If the field does exist, check that the sheet names the right SOR endpoint. |
| R011 | `Card Details_Retrieve!M87` | 'productIdentification' matched the SOR field 'value' in 3 different places in the response of GET /v1/account/info: accounts.accountProduct.value, accounts.reclassification.value, accounts.repayment.value | write the fuller path in cell M87 so it is clear which one is meant. Matching is on the last segment, so the element was kept, but a reader cannot tell which SOR field feeds it. |
| R011 | `Card Details_Retrieve!M88` | 'productDescription' matched the SOR field 'description' in 3 different places in the response of GET /v1/account/info: accounts.accountProduct.description, accounts.reclassification.description, accounts.repayment.description | write the fuller path in cell M88 so it is clear which one is meant. Matching is on the last segment, so the element was kept, but a reader cannot tell which SOR field feeds it. |
| R011 | `Card Details_Retrieve!M90` | 'accountReclassificationValue' matched the SOR field 'value' in 3 different places in the response of GET /v1/account/info: accounts.accountProduct.value, accounts.reclassification.value, accounts.repayment.value | write the fuller path in cell M90 so it is clear which one is meant. Matching is on the last segment, so the element was kept, but a reader cannot tell which SOR field feeds it. |
| R011 | `Card Details_Retrieve!M91` | 'accountReclassificationDescription' matched the SOR field 'description' in 3 different places in the response of GET /v1/account/info: accounts.accountProduct.description, accounts.reclassification.description, accounts.repayment.description | write the fuller path in cell M91 so it is clear which one is meant. Matching is on the last segment, so the element was kept, but a reader cannot tell which SOR field feeds it. |
| R011 | `Card Details_Retrieve!M93` | 'repaymentValue' matched the SOR field 'value' in 3 different places in the response of GET /v1/account/info: accounts.accountProduct.value, accounts.reclassification.value, accounts.repayment.value | write the fuller path in cell M93 so it is clear which one is meant. Matching is on the last segment, so the element was kept, but a reader cannot tell which SOR field feeds it. |
| R011 | `Card Details_Retrieve!M94` | 'repaymentDescription' matched the SOR field 'description' in 3 different places in the response of GET /v1/account/info: accounts.accountProduct.description, accounts.reclassification.description, accounts.repayment.description | write the fuller path in cell M94 so it is clear which one is meant. Matching is on the last segment, so the element was kept, but a reader cannot tell which SOR field feeds it. |
| R021 | `Card Details_Retrieve` | 1 of 73 SOR field names on this sheet do not exist in the SOR endpoint, so 1 element was removed from the interface | the R010 findings name each one with its cell. Correct them in the workbook, or accept the removals if the SOR really has changed. |
| V003 | `Card Details_Retrieve!L1` | variant 02 'Accunt Deliquency' was paired with SOR column 'SOR /v1/account/delinquency' by position, but the two names have nothing in common | reorder the SOR columns to match the order of the requestVariant values, or reword one of them so the pairing is obvious |
| G001 | `Card Details_Retrieve` | the sheet name implies GET but the Request Body section declares a schema, and GET cannot carry a body, so the operation was emitted as POST | confirm POST for this operation, or move the request attributes into a Request Parameter section so they become query parameters |
| L003 | `Card Details_Retrieve!H1` | there is no Dummy Value column | add a Dummy Value header to the right of the Level columns |
| D002 | `Card Details_Retrieve!A12` | 'accountIdentification' is Mandatory but has no example value | put a realistic value in the Dummy Value column for 'accountIdentification' |
| D002 | `Card Details_Retrieve!A16` | 'issueddeviceIdentification' is Mandatory but has no example value | put a realistic value in the Dummy Value column for 'issueddeviceIdentification' |
| D001 | `Card Details_Retrieve!J23` | 'AccountDetials' has no description | describe what 'AccountDetials' means to a consumer of the API |
| D001 | `Card Details_Retrieve!J24` | 'Account' has no description | describe what 'Account' means to a consumer of the API |
| D001 | `Card Details_Retrieve!J31` | 'accountOpenDate' has no description | describe what 'accountOpenDate' means to a consumer of the API |
| D001 | `Card Details_Retrieve!J34` | 'accountDate' has no description | describe what 'accountDate' means to a consumer of the API |
| D001 | `Card Details_Retrieve!J83` | 'DeliveryOptions' has no description | describe what 'DeliveryOptions' means to a consumer of the API |
| L003 | `Card issuedDevice_Initiate!F1` | there is no Dummy Value column | add a Dummy Value header to the right of the Level columns |
| D001 | `Card issuedDevice_Initiate!I9` | 'IssuedDeviceIdentifier' has no description | describe what 'IssuedDeviceIdentifier' means to a consumer of the API |
| D001 | `Card issuedDevice_Initiate!I11` | 'issueddeviceIdentification' has no description | describe what 'issueddeviceIdentification' means to a consumer of the API |
| D002 | `Card issuedDevice_Initiate!A11` | 'issueddeviceIdentification' is Mandatory but has no example value | put a realistic value in the Dummy Value column for 'issueddeviceIdentification' |
| D001 | `Card issuedDevice_Initiate!I12` | 'issueddeviceIdentificationMasked' has no description | describe what 'issueddeviceIdentificationMasked' means to a consumer of the API |
| D001 | `Card issuedDevice_Initiate!I15` | 'IssuedDeviceIdentifier' has no description | describe what 'IssuedDeviceIdentifier' means to a consumer of the API |
| D001 | `Card issuedDevice_Initiate!I17` | 'issueddeviceIdentification' has no description | describe what 'issueddeviceIdentification' means to a consumer of the API |
| D001 | `Card issuedDevice_Initiate!I18` | 'issueddeviceIdentificationMasked' has no description | describe what 'issueddeviceIdentificationMasked' means to a consumer of the API |
| D001 | `Card issuedDevice_Initiate!I19` | 'Status' has no description | describe what 'Status' means to a consumer of the API |
| D001 | `Card issuedDevice_Initiate!I20` | 'statusCode' has no description | describe what 'statusCode' means to a consumer of the API |
| D001 | `Card issuedDevice_Initiate!I21` | 'statusDescription' has no description | describe what 'statusDescription' means to a consumer of the API |
| D001 | `Card issuedDevice_Initiate!I22` | 'statusEffectiveDate' has no description | describe what 'statusEffectiveDate' means to a consumer of the API |
| R005 | `Card issuedDevice_Update!H1` | the workbook states DELETE for '/v1/card/cancel' but the SOR specification defines only PUT, so PUT was used | correct the method in the SOR API Endpoint cell, or confirm that PUT is the operation intended |
| L003 | `Card issuedDevice_Update!F1` | there is no Dummy Value column | add a Dummy Value header to the right of the Level columns |
| D001 | `Card issuedDevice_Update!I9` | 'IssuedDeviceIdentifier' has no description | describe what 'IssuedDeviceIdentifier' means to a consumer of the API |
| D001 | `Card issuedDevice_Update!I11` | 'issueddeviceIdentification' has no description | describe what 'issueddeviceIdentification' means to a consumer of the API |
| D002 | `Card issuedDevice_Update!A11` | 'issueddeviceIdentification' is Mandatory but has no example value | put a realistic value in the Dummy Value column for 'issueddeviceIdentification' |
| D001 | `Card issuedDevice_Update!I12` | 'issueddeviceIdentificationMasked' has no description | describe what 'issueddeviceIdentificationMasked' means to a consumer of the API |
| D001 | `Card issuedDevice_Update!I15` | 'IssuedDeviceIdentifier' has no description | describe what 'IssuedDeviceIdentifier' means to a consumer of the API |
| D001 | `Card issuedDevice_Update!I17` | 'issueddeviceIdentification' has no description | describe what 'issueddeviceIdentification' means to a consumer of the API |
| D001 | `Card issuedDevice_Update!I18` | 'issueddeviceIdentificationMasked' has no description | describe what 'issueddeviceIdentificationMasked' means to a consumer of the API |
| D001 | `Card issuedDevice_Update!I19` | 'Status' has no description | describe what 'Status' means to a consumer of the API |
| D001 | `Card issuedDevice_Update!I20` | 'statusCode' has no description | describe what 'statusCode' means to a consumer of the API |
| D001 | `Card issuedDevice_Update!I21` | 'statusDescription' has no description | describe what 'statusDescription' means to a consumer of the API |
| D001 | `Card issuedDevice_Update!I22` | 'statusEffectiveDate' has no description | describe what 'statusEffectiveDate' means to a consumer of the API |
| T002 | `Card Issued Device_Retrieve!H28` | 'partyIdentificationMasked' has no Data Type and no nested rows, so it was treated as a string | set the Data Type of 'partyIdentificationMasked', for example String (35) or Number(12) |
| V001 | `Card Issued Device_Retrieve!I1` | this sheet has 3 SOR columns but the request declares no requestVariant values | add requestVariant to the request and list one numbered variant per SOR column, or confirm that the variant is chosen by which identifier the caller supplies |
| L003 | `Card Issued Device_Retrieve!G1` | there is no Dummy Value column | add a Dummy Value header to the right of the Level columns |
| D001 | `Card Issued Device_Retrieve!J28` | 'partyIdentificationMasked' has no description | describe what 'partyIdentificationMasked' means to a consumer of the API |
| D001 | `Card Cancel!I9` | 'IssuedDeviceIdentifier' has no description | describe what 'IssuedDeviceIdentifier' means to a consumer of the API |
| D001 | `Card Cancel!I11` | 'issueddeviceIdentification' has no description | describe what 'issueddeviceIdentification' means to a consumer of the API |
| D001 | `Card Cancel!I12` | 'issueddeviceIdentificationMasked' has no description | describe what 'issueddeviceIdentificationMasked' means to a consumer of the API |
| D002 | `Card Cancel!J14` | 'cancelNotes' is Mandatory but has no example value | put a realistic value in the Dummy Value column for 'cancelNotes' |
| D001 | `Card Cancel!I25` | 'IssuedDeviceIdentifier' has no description | describe what 'IssuedDeviceIdentifier' means to a consumer of the API |
| D001 | `Card Cancel!I29` | 'Status' has no description | describe what 'Status' means to a consumer of the API |
| D001 | `Card Cancel!I30` | 'statusCode' has no description | describe what 'statusCode' means to a consumer of the API |
| D001 | `Card Cancel!I31` | 'statusDescription' has no description | describe what 'statusDescription' means to a consumer of the API |
| G001 | `Customer Addl Data` | the sheet name implies GET but the Request Body section declares a schema, and GET cannot carry a body, so the operation was emitted as POST | confirm POST for this operation, or move the request attributes into a Request Parameter section so they become query parameters |
| D001 | `Customer Addl Data!I9` | 'PartyIdentifier' has no description | describe what 'PartyIdentifier' means to a consumer of the API |
| D001 | `Customer Addl Data!I13` | 'IssuedDeviceIdentifier' has no description | describe what 'IssuedDeviceIdentifier' means to a consumer of the API |
| R001 | `Card RSAEncrypted!H1` | the column header names '/v1/card/cancel' but the SOR API Endpoint banner names '/v1/card/{}/cardRSAEncrypted' for this column, and the banner is definitive, so '/v1/card/{}/cardRSAEncrypted' was used | correct the column header to 'SOR /v1/card/{}/cardRSAEncrypted', or correct the banner. Only the banner is read, so the interface is unaffected either way, but a reader of the workbook will be misled. |
| E001 | `Card RSAEncrypted!H17` | the Response Body section declares 3 attributes and not one of them names an SOR field, so no payload was published for it | fill in the SOR column for the attributes this message carries. If it carries none, delete the rows and leave the Response Body label in place. |
| G001 | `Card RSAEncrypted` | the sheet name implies GET but the Request Body section declares a schema, and GET cannot carry a body, so the operation was emitted as POST | confirm POST for this operation, or move the request attributes into a Request Parameter section so they become query parameters |
| L003 | `Card RSAEncrypted!F1` | there is no Dummy Value column | add a Dummy Value header to the right of the Level columns |
| D002 | `Card RSAEncrypted!A9` | 'encodedKey' is Mandatory but has no example value | put a realistic value in the Dummy Value column for 'encodedKey' |
| D001 | `Card RSAEncrypted!I12` | 'IssuedDeviceIdentifier' has no description | describe what 'IssuedDeviceIdentifier' means to a consumer of the API |
| D002 | `Card RSAEncrypted!A14` | 'issueddeviceIdentification' is Mandatory but has no example value | put a realistic value in the Dummy Value column for 'issueddeviceIdentification' |
| D001 | `Card RSAEncrypted!I15` | 'issueddeviceIdentificationMasked' has no description | describe what 'issueddeviceIdentificationMasked' means to a consumer of the API |

## Findings by code

| Code | Count |
|---|---|
| A006 | 3 |
| B001 | 4 |
| D001 | 96 |
| D002 | 8 |
| E001 | 1 |
| G001 | 7 |
| L003 | 6 |
| R001 | 1 |
| R005 | 1 |
| R010 | 2 |
| R011 | 12 |
| R021 | 2 |
| S001 | 1 |
| T002 | 5 |
| V001 | 1 |
| V003 | 3 |
