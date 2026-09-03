# Specification generated from the mapping workbook

* Workbook : `/tmp/casa_map.xlsx`
* Output   : `samples/casa_generated.yaml`
* Generator: SOR Polymorphizer v5.2 (workbook-only mode)

## What was built

* Operations            : **26**
* Request schemas       : 26
* Response schemas      : 26
* Parameters            : 182
* Response headers      : 78

## What the workbook contained

* Body rows read        : 2111
* Container rows        : 506
* Mapped leaf attributes: **653**  (emitted)
* Marked no SOR field   : 952  (never emitted)

* Usage column populated: 298 of 2111 (14%). `required` was emitted only where Usage reads Mandatory, so the remaining attributes carry no requiredness. Populate the Usage column to fix this at source.

## There is no match rate in this mode

The workbook is both the input and the arbiter, so a coverage figure would be 100% by construction and would mean nothing. The findings below are the substitute: they are the workbook's own data-quality defects.

## Data-quality findings (7)

* CASA_AcctRelation_retrieve!60 CurrentAndSavingsAccountFacilityAccountRelationRetrieveResponse.AccountInvolvement.Party.PartyObligationOrEntitlement.EntitlementArrangement: data type 'string' but rows exist beneath it; treated as an object
* CASA_AcctDetails_retrieve!73 CurrentAndSavingsAccountFacilityRetrieveResponse.Account.Associations.AccountInvolvement.Party.PartyObligationOrEntitlement.EntitlementArrangement: data type 'string' but rows exist beneath it; treated as an object
* AccountLimit_Retrieve!45 AccountLimitResponse.status.statusCode: written as 'status' but declared as 'Status' (casing differs; matched anyway)
* DebitandCredit_Retrieve!91 DebitAndCreditResponse.DepositTransaction.TransactionUserReference: data type 'string' but rows exist beneath it; treated as an object
* AmountBlock_Retrieve!32 AmountBlockRetrieveRequest.AmountBlockProcessingInstruction.Instruction: data type 'String' but rows exist beneath it; treated as an object
* Passbook_Retrieve!65 PassbookRetrieveResponse.PassbookInstanceRecord.PassbookEntryTransaction.transactionAmount: data type 'Number' but rows exist beneath it; treated as an object
* Passbook_Retrieve!66 PassbookRetrieveResponse.PassbookInstanceRecord.PassbookEntryTransaction.TransactionAmount.amountValue: written as 'PassbookInstanceRecord.PassbookEntryTransaction.TransactionAmount' but declared as 'PassbookInstanceRecord.PassbookEntryTransaction.transactionAmount' (casing differs; matched anyway)

## Output assertions

* Passed: no empty object schemas, no dangling references, no orphaned required entries.
