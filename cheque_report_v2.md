# SOR-driven polymorphism refactor

* Source swagger: `/root/.claude/uploads/65f30f0e-ba98-58cd-a210-150f3bfe3d35/e5c13e2b-apicoechequeprocessingswaggerv1.0.0.yaml`
* SOR mapping:   `/root/.claude/uploads/65f30f0e-ba98-58cd-a210-150f3bfe3d35/91868040-datamodelChequeProcessingV1__Example.xlsx`
* Output:        `cheque_polymorphic_v2.yaml`

## Attributes eliminated (8)

* `AccountIdentifier.accountIdentificationType`  — Not Used In SOR / blank
* `ChequeImage.Image Description`  — Not Used In SOR / blank
* `Amount.amountType`  — absent-from-mapping
* `ChequeItem.chequeReferenceNumber`  — absent-from-mapping
* `Account.accountNickName`  — absent-from-mapping
* `Account.accountOpenDate`  — absent-from-mapping
* `ChequeSubmissionResponse.ChequeItem[].chequeReferenceNumber`  — absent-from-mapping
* `ChequeEnquiryResponse.ChequeItem[].chequeReferenceNumber`  — absent-from-mapping

## Polymorphism applied

Base ChequeProcessingResultBase holds common ['ChequeItem'] + discriminator 'resultContext'.
ChequeSubmissionResponse -> allOf[ChequeProcessingResultBase]
ChequeEnquiryResponse -> allOf[ChequeProcessingResultBase, +['ChequeLifestyleStatus']]

## De-duplication

* ChequeEnquiryRequest.chequeProcessingOperatingSession -> $ref chequeProcessingOperatingSession (identical shape)
