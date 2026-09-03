# FAILED: Card RSAEncrypted

No specification was generated for this endpoint. The rest of the workbook was processed normally.

## R001 at Card RSAEncrypted!H1

**Found:** the column header names the SOR endpoint '/v1/card/cancel', but the SOR API Endpoint banner names '/v1/card/{}/cardRSAEncrypted'. The two disagree and the tool does not guess which is right

**Expected:** the column header and the banner to name the same endpoint

**Fix:** correct whichever is wrong. Put every SOR endpoint this operation calls in the SOR API Endpoint banner cell, and head each SOR column with the endpoint that column maps to.
