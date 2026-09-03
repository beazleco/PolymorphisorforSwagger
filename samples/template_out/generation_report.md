# SOR Polymorphizer generation report

Workbook: `SOR_mapping_template.xlsx`

Generated 3 of 3 endpoints. **0 failed.**

## Endpoints generated

| Sheet | Pattern | SOR endpoints | Method | Path |
|---|---|---|---|---|
| Example_SingleSor | P1 | 1 | POST | `/service-domain/resource/action` |
| Example_MultiSor | P2 | 3 | POST | `/service-domain/resource/action` |
| Operation_Template | P1 | 1 | POST | `/service-domain/resource/action` |

### Patterns

- **P1**, one SOR endpoint: a single schema per message, pruned to what SOR supplies.
- **P2**, several SOR endpoints with a `requestVariant` enumeration: base plus `allOf` variants, wrapped in `oneOf` with a discriminator on the request.
- **P3**, several SOR endpoints with no tag: `oneOf` without a discriminator.

## Element reduction on the polymorphic endpoints

| Sheet | Variant | Attributes in the flat schema | In this variant |
|---|---|---|---|
| Example_MultiSor | AccountRetrieveRequestForAccountCardList | 4 | 4 |
| Example_MultiSor | AccountRetrieveRequestForAccountContactDetails | 4 | 4 |
| Example_MultiSor | AccountRetrieveRequestForAccountInformation | 4 | 3 |

## Support sheets

No Level columns, so these are not operation sheets: `How to use`, `Vocabulary`.
