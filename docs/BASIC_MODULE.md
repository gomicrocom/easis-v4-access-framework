# BasicModule v1

## Overview

BasicModule v1 is the first concrete business application layer built on top of the Easis v4 Access framework.

It marks the transition from a framework-first repository into an actively used business solution with tenant-specific master data, transactions, reporting, document generation, and operational output flows.

The module uses the existing backend separation model:

- `FE.accdb`
- `Data\<TENANT>_be.accdb`
- `Data\sys_be.accdb`
- `Data\log_be.accdb`

Business data is primarily stored in the tenant backend.

## Scope

BasicModule v1 currently covers the first core business entities and process chain:

- addresses
- article master data
- product groups
- order headers
- order lines
- payment terms
- VAT codes
- units

The focus is a stable baseline for order entry and document preparation.

## Intended Workflow

The intended operational flow is:

`Address -> Order -> OrderLines -> Document -> PDF -> Mail`

In more detail:

1. maintain customer and address master data
2. maintain articles and article grouping
3. create an order header
4. add order lines
5. calculate line and order totals
6. prepare the business document layer
7. generate PDF output
8. send the result by mail

The framework is responsible for reusable runtime behavior, while the business module owns the domain entities and process flow.

## Current Implementation Status

The current repository state supports the first business phase through:

- tenant backend usage for active business data
- framework-driven forms, validation, and translation infrastructure
- document-related calculation and reporting groundwork
- PDF export groundwork
- translation maintenance tooling
- tag maintenance tooling

BasicModule v1 database scope is now explicitly active in the tenant backend through:

- `tblAddresses`
- `art_product_group`
- `art_article`
- `ord_order`
- `ord_order_line`
- `ref_payment_term`
- `ref_vat_code`
- `ref_unit`

At this stage, the technical framework and the first business entities already interact, but the end-to-end business UI is still being expanded.

## Framework Integration

BasicModule v1 relies on the framework for:

- translations
  - multilingual captions, labels, and document texts
- tags
  - declarative control behavior and validation metadata
- validation
  - centralized runtime validation via framework services
- module access
  - role and module based access handling
- reporting
  - document/report rendering on top of prepared business data
- logging
  - diagnostics, runtime tracing, and error handling

The business module should use these services rather than reimplementing technical logic locally in forms.

## Planned Next Steps

- `frmAddresses`
- `frmArticles`
- `frmOrders`
- order line handling
- calculation services
- document generation
- mail handling

## Notes

BasicModule v1 is intentionally a first business baseline, not a complete ERP layer.

The current direction is to keep:

- business rules inside services and repositories
- UI logic thin
- framework services reusable across future modules
- tenant data isolated in the tenant backend
