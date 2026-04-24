# Database Model V1

## Purpose

Version 1 of the data model defines the baseline conventions for tenant-specific backend databases used by the Access frontend.

## Model Assumptions

- each tenant has its own backend database file
- framework metadata is stored separately from transactional business data where practical
- optional modules extend the model in a controlled way
- naming should remain stable across languages; translations belong in UI or metadata tables

## Baseline Areas

- system metadata: tenant identity, schema version, language defaults, and configuration markers
- licensing metadata: enabled features, entitlement dates, and validation status
- shared reference data: reusable lookup tables required by the framework
- module data: tables introduced by optional features like `CAMT054`, `PROPERTY_MGMT`, and `WINE_MGMT`

## Versioning

Schema changes should be introduced through explicit framework versioning so the frontend can detect compatibility and apply controlled migrations when needed.

## Open Direction

Detailed table definitions are intentionally deferred until the framework services, licensing rules, and first module boundaries are finalized.

## Framework Tables (Frontend or Backend)

### Translations
- tblFwTranslations
  - key
  - language
  - value
  - active flag

### Help System
- tblFwTagHelp
  - token
  - description
  - example
  - category

### Tag Composer (temporary)
- tblTmpTagComposer
  - ControlName
  - TagValue