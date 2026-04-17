# Easis Version 4

Easis Version 4 is an MS Access VBA framework for building tenant-aware business applications with a shared frontend structure and one backend database per tenant.

## Scope

- MS Access frontend with VBA-based application framework
- One backend per tenant for data isolation and deployment flexibility
- INI-based system configuration for environment and tenant settings
- Multilingual support for user interface text and module labels
- Feature-based licensing for controlled module activation
- Optional domain modules such as `CAMT054`, `PROPERTY_MGMT`, and `WINE_MGMT`

## Repository Structure

- `docs/` contains architecture, roadmap, licensing, and data model notes
- `src/access/exported/` contains exported Access objects grouped by type
- `scripts/` is reserved for automation and export/import helper scripts
- `tests/` is reserved for framework-level validation assets and test utilities

## Current Status

This repository currently contains the initial project scaffold and documentation baseline. Business logic, VBA implementations, and Access object exports will be added incrementally.
