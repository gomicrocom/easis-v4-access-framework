# Easis Version 4

Easis Version 4 is an MS Access VBA framework for tenant-aware business applications with a shared frontend pattern and one backend database per tenant.

## Foundation

- MS Access frontend with VBA-based application framework
- one backend per tenant for data isolation and deployment flexibility
- INI-based configuration for environment and tenant settings
- multilingual support for UI text, captions, and module labels
- feature-based licensing for controlled module activation
- Optional domain modules such as `CAMT054`, `PROPERTY_MGMT`, and `WINE_MGMT`

## Repository Layout

- `docs/` contains architecture, roadmap, licensing, and data model notes
- `src/access/exported/` contains exported Access objects grouped by type
- `scripts/` contains local automation and Access export/import helpers
- `tests/` contains validation assets for framework-level checks

## Current Status

This repository currently provides the initial project scaffold and documentation baseline. Business logic, VBA implementations, and Access object exports will be added incrementally.
