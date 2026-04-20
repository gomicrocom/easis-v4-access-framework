# Module Catalog

## Framework Core

The framework core contains startup, configuration, translation, licensing, navigation, and shared utility services needed by every tenant installation.

## Optional Modules

### `CAMT054`

Intended for bank statement or payment notification processing based on CAMT.054-related workflows.

### `PROPERTY_MGMT`

Intended for property management scenarios such as units, contracts, charges, and operational administration.

### `WINE_MGMT`

Intended for wine-related operations such as inventory, classification, movements, and domain-specific reporting.

## Rules

- modules should be independently licensable
- modules may add forms, reports, queries, classes, and VBA modules
- modules should integrate through framework services rather than direct hard coupling
- multilingual labels and messages should be provided consistently per module

## Extension Direction

Additional modules can be added to the catalog once their business boundaries, licensing needs, and backend model extensions are defined.
