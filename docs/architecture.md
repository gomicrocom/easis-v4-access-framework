# Architecture

## Overview

Easis Version 4 follows an MS Access VBA architecture with a tenant-facing frontend and a dedicated backend database for each tenant.

The system combines:
- a modular framework layer (runtime, services, UI patterns)
- optional business modules
- tenant-isolated data backends

---

## Core Principles

- one backend per tenant to simplify data isolation and operational ownership
- shared framework patterns for forms, reports, classes, queries, and modules
- tag-driven UI behavior using Control.Tag
- centralized validation and UX handling
- multilingual support through translation services
- feature-based licensing for module activation
- service-oriented VBA modules (low coupling, reusable logic)

---

## Logical Layers

### 1. Startup & Bootstrap
- configuration loading (INI-based)
- tenant resolution
- license initialization
- translation initialization
- module activation

### 2. Framework Services
- logging
- configuration
- translation (T / TEx)
- licensing
- navigation
- validation engine
- UI policy engine (Tag-based)

### 3. Application Modules
Optional feature packages:
- CAMT054
- PROPERTY_MGMT
- WINE_MGMT

Modules integrate via framework services.

### 4. Data Access
- linked tables per tenant backend
- query-based access
- import/export services

---

## Runtime Framework (Access UI Layer)

### Tag System

Controls use the `Tag` property for declarative behavior.

Supported tokens:

- REQUIRED
- NUMERIC
- INTEGER
- DATE
- MIN / MAX
- MINLEN / MAXLEN
- READONLY / LOCKED / DISABLED
- ROLE / HIDDEN / SETFOCUS

Parsed via:
- `ParseTagTokens`

---

### Validation Engine

Centralized in:
- `modFormRuntime`

Features:

- rule-based validation via tags
- per-field validation messages
- summary message output
- first invalid control gets focus
- inline highlighting of invalid controls
- original control colors restored after validation
- hidden and disabled controls are excluded

---

### Translation System (i18n)

Module:
- `modTranslationService`

Functions:
- `T(key, fallback)`
- `TEx(key, fallback, args...)`

Features:
- placeholder support `{0}`, `{1}`
- multi-language (EN / DE)
- table-driven (`tblFwTranslations`)

---

### Tag Composer

Form:
- `frmTagComposer`

Features:
- visual editing of Tag strings
- multi-control editing
- temporary storage via `tblTmpTagComposer`
- preserves `TR:*` tags
- prevents syntax errors

---

### Help System

- token documentation stored in table
- seeded via setup
- accessible from UI tools

---

## Configuration Direction

System configuration is provided via INI files:

- backend location
- tenant identifier
- default language
- enabled modules
- licensing parameters
- output directories (planned)

Config location:
- `<AppPath>\Cfg\easis.ini`

---

## Deployment Direction

- frontend distributed centrally
- backend per tenant
- supports controlled updates and module rollout

---

## Next Architecture Phase

### Core Services (planned)

- Document / PDF generation
- Output path service:
  - `<DocumentDirectory>\<CustomerName>\DocNumber.pdf`
- QR code integration
- Email service (CDO)
- Batch processing (print, email, dunning, subscriptions)
- CAMT.054 import
- NAPS2 scan integration

---

## Target Service Architecture

### Core Services
- modOutputPathService
- modPdfExportService
- modDocumentService
- modEmailService

### Orchestration
- modBatchHandler

### Integrations
- modCamt054Service
- modScanIntegrationService
