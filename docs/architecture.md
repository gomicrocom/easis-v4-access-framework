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

### 5. Business Application Layer

The project has now moved from the initial framework and infrastructure phase into the first business application phase.

This means the framework is no longer only preparing technical services. It is now actively hosting business entities, transactional workflows, reporting, and document output on top of the tenant backend structure.

Active storage landscape:

- `FE.accdb`
- `Data\<TENANT>_be.accdb`
- `Data\sys_be.accdb`
- `Data\log_be.accdb`

The tenant backend is now the primary home for business data and the first functional module scope: `BasicModule v1`.

---

## BasicModule v1

### Transition from Framework Phase to Business Phase

The framework phase established:

- startup and configuration loading
- tenant/backend separation
- translation handling
- tag-driven UI policy handling
- validation services
- reporting and export foundations
- logging and diagnostics

BasicModule v1 is the first module that uses these technical foundations to implement a concrete business workflow.

### Tenant Backend Usage

`Data\<TENANT>_be.accdb` is now actively used for:

- business master data
- order transactions
- payment and VAT references needed by order entry
- business document generation inputs

`sys_be.accdb` continues to hold global reference pools.

`log_be.accdb` remains the technical destination for operational logging as the logging architecture evolves.

### Core Business Entities

BasicModule v1 currently centers around the following tenant-backend tables:

- `tblAddresses`
  - address master data for customers, contacts, invoice addresses, and delivery addresses
- `art_product_group`
  - logical grouping and classification of sellable items
- `art_article`
  - article and service master data used in order lines
- `ord_order`
  - order header data including customer, status, dates, totals, and downstream document context
- `ord_order_line`
  - transactional line items belonging to an order
- `ten_payment_term`
  - tenant-side payment-term master data used by orders and document generation
- `ref_vat_code`
  - translated VAT code definitions used for pricing, tax logic, and document totals
- `ref_unit`
  - translated unit definitions for articles and order lines

### Order Workflow

The intended baseline business flow is:

1. maintain business partners and addresses
2. maintain articles and supporting references
3. create an order header
4. add order lines
5. calculate totals and VAT
6. generate a business document
7. export PDF and trigger mail delivery

This workflow is intentionally built on framework services instead of duplicating infrastructure logic inside forms.

### Framework Integration

BasicModule v1 depends on the framework layer in the following way:

- translations
  - UI captions, report labels, and document titles are resolved through the translation service
- tags
  - controls use managed `Tag` tokens for validation, behavior, access restrictions, and translation metadata
- validation
  - form input rules remain centralized in the framework runtime instead of being duplicated per form
- module access
  - role- and module-dependent availability is enforced through the existing framework access patterns
- reporting
  - reports consume prepared business data and framework translation logic
- logging
  - runtime diagnostics, validation issues, and service errors are written through centralized logging helpers

### UI Form Architecture

The business application layer now follows a standardized form naming convention.

This convention is intended to:

- keep modules structurally predictable
- simplify navigation and runtime handling
- reduce ad hoc naming decisions in future UI work
- support reusable workflow-oriented form patterns across modules

### Form Naming Convention

The following naming patterns are the standard for future business forms:

- `frm<Entity>List`
- `frm<Entity>Detail`
- `frm<Entity>Select`
- `frm<Entity>Dialog`
- `frm<Entity>Wizard`

### Form Type Purpose

- `frm<Entity>List`
  - list, navigation, search, and workflow entry form
- `frm<Entity>Detail`
  - record maintenance and editing form
- `frm<Entity>Select`
  - compact selection or lookup form used from other workflows
- `frm<Entity>Dialog`
  - focused modal or short interaction form
- `frm<Entity>Wizard`
  - guided multi-step workflow form

### Examples

- `frmAddressList`
- `frmAddressDetail`
- `frmOrderList`
- `frmOrderDetail`
- `frmArticleList`
- `frmArticleDetail`
- `frmInvoiceList`
- `frmInvoiceDetail`
- `frmCustomerAccountDialog`
- `frmSubscriptionWizard`

### Architectural Intent

The naming convention is not cosmetic only. It defines expected workflow roles:

- list forms are entry, navigation, and search forms
- detail forms are record editing forms
- select forms are helper forms for choosing existing business entities
- dialog forms are scoped interactions with a narrow purpose
- wizard forms guide the user through sequential business steps

The framework is therefore designed around reusable workflow-oriented UI patterns instead of isolated one-off forms.

### Application Shell

The frontend now also includes a first application-shell pattern:

- `frmAppShell`
  - persistent host form
- `frmAppNavigation`
  - left-side navigation surface
- `frmAppDashboard`
  - default workspace landing view

Shell behavior is service-driven rather than form-driven:

- `modAppShell`
- `modAppNavigationService`
- `modAppWorkspaceService`
- `modAppDashboardService`

Detailed shell notes are documented in [app-shell.md](./app-shell.md).

### Relationship to `fw_list_action`

Dynamic navigation from list forms should be driven through:

- `fw_list_action`

This means:

- `fw_list_action` stores `target_form` values
- naming consistency is important for generic navigation handlers
- future runtime handlers may dynamically open target forms based on naming conventions and action metadata

The preferred pattern is:

- `frm<Entity>List` as the navigation host
- `fw_list_action` as the configurable action source
- `frm<Entity>Detail` and related forms as targets

This keeps business navigation extensible without adding many hard-coded per-row UI actions.

---

## Business Tables Overview

### Master and Reference Scope

- `tblAddresses`
  - stores business partner and address master records
- `art_product_group`
  - stores product or service grouping definitions
- `art_article`
  - stores article master records including pricing and unit defaults
- `ten_payment_term`
  - stores tenant-level payment-term definitions for the business module
- `ref_vat_code`
  - stores VAT code and VAT-rate related reference definitions
- `ref_unit`
  - stores translated reusable units for quantities and article definitions

### Transaction Scope

- `ord_order`
  - stores order headers and commercial context
- `ord_order_line`
  - stores order positions and line-level commercial detail

### Business Document Direction

Orders are expected to become the operational basis for:

- printed documents
- PDF output
- email delivery
- later business document lifecycles such as invoice, delivery note, and follow-up handling

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
- table-driven (`fw_translation`)

---

### Tag Composer

Form:
- `frmTagComposer`

Features:
- visual editing of Tag strings
- multi-control editing
- temporary storage via `fw_tmp_tag_composer`
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

### Business Application Roadmap

Planned next steps:

- `frmAddresses`
- `frmArticles`
- `frmOrders`
- order line handling
- calculation services
- document generation
- mail handling

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
