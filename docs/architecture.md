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
- centralized Access SQL literal helpers (`modSqlHelper`)

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

- `adr_address`
  - address master data for customers, contacts, invoice addresses, and delivery addresses
- `art_product_group`
  - tenant-specific article-group master data used for article classification and future business grouping
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

### Language-Neutral Business Data

Business and master-data tables are intentionally language-neutral.

That includes values such as:

- `art_product_group.product_group_name`
- country names stored as business values
- currency names stored as business values

Localization belongs to UI and rendering layers only:

- forms
- navigation
- messages
- status texts
- reports
- document output

The framework therefore does not model multilingual duplicate business rows and
does not require translation joins for normal master-data access.

### Audit Field Convention

The framework uses a shared audit-field convention for bound record forms.

Standard field names:

- `created_at`
- `created_by`
- `updated_at`
- `updated_by`

Rules:

- `created_at` and `created_by`
  - are set only when a record is created
  - are not overwritten on later saves
- `updated_at` and `updated_by`
  - are refreshed whenever a record is saved

Implementation direction:

- reusable form-level audit handling belongs in `modAuditHelper`
- detail forms should use the shared helper instead of duplicating audit logic
- forms without audit fields must continue to save safely without runtime errors

### Schema Naming Convention

Physical schema objects must follow the documented SQL naming rules:

- table names use `snake_case`
- field names use `snake_case`
- no mixed-case field names
- no hyphenated field names
- no space-containing field names

Examples of valid physical field names:

- `order_no`
- `address_id`
- `document_status_code`
- `line_total`

Examples that must fail review if introduced into schema/code generation:

- `OrderNo`
- `Order-No`
- `Order No`
- `CustomerAddressId`

Compatibility note:

- `adr_address` is the authoritative address master table
- `tblAddresses` is deprecated and must not be recreated by current setup paths
- new field definitions inside schema creation must still use `snake_case`

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
- `frmArticleGroupList`
- `frmArticleGroupDetail`
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

### Shell-Owned List Command Bar

The framework now standardizes list actions through the shell itself.

Owner:

- `frmAppShell`

Shell command bar purpose:

- standardize common list actions
- reduce duplicated search and button wiring across list forms
- keep list-specific business navigation in the active workspace form

Standard layout:

- `Home`
- `Zurueck`
- `Suchen`
- `Leeren`
- `Neu`
- `Bearbeiten`
- `Aktualisieren`

List-form contract:

- `Public Function SupportsListCommandBar() As Boolean`
- `Public Sub ListSearch(ByVal searchText As String)`
- `Public Sub ListClearSearch()`
- `Public Sub ListNew()`
- `Public Sub ListEdit()`
- `Public Sub ListRefresh()`

Optional per-action capability flags:

- `Public Function SupportsListNew() As Boolean`
- `Public Function SupportsListEdit() As Boolean`

The shell command bar dispatches to the active workspace form through
`subWorkspaceHost.Form` and should not know:

- which detail form opens
- which primary key is selected
- which filters are applied
- which workspace target should be loaded

### List Form Architecture

List forms now follow one shared shell-driven pattern.

Ownership:

- `frmAppShell`
  - owns `Home`, `Zurueck`, `Suche`, `Leeren`, `Neu`, `Bearbeiten`, and `Aktualisieren`
- workspace list forms
  - own list-specific row sources, selected-record logic, workspace state, and detail-form targets

Required contract:

- `Public Function SupportsListCommandBar() As Boolean`
- `Public Sub ListSearch(ByVal searchText As String)`
- `Public Sub ListClearSearch()`
- `Public Sub ListNew()`
- `Public Sub ListEdit()`
- `Public Sub ListRefresh()`

Optional capability flags:

- `Public Function SupportsListNew() As Boolean`
- `Public Function SupportsListEdit() As Boolean`

Search-state pattern:

- `Private m_currentSearchText As String`
- `ListSearch(...)`
  - updates the current search text
- `ListClearSearch()`
  - clears the current search text
- filter application
  - is based on the list form's current search state
  - is no longer driven by a local search textbox

Search-index standard:

- `address_search_text`
- `article_group_search_text`
- `article_search_text`

Purpose:

- provide one centralized searchable text field per list row source
- keep filter expressions simple and Access-safe
- avoid duplicating multi-column search logic inside forms

Current note:

- `frmAddressList` still filters on legacy `AddressSearchText`
- this should be aligned to `address_search_text` in a future low-risk cleanup pass

Workspace interaction:

- list forms capture search and current-row context in `GetWorkspaceState()`
- detail forms are opened through `modAppWorkspaceService.OpenWorkspaceForm(...)`
- returning from detail uses workspace history and `RestoreWorkspaceState(...)`

Navigation philosophy:

- navigation contains areas and worklists only
- creation actions belong to `frmAppShell -> Neu`
- standard master-data objects should not add separate `Neue ...` navigation entries

Examples:

- `frmAddressList`
- `frmArticleGroupList`
- `frmArticleList`
- `frmFwTranslationAudit`

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

Current direction:

- the shell command bar is the default entry point for standard list actions
- `fw_list_action` remains available for richer object-specific workflows such as Address follow-up actions
- standard `New`, `Edit`, and `Refresh` behavior should not depend on `fw_list_action`

### Article Groups

The first additional tenant master-data object after the initial shell and framework rollout is:

- `Article Groups`

Implemented objects:

- table
  - `art_product_group`
  - primary key: `product_group_id`
- list form
  - `frmArticleGroupList`
- detail form
  - `frmArticleGroupDetail`
- service module
  - `modArticleGroupService`

The UI and business term remains `Artikelgruppe` / `Article Group`, while the
physical tenant table follows the existing Easis naming convention:

- physical table
  - `art_product_group`
- physical key fields
  - `product_group_id`
  - `product_group_code`
  - `product_group_name`

Navigation placement:

- `Mandant`
  - `Artikelgruppen`

The implementation follows the established shell-aware list/detail workflow:

- list form opens in the workspace
- detail form opens in the workspace
- add mode is driven by the shell command bar `Neu`
- back navigation restores the list state where possible

### Articles

The first article UI step is intentionally lightweight:

- physical table
  - `art_article`
- UI form
  - `frmArticleList`
- service module
  - `modArticleService`

Physical linkage:

- `art_article.product_group_id`
  - links to `art_product_group.product_group_id`
- `art_article.article_type_code`
  - links to `ref_article_type_code.article_type_code`

Current scope:

- article master data is listed in the workspace shell
- article detail maintenance is available through `frmArticleDetail`
- product-group names are resolved through a simple join for display
- shell command bar `Neu` and `Bearbeiten` are active for the article workflow
- `article_no` is auto-generated for new records when empty, using the first simple pattern `ART-000001`
- `article_type_code` remains a technical future-classification field and currently defaults to `ITEM`
- article business values remain language-neutral
- no inventory, stock tracking, variants, supplier logic, or complex pricing workflow is introduced yet

### Article Type Concept

Articles now distinguish between two independent concepts:

- business grouping
  - `product_group_id`
- behavior and future specialization
  - `article_type_code`

Examples of business grouping:

- Wine
- Furniture
- Services

Examples of article type behavior:

- `PRODUCT`
- `SERVICE`
- `SUBSCRIPTION`
- `FEE`
- `DISCOUNT`
- `WINE`
- `CUSTOM_SIZE`
- `APPAREL_SIZE`

This distinction is intentional:

- `product_group_id`
  - is used for business organization and reporting groups
- `article_type_code`
  - is reserved for runtime behavior, specialization, and future extensibility

### Future Article Extension Strategy

Future specialized article structures should extend `art_article` by article type, not by product group.

Planned extension directions include:

- `art_article_wine`
- `art_article_dimension`
- `art_article_apparel`
- `art_article_subscription`

These tables are not implemented yet.

The architectural rule is:

- `art_article`
  - remains the shared base article record
- future extension tables
  - attach to the base article by `article_id`
  - are activated or interpreted through `article_type_code`

Navigation placement:

- `Mandant`
  - `Artikel`

---

## Business Tables Overview

### Master and Reference Scope

- `adr_address`
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

### Address Table Authority

The address module now treats `adr_address` as the only authoritative address table.

Rules:

- current setup paths must not create `tblAddresses`
- new code must not query `tblAddresses`
- document and report joins must use `adr_address`
- any remaining `tblAddresses` mention is deprecated historical context only

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
- runtime translation marker support through `Control.Tag`

#### Official UI Translation Rule

For translatable UI controls:

- `Caption`
  - contains the readable fallback text shown in design mode
- `Tag`
  - contains `TR:<translation_key>`

Example:

- `Caption = Access Framework`
- `Tag = TR:FORM.FRMAPPSHELL.APP_SUBTITLE`

Important:

- `fw_translation.translation_key` stores only the pure key
- `TR:` is not stored in `fw_translation`
- new translation-tag maintenance writes the marker into `Tag`
- legacy caption-based `TR:` markers may still be tolerated by the runtime for backward compatibility

### SQL Helper Convention

Module:
- `modSqlHelper`

Purpose:
- central Access SQL literal formatting
- one shared place for string, nullable text, boolean, numeric-id, and date-time SQL literals

Recommended helpers:
- `SqlText(...)`
- `SqlNullableText(...)`
- `SqlBoolean(...)`
- `SqlLongOrNull(...)`
- `SqlDateTime(...)`

Project rule:
- modules and form classes should reuse the shared helpers
- new local `Private Function SqlText(...)` copies should not be introduced

---

### Tag Composer

Form:
- `frmTagComposer`

Features:
- visual editing of Tag strings
- multi-control editing
- temporary storage via `fw_tmp_tag_composer`
- preserves unrelated Tag segments while translation markers are managed separately
- prevents syntax errors

### Translation Maintenance

Form:
- `frmFwTranslations`

Purpose:
- translation maintenance
- translation key assignment
- `fw_translation` editing
- safe management of the `TR:` marker inside `Control.Tag`
- maintenance of both UI-bound and free/system translation namespaces

Official namespaces:

- `FORM.*`
  - form and control captions
- `NAV.*`
  - shell and navigation captions
- `MSG.*`
  - messages and dialog texts
- `STATUS.*`
  - status labels and status texts
- `REPORT.*`
  - report-focused captions and report text resources
- `DOCUMENT.*`
  - document generation texts and reusable output labels
- `REF.*`
  - reference and reusable display labels

Scope workflow:

- `FORM`
  - keeps the existing form/control workflow
  - translation keys are derived from the selected form/report and control context
- `NAV`, `MSG`, `STATUS`, `REPORT`, `DOCUMENT`, `REF`
  - allow direct maintenance of free keys
  - no form/control selection is required
  - key prefix must match the selected scope
- `ALL`
  - shows non-form keys across free/system namespaces
  - excludes `FORM.*`

Free-key scope specifics:

- `MSG`
  - accepts both `MSG.*` and legacy `MSG_*`
- `REF`
  - includes:
    - `ADDRESS_TYPE.*`
    - `SALUTATION.*`
    - `CONTACT_TYPE.*`
    - `UNIT.*`
    - `VAT.*`
    - `REF.*`

Important rules:

- `fw_translation.translation_key`
  - stores only the pure key such as `FORM.FRMAPPSHELL.APP_SUBTITLE`
- `TR:`
  - is a runtime/designer marker only
  - belongs in `Control.Tag`
  - must not be stored in `fw_translation.translation_key`

This is intentionally separate from `frmTagComposer`:

- `frmFwTranslations`
  - manages translation keys and translation data
  - manages `FORM.*`, `NAV.*`, `MSG.*`, `STATUS.*`, and `REF.*`
- `frmTagComposer`
  - manages validation, behavior, access, and other Tag metadata

### Translation Audit

Module:
- `modFwTranslationAuditService`

Purpose:
- build deterministic translation audit work data
- treat missing translations as an audit concern before they become a UI problem
- compare expected translation keys against `fw_translation`

Required languages:

- `DE-CH`
- `EN-US`
- `FR-FR`

Audit statuses:

- `OK`
- `MISSING_ROW`
- `EMPTY_VALUE`
- `ORPHAN`
- `LEGACY_KEY`

Current MVP expected-key sources:

- `fw_navigation.caption_key`
- reference tables with `translation_key` fields:
  - `ref_unit`
  - `ref_vat_code`
  - `ref_article_type_code`
  - `ref_address_type`
  - `ref_salutation`
  - `ref_addressing_mode`
  - `ref_contact_type`
- `FORM.*`
  - discovered from form and control `Tag` values containing `TR:...`
- registry-backed keys from `fw_translation_expected`
  - authoritative source for:
    - `MSG.*`
    - `STATUS.*`
    - `COMMON.*`
    - active legacy framework keys such as `MSG_*` and `ERR_*`
    - future manually registered framework keys

Current MVP limitation:

- report control scanning is not implemented yet
- VBA source parsing is intentionally avoided

Registry rationale:

- runtime fallback and runtime logging are not enough to manage translation completeness
- parsing VBA for `MSG.*` or `STATUS.*` usage would be fragile and expensive
- `fw_translation_expected` provides one deterministic audit source for framework message and status namespaces

Legacy translation treatment:

- historical translation rows are not deleted automatically
- actively referenced legacy framework keys should be registered in `fw_translation_expected`
- inactive historical keys should be classified as `LEGACY_KEY`
- true `ORPHAN` rows should represent audit-model inconsistency, not harmless historical leftovers

Runtime vs audit rule:

- runtime may keep using readable fallback text
- translation completeness must be measured through audit data, not only through runtime warnings

### Translation Workflow Split

Translation administration is intentionally split into two focused workspaces:

- `frmFwTranslationAudit`
  - translation control center
  - shell-command-bar list form
  - shell search handles free-text audit filtering
  - local dropdowns remain responsible for:
    - `scope_code`
    - `language_code`
    - `audit_status`
  - identifies completeness problems
  - opens the focused edit workspace for one key
  - keeps audit-specific action `Fehlende Eintraege erzeugen` local
- `frmFwTranslationEdit`
  - edits exactly one `translation_key`
  - shows the required language set together:
    - `DE-CH`
    - `EN-US`
    - `FR-FR`
  - saves values and returns to the previous workspace

`frmFwTranslations` is deprecated and should not be extended further. It may
remain temporarily during the transition, but the primary workflow is now:

`frmFwTranslationAudit` -> `frmFwTranslationEdit`

### DeepL Suggestion Workflow

`frmFwTranslationEdit` now supports a guided `DeepL Vorschlag` action.

Rules:

- DeepL provides suggestions only
- suggested values are filled into empty language fields
- existing non-empty target values are not overwritten without confirmation
- DeepL suggestions are not saved automatically
- the user must still review and click `Speichern`

Configuration:

- `DEEPL_API_KEY`
  - tenant parameter in `ten_parameter`
- `DEEPL_API_BASE_URL`
  - optional tenant parameter in `ten_parameter`
  - defaults to `https://api-free.deepl.com`
  - may be overridden with `https://api.deepl.com`

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
