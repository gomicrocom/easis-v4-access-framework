# Translation Rules

## Purpose

This document defines the official translation policy for Easis v4.

The policy is intentionally practical:

- new user-visible UI text must be translation-key based
- working forms are not mass-migrated in one step
- cleanup happens in phases
- translation metadata must stay predictable for designers, runtime code, and maintenance tools

## Official Rules

### UI Controls

For translatable controls:

- `Caption`
  - contains the readable fallback or design text
- `Tag`
  - contains `TR:<translation_key>`
- `fw_translation.translation_key`
  - stores the pure key only
  - never stores the `TR:` prefix

Example:

- `Caption = Access Framework`
- `Tag = TR:FORM.FRMAPPSHELL.APP_SUBTITLE`
- `fw_translation.translation_key = FORM.FRMAPPSHELL.APP_SUBTITLE`

Important:

- `TR:` is a runtime and designer marker only
- `TR:` belongs in `Control.Tag`
- `TR:` must not be stored in `fw_translation.translation_key`

### Code Messages

User-facing code messages should use translation helpers instead of hardcoded visible text.

Recommended pattern:

```vb
MsgBox modFwTranslationRuntime.ResolveText( _
    "MSG.TRANSLATION_KEY_MISSING", _
    "translation_key fehlt.")
```

The fallback text must remain readable and safe if no translation exists.

### New Translation Structure Standard

All newly introduced translation structures must include at least:

- `DE-CH`
- `EN-US`
- `FR-FR`

This is now the framework minimum standard for new reference namespaces,
framework UI structures, and newly introduced translation-backed concepts.

## Translation Audit Responsibility

Missing translations are primarily an audit problem, not a runtime problem.

Runtime behavior:

- runtime may continue to use fallback captions and fallback text
- runtime warnings are still acceptable as a secondary safety net

Audit responsibility:

- expected translation keys must be collected deterministically
- required languages must be checked centrally
- missing, empty, and orphaned translation rows must be reported through audit data

Required audit statuses:

- `OK`
- `MISSING_ROW`
- `EMPTY_VALUE`
- `ORPHAN`
- `LEGACY_KEY`

Meaning:

- `OK`
  - expected translation row exists and contains a value
- `MISSING_ROW`
  - expected translation row does not exist
- `EMPTY_VALUE`
  - translation row exists but has no text value
- `ORPHAN`
  - translation exists but the expected-source model is inconsistent
- `LEGACY_KEY`
  - historical translation exists, but no active expected-source reference is currently modeled

Current audit MVP sources:

- `fw_navigation.caption_key`
- reference tables with `translation_key` fields:
  - `ref_unit`
  - `ref_vat_code`
  - `ref_article_type_code`
  - `ref_address_type`
  - `ref_salutation`
  - `ref_addressing_mode`
  - `ref_contact_type`
- form and control `Tag` values containing `TR:FORM.*`
- `fw_translation_expected`
  - authoritative registry for:
    - `MSG.*`
    - `STATUS.*`
    - `COMMON.*`
    - active legacy framework keys such as `MSG_*` and `ERR_*`
    - future manually registered framework keys

Current audit MVP limitation:

- `REPORT.*` control-tag scanning is not implemented yet
- VBA source parsing is intentionally not part of the audit design

Why VBA parsing is avoided:

- deterministic registry rows are easier to maintain and review
- runtime message usage can evolve without making the audit brittle
- framework keys such as `MSG.*`, `STATUS.*`, and `COMMON.*` remain centrally auditable without a code parser

Legacy-key treatment:

- historical translations are not deleted automatically
- actively used legacy framework keys should be registered in `fw_translation_expected`
- inactive historical keys should be classified as `LEGACY_KEY`
- namespace migration can happen later without changing current runtime behavior

## Language-Neutral Business Data Policy

Business, master, and reference table content is intentionally language-neutral.

This means values stored in business tables stay stable and are not translated in
place and are not duplicated per language.

Examples:

- `art_product_group.product_group_name`
- `refCountries.country_name`
- `refCurrencies.currency_name`

These remain business data values, not UI translation resources.

### What Is Localized

Localization belongs to the presentation layers only:

- UI
- navigation
- messages
- status texts
- reports
- document rendering and output

These use translation namespaces such as:

- `FORM.*`
- `NAV.*`
- `MSG.*`
- `STATUS.*`
- `COMMON.*`
- `REPORT.*`

Document output must later render in the correspondence language of the
recipient, but the underlying business data rows remain language-neutral.

### What This Means Architecturally

- no translation joins for business tables
- no multilingual duplicate rows in master data tables
- no `PRODUCT_GROUP.*` translation records for business row content
- demo and seed business values remain stable business values

This is an intentional architecture decision, not a limitation.

## Current And Future FORM/REPORT Architecture

### Current State

Today the translation architecture is intentionally asymmetric:

- `FORM`
  - object-aware
  - control-aware
  - Tag-based
  - runtime-localized
- `REPORT`
  - currently treated as a free-key namespace in `frmFwTranslations`
  - valid for keys such as `REPORT.*`
  - not yet fully implemented as an object-aware report translation workflow

This transitional state is acceptable and should remain stable until the report-aware architecture is introduced deliberately.

### Official Future Target

The future target is architectural parity between `FORM` and `REPORT`.

Target outcome:

- `FORM`
  - object-aware
  - control-aware
  - Tag-based
  - runtime-localized
- `REPORT`
  - object-aware
  - report-control-aware
  - Tag-based
  - runtime-localized
  - integrated into `frmFwTranslations` similarly to `FORM`

### FORM Rule

For forms:

- `Caption`
  - readable design-time fallback
- `Tag`
  - `TR:FORM.<FORMNAME>.<KEY>`
- `fw_translation.translation_key`
  - stores the pure key only
  - never the `TR:` prefix

Examples:

- `FORM.FRMADDRESSDETAIL.CUSTOMER_NAME`
- `FORM.FRMAPPSHELL.APP_TITLE`

### REPORT Rule

Future rule for reports:

- `Caption`
  - readable design-time fallback
- `Tag`
  - `TR:REPORT.<REPORTNAME>.<KEY>`
- `fw_translation.translation_key`
  - stores the pure key only
  - never the `TR:` prefix

Future examples:

- `REPORT.RPTINVOICE.TITLE`
- `REPORT.RPTINVOICE.FOOTER`
- `REPORT.RPTORDER.POSITION_HEADER`
- `REPORT.RPTPAYMENTREMINDER.WARNING`

### Future Runtime Target

The future report runtime should:

- iterate report controls
- evaluate `TR:` tags
- resolve `REPORT.*` keys
- apply translated captions or text
- use fallback captions when translations are missing
- behave similarly to form runtime localization

Important:

- this future report-aware runtime is not fully implemented yet
- current free-key `REPORT.*` usage remains valid during transition
- existing report behavior should not be broken during the transition

## Official Namespaces

### `FORM.*`

Use for:

- form captions
- label captions
- button captions
- user-visible captions belonging to a specific form or report control

Pattern:

- `FORM.<FORMNAME>.<CONTROL_OR_MEANING>`

Examples:

- `FORM.FRMAPPSHELL.APP_TITLE`
- `FORM.FRMAPPSHELL.APP_SUBTITLE`
- `FORM.FRMADDRESSDETAIL.FORM_TITLE`

### `NAV.*`

Use for:

- shell navigation captions
- menu items
- navigation groups

Examples:

- `NAV.GROUP.ADDRESSES`
- `NAV.ADDRESS_LIST`

### `MSG.*`

Use for:

- message boxes
- validation messages
- user-facing warnings
- user-facing confirmations

Legacy support:

- existing `MSG_*` keys may still exist
- new keys should prefer `MSG.*`

### `STATUS.*`

Use for:

- status bar text
- runtime status labels
- user-facing state text

Example:

- `STATUS.READY`

### `COMMON.*`

Use for:

- reusable standard UI actions
- shared button captions
- common list-form verbs

Examples:

- `COMMON.NEW`
- `COMMON.EDIT`
- `COMMON.REFRESH`

### `REPORT.*`

Use for:

- report labels
- report captions
- report-specific visible text

Current state:

- supported as a free-key namespace
- transitional until report-aware translation composition is introduced

### `DOCUMENT.*`

Use for:

- document type labels
- document-specific output labels
- reusable visible text in document generation

### `REF.*`

Use for:

- generic reference values where no more specific namespace exists

Reference-specific namespaces are also allowed and preferred where clearer:

- `ADDRESS_TYPE.*`
- `REF.ARTICLE_TYPE.*`
- `SALUTATION.*`
- `CONTACT_TYPE.*`
- `UNIT.*`
- `VAT.*`

## frmFwTranslations Scope Behavior

`frmFwTranslations` now supports these scopes:

- `FORM`
- `NAV`
- `MSG`
- `STATUS`
- `REPORT`
- `DOCUMENT`
- `REF`
- `ALL`

### `FORM`

- keeps the existing form/control workflow
- translation keys are derived from form/report and control context
- this workflow should not be broken by free-key maintenance changes

### `REPORT`

Current behavior:

- works as a free-key namespace
- no report selection is required
- no report-control selection is required

Future target:

- optional report selection
- optional report-control selection
- automatic `REPORT.*` key generation
- report translation composition similar to `FORM`

### Free-Key Scopes

For `NAV`, `MSG`, `STATUS`, `REPORT`, `DOCUMENT`, and `REF`:

- no object/control selection is required
- translations are maintained directly in `fw_translation`
- `translation_key`, `language_code`, `translation_value`, `module_code`, `is_active`, and `sort_order` are edited directly

### `REF`

`REF` scope includes:

- `ADDRESS_TYPE.*`
- `SALUTATION.*`
- `CONTACT_TYPE.*`
- `UNIT.*`
- `VAT.*`
- `REF.*`

### `ALL`

`ALL` is a free-key maintenance view for non-form keys.

Rules:

- includes non-`FORM.*` keys
- excludes `FORM.*`

## Validation Rules

### Required

On save:

- `translation_key` is required
- `language_code` is required
- `translation_value` is required unless a future project convention explicitly allows placeholders

### Forbidden

Do not store:

- `TR:FORM.X.Y`
- caption fallback text instead of a key

Reject keys with:

- `TR:`

### Scope Prefix Rules

Expected prefixes:

- `FORM` -> `FORM.*` and existing `REPORT.*` where the form workflow still uses them
- `NAV` -> `NAV.*`
- `MSG` -> `MSG.*` and legacy `MSG_*`
- `STATUS` -> `STATUS.*`
- `REPORT` -> `REPORT.*`
- `DOCUMENT` -> `DOCUMENT.*`
- `REF` -> `ADDRESS_TYPE.*`, `SALUTATION.*`, `CONTACT_TYPE.*`, `UNIT.*`, `VAT.*`, `REF.*`
- `ALL` -> any non-`FORM.*` key

## Cleanup Policy

### Phase 1

- no new hardcoded user-visible UI text
- new forms and modules must use translation keys immediately

### Phase 2

Clean framework, shell, and admin forms first:

- `frmAppShell`
- `frmAppNavigation`
- `frmAppDashboard`
- `frmFwTranslations`
- `frmFwNavigationAdmin`
- `frmTagComposer`
- `frmTagHelp`

### Phase 3

- migrate user-facing `MsgBox` text to `MSG.*`
- migrate user-facing status text to `STATUS.*`
- keep technical logs readable
- do not over-translate purely technical diagnostics unless they are user-facing

### Phase 4

Migrate business forms when they are touched:

- Address
- Articles
- Orders
- Finance
- Reports

This is an incremental cleanup strategy, not a one-shot rewrite.

## FORM/REPORT Transition Strategy

### Phase 1

- stabilize `FORM` architecture
- stabilize translation runtime
- stabilize `frmFwTranslations` scopes

### Phase 2

- use `REPORT.*` safely as a free-key namespace
- keep existing report behavior stable

### Phase 3

- introduce report-aware translation scanning and composition
- prepare report object/control selection workflows where justified

### Phase 4

- unify `FORM` and `REPORT` runtime localization architecture
- converge maintenance workflows in `frmFwTranslations`

## What Not To Do

Do not:

- put `TR:` into `fw_translation.translation_key`
- overwrite `Caption` with unreadable key text for new work
- add new user-facing hardcoded captions when a translation key should be used
- mass-migrate old forms without testing
- break existing `FORM` workflow while improving free-key scopes

## Helper Usage

Shared runtime helper:

- `modFwTranslationRuntime.ResolveText(translation_key, fallback_text)`

Purpose:

- returns translated text if available
- returns fallback text if missing
- never raises an error to the caller
- avoids excessive missing-translation warnings for normal fallback usage

Existing form validation and maintenance tools in `frmFwTranslations` can help identify missing translations over time, but this document does not require a full repository scanner in the current phase.

## Translation Administration Workflow

The framework now uses a focused administrator workflow:

- `frmFwTranslationAudit`
  - translation quality-control workspace
  - shell-command-bar list form
  - shell search handles free-text audit filtering
  - local scope/language/status dropdowns remain audit-specific filters
  - opens the focused edit workspace for one key
- `frmFwTranslationEdit`
  - edits exactly one `translation_key`
  - shows the required language set together:
    - `DE-CH`
    - `EN-US`
    - `FR-FR`
  - saves values and returns to the previous workspace

`frmFwTranslations` is deprecated. It may remain temporarily during the
transition, but it must not be extended further.

The intended workflow is:

- open `frmFwTranslationAudit`
- select one audit row / `translation_key`
- open `frmFwTranslationEdit`
- maintain the three required language values for that key

Audit-list action model:

- shell command bar
  - `Search`
  - `Clear Search`
  - `Edit`
  - `Refresh`
- local audit form
  - keeps domain-specific filters
  - keeps domain-specific action `Fehlende Eintraege erzeugen`

## DeepL Suggestion Rule

`frmFwTranslationEdit` may request DeepL suggestions for missing language
values, but this remains a review-first workflow.

Rules:

- DeepL fills suggestions only
- existing values are not overwritten without confirmation
- DeepL suggestions are never auto-saved
- final persistence still happens only through `Speichern`

Configuration:

- `DEEPL_API_KEY`
  - expected in tenant parameter `ten_parameter.param_key`
- `DEEPL_API_BASE_URL`
  - optional override in tenant parameter `ten_parameter.param_key`
  - default is `https://api-free.deepl.com`
