# UI Layout Guide

## Goal

The Easis v4 Access Framework uses a fixed shell and workspace layout so future forms can be designed consistently.

This keeps the frontend predictable across modules:

- the shell remains stable
- navigation stays in one place
- workspace forms share one standard canvas
- status information stays compact and visible

## Shell Architecture Status

The shell architecture is now implemented and no longer only conceptual.

Currently implemented:

- `frmAppShell`
- accordion navigation
- workspace form hosting
- workspace history
- back navigation
- workspace-aware list and detail workflows
- embedded form loading
- shell-aware navigation service
- workspace-aware filtering

The application now behaves as a single-window, workspace-oriented Access frontend.

Shell translation behavior is now part of that architecture:

- navigation captions can be translated
- shell labels can be translated
- dashboard card titles can be translated
- fallback text remains mandatory for safe operation

## Shell Layout

Recommended `frmAppShell` structure:

- Form Header
  - not used
  - hidden
- Detail section
  - custom header zone
  - toolbar zone
  - left navigation host
  - right workspace host
- Form Footer
  - status bar

The shell should remain open as the application host.

## Recommended Shell Size

Recommended overall shell size:

- `frmAppShell` total: approx. `1400 x 850 px`
- Access design size: approx. `37.0 cm x 22.5 cm`

These values are intended as a practical baseline, not a pixel-perfect restriction.

## Detail Section Zones

Recommended approximate positions inside the detail section:

### Header Zone

- Top: `0.0 cm`
- Left: `0.0 cm`
- Width: `37.0 cm`
- Height: `2.5 cm`

Use this zone for:

- application title
- tenant branding if needed
- lightweight shell commands

### Toolbar Zone

- Top: `2.6 cm`
- Left: `0.2 cm`
- Width: `36.0 cm`
- Height: `1.0 cm`

Use this zone for:

- context actions
- simple navigation shortcuts
- refresh or workflow buttons

### Navigation Host

- Top: `3.8 cm`
- Left: `0.2 cm`
- Width: `7.0 cm`
- Height: `17.0 cm`

Recommended host:

- `subNavigationHost`

### Workspace Host

- Top: `3.8 cm`
- Left: `7.4 cm`
- Width: `28.5 cm`
- Height: `17.0 cm`

Recommended host:

- `subWorkspaceHost`

This workspace area is the standard content canvas for most application forms.

## Workspace Hosting Rules

Workspace forms should behave as embedded forms inside `frmAppShell.subWorkspaceHost`.

Recommended rules:

- avoid `DoCmd.OpenForm` for normal workspace navigation
- use `modAppWorkspaceService.OpenWorkspaceForm(...)`
- avoid popup workflows
- avoid modal workflows unless a true dialog is needed
- reuse the shell workspace instead of stacking floating forms

The same conservative design principle also applies to shell translations:

- translated labels should improve usability
- missing translations must not block navigation or shell loading

Example:

Old:

```vb
DoCmd.OpenForm "frmAddressDetail"
```

Preferred:

```vb
modAppWorkspaceService.OpenWorkspaceForm _
    Me.Parent, _
    "frmAddressDetail", _
    "[address_id]=123"
```

## Footer / Status Bar

Recommended footer usage:

- Form Footer height: `0.7 cm` to `1.0 cm`
- used for status bar only

Suggested compact status line:

`v0.5.0-dev | DEFAULT | ADMIN | Ready`

Recommended status controls:

- `txtStatusAppVersion`
- `txtStatusCurrentUser`
- `txtStatusCurrentTenant`
- `txtStatusCurrentRole`
- `txtStatusBackend`
- `txtStatusEnvironment`

Optional future simplification:

- `txtStatusLine`

## Standard Workspace Canvas

Recommended standard workspace form settings:

- Width: `28.5 cm`
- Height: `17.0 cm`
- `AutoResize = No`
- `AutoCenter = No`
- `PopUp = No`
- `Modal = No`
- `RecordSelectors = No`
- `NavigationButtons = No`
- `ScrollBars = None` or `Vertical only`

This standard canvas helps workspace forms fit reliably inside `frmAppShell` without ad hoc resizing.

## Translation-Aware Shell Notes

The shell layer is translation-aware, but it remains fallback-first.

Recommended practice:

- define translation keys for shell labels and navigation entries
- always keep a useful fallback caption
- avoid relying on translated text for program logic

Navigation entries should therefore always store:

- `caption_key`
- `fallback_caption`

This keeps the shell usable even when:

- translation seeding has not run yet
- the selected language has missing values

### Form Translation Marker Rule

For shell controls and other translatable form controls, the standard rule is:

- `Caption`
  - readable fallback or design text
- `Tag`
  - `TR:<translation_key>`

Example:

- `Caption = Access Framework`
- `Tag = TR:FORM.FRMAPPSHELL.APP_SUBTITLE`

Important:

- `fw_translation.translation_key` stores only the pure key such as `FORM.FRMAPPSHELL.APP_SUBTITLE`
- `TR:` is only a runtime and designer marker
- new translation assignments should update `Tag`, not `Caption`
- legacy caption-based markers may still be tolerated temporarily, but they are no longer the preferred pattern

## Workspace History / Back Navigation

Workspace history is now part of the implemented shell behavior.

Back navigation can restore:

- the previous form
- filters
- search state
- selected record when possible

Key shell behavior:

- previous workspace states are preserved in memory
- `cmdBack` can trigger back navigation from the shell
- `modAppWorkspaceService.GoBack(...)` restores the previous workspace
- `modAppWorkspaceService.CanGoBack()` can be used to enable or disable back UI

Optional workspace state contract for forms:

```vb
Public Function GetWorkspaceState() As String
Public Sub RestoreWorkspaceState(ByVal stateText As String)
```

Workspace-aware forms should implement these methods when meaningful.

This is especially useful for:

- list forms with search/filter context
- list/detail workflows
- forms that need to restore selected records

## Workspace Workflow Pattern

Preferred shell workflow:

- `List -> Detail -> Back -> continue workflow`

Examples:

- `AddressList -> AddressDetail -> Back`
- `DocumentList -> DocumentDetail -> Back`
- `ArticleList -> ArticleDetail -> Back`

The goal is that users continue working without losing context, search position, or list selection.

### Detail Form Action Convention

For workspace-hosted detail forms, the standard action convention is:

- `Save`
  - save the record
  - return to the previous workspace or list
- `Cancel`
  - discard unsaved changes
  - return to the previous workspace or list
- `Apply` / `Uebernehmen`
  - reserved for future complex forms
  - save and keep the detail form open

Current rule:

- do not add `Apply` by default
- use `modAppWorkspaceService.GoBack(...)` for the return flow
- if no workspace history exists, the detail form may stay open safely

### Future Address Direction

`frmAddressDetail` is expected to evolve later into a customer account or action board.

Planned direction only:

- customer activity overview
- contextual actions such as:
  - new order
  - documents
  - invoices
  - subscriptions
  - communication
  - history

This is a future UX direction only and is not part of the current implementation pass.

## Navigation

Navigation conventions:

- `frmAppNavigation` is hosted in `subNavigationHost`
- accordion navigation is used
- only one group should be open at a time
- group clicks collapse other groups
- `GROUP` rows toggle expand and collapse
- `FORM` and `REPORT` rows open workspace content
- navigation open behavior can be driven by `fw_navigation.open_mode`
- `open_mode=ADD` is the preferred shell-safe way to open "New ..." workspace forms
- grouping is data-driven through `fw_navigation`
- no Access report-style grouping is used
- standard master-data object creation should come from `cmdNew` on the list form instead of separate shell navigation rows

Navigation maintenance conventions:

- use `frmFwNavigationAdmin` for controlled maintenance of `fw_navigation`
- do not manually delete seeded rows
- prefer `is_active` and `is_visible` flags over destructive cleanup

### Shell Command Bar

The experimental `sfrmFwListCommandBar` approach was replaced by a shell-owned command bar.

Owner:

- `frmAppShell`

Official layout:

- `[Home] [Zurueck] | [Suchen: __________] [Leeren] [Neu] [Bearbeiten] [Aktualisieren]`

Purpose:

- standardize list-local actions
- reduce duplicated search and button implementations
- keep shell navigation lean

List-form contract:

- `Public Function SupportsListCommandBar() As Boolean`
- `Public Sub ListSearch(ByVal searchText As String)`
- `Public Sub ListClearSearch()`
- `Public Sub ListNew()`
- `Public Sub ListEdit()`
- `Public Sub ListRefresh()`

Optional per-action support:

- `Public Function SupportsListNew() As Boolean`
- `Public Function SupportsListEdit() As Boolean`

Current migrated list forms:

- `frmArticleGroupList`
- `frmAddressList`
- `frmArticleList`
- `frmFwTranslationAudit`

## List Form Architecture

List forms are now built around a shell-owned command surface rather than local duplicated search controls.

Shell ownership:

- `frmAppShell`
  - owns `Home`
  - owns `Zurueck`
  - owns `Suche`
  - owns `Leeren`
  - owns `Neu`
  - owns `Bearbeiten`
  - owns `Aktualisieren`

List-form contract:

- `SupportsListCommandBar()`
- `ListSearch(ByVal searchText As String)`
- `ListClearSearch()`
- `ListNew()`
- `ListEdit()`
- `ListRefresh()`

Optional capability flags:

- `SupportsListNew()`
- `SupportsListEdit()`

Search architecture:

- list forms keep `Private m_currentSearchText As String`
- the shell command bar supplies the current search text
- local search textboxes are no longer required for migrated list forms
- filters should operate on one dedicated search-index field per row source

Standard search-index fields:

- `address_search_text`
- `article_group_search_text`
- `article_search_text`

Current note:

- `frmAddressList` still exposes legacy `AddressSearchText`
- this should be normalized to `address_search_text` in a later low-risk cleanup pass
- `frmFwTranslationAudit`
  - uses shell search for free-text filtering
  - keeps local dropdown filters for:
    - `scope`
    - `language`
    - `audit status`
  - keeps local domain action:
    - `Fehlende Eintraege erzeugen`

Workspace interaction:

- list forms store search and current-row state in `GetWorkspaceState()`
- detail forms open through `modAppWorkspaceService.OpenWorkspaceForm(...)`
- detail save and cancel return through workspace history

Navigation philosophy:

- navigation contains areas and worklists only
- standard create actions belong to `Neu` in the shell command bar
- standard master-data objects should not add separate `Neue ...` navigation entries

### Lean Navigation Rule

Shell navigation should remain lean:

- navigation represents areas and worklists
- standard object creation belongs to the shell command bar
- standard master-data objects should not need separate `Neue ...` shell entries

Navigation is intentionally lightweight and Access-safe.

This keeps navigation behavior simple, predictable, and easy to extend through data.

## Toolbar

Toolbar conventions:

- toolbar should live in the detail section, not the form header
- preferred host name: `subToolbarHost`
- preferred form name: `frmAppToolbar`
- shell may contain `cmdBack`
- initial implementation may be static
- future implementation may become context-sensitive
- workspace-aware toolbar integration is planned
- hybrid toolbar/dropdown action models are preferred over ribbon-style complexity

Current lightweight examples already fit this direction:

- `New`
- `Edit`
- `Refresh`

Keeping the toolbar in the detail section gives more control over layout and embedded shell composition.

## Dialog Forms

Dialogs should be modal standalone forms and should not be embedded in the workspace.

Recommended small dialog size:

- width: `12-16 cm`
- height: `6-10 cm`

Recommended medium dialog size:

- width: `16-22 cm`
- height: `10-14 cm`

Examples:

- selection dialogs
- confirmation dialogs
- lookup disambiguation dialogs

Dialogs should not be used as a workaround for normal workspace navigation.

## Workspace Form Design Standards

Workspace forms should be designed and coded as embedded forms first.

Recommended:

- workspace forms behave as embedded forms
- keep layouts stable and simple
- avoid focus-heavy hacks
- avoid unnecessary popup forms
- avoid unsupported transparency tricks
- avoid excessive conditional formatting

Important Access behavior:

Workspace forms may lose focus during shell navigation and history operations.

Code should therefore avoid using:

```vb
control.Text
```

outside active control focus scenarios.

Prefer:

```vb
control.Value
```

whenever possible.

## Focus Management Notes

Practical Access-specific focus issues still matter for migrated list forms:

- filtering or requerying while a bound control has focus may raise `Err 2185`
- workspace navigation can trigger focus transitions
- embedded forms behave differently from standalone forms

Recommended pattern:

- use hidden or lightweight focus sink controls when necessary
- move focus before filter or requery operations
- avoid `.Text` where possible
- keep search state in form variables instead of local search controls

## Design Rules

Recommended UI conventions:

- use `Segoe UI`
- prefer a bright modern layout
- avoid dense button grids
- prefer simple flat controls
- avoid too much conditional formatting in continuous forms
- use transparent click buttons carefully
- run Compact and Repair after heavy UI layout changes
- commit after stable UI milestones

These conventions are intended to keep the Access frontend clean, maintainable, and consistent.

## UI Stability Recommendations

Operational recommendations learned during implementation:

- run Compact and Repair regularly during UI-heavy development
- avoid excessive redesign inside continuous forms
- save and reopen forms after major layout changes
- commit after stable UI milestones

Access form designer instability is a practical architectural consideration and should be planned for during implementation work.

## Implementation Note

Existing Access form layout may need manual setup in Access because form geometry is not always reliably represented in exported text files.

This is especially relevant for:

- section heights
- exact control coordinates
- stacking order
- visual spacing
- embedded subform placement

Behavioral architecture belongs in modules and services, not in duplicated form event logic where avoidable.
