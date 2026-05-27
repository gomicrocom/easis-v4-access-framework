# Application Shell

## Purpose

`frmAppShell` is the first lightweight application host for the Easis v4 Access frontend.

It keeps one stable shell form open and separates the UI into:

- header
- left navigation
- right workspace
- bottom status bar

The shell is intentionally thin. Navigation logic, workspace loading, and dashboard status population live in service modules.

## Main Objects

### Forms

- `frmAppShell`
  - host form
- `frmAppDashboard`
  - default workspace landing page built as a normal unbound Access form
- `frmAppNavigation`
  - navigation menu loaded into the shell

### Modules

- `modAppShell`
  - shell startup orchestration
- `modAppNavigationService`
  - table setup, seed logic, navigation query, navigation clicks
- `modAppWorkspaceService`
  - workspace form loading and report preview
- `modAppDashboardService`
  - dashboard card population

## Navigation Architecture

Navigation data is stored in frontend table `fw_navigation`.

Important fields:

- `navigation_id`
- `parent_navigation_id`
- `navigation_group`
- `caption_key`
- `fallback_caption`
- `object_name`
- `object_type`
- `open_mode`
- `icon_key`
- `sort_order`
- `is_active`
- `is_expanded`
- `is_visible`

Role assignments are prepared in `fw_navigation_role`.

Navigation is now hierarchical:

- `GROUP` rows are top-level accordion entries
- `FORM` and `REPORT` rows are child entries
- child rows point to their parent through `parent_navigation_id`
- only children of expanded groups are shown

The current version keeps role filtering lightweight:

- if no role rows exist for a navigation entry, it is visible
- if role rows exist, `GetNavigationRowSource()` can filter by `role_code`

### Open Mode Metadata

Navigation open behavior is metadata-driven through `fw_navigation.open_mode`.

Allowed values:

- `NORMAL`
- `ADD`
- `EDIT`
- `READONLY`

Current behavior:

- `NORMAL`
  - opens a form with standard workspace behavior
- `ADD`
  - opens a form in add or new-record mode through `modAppWorkspaceService`
- `EDIT`
  - currently falls back to `NORMAL`
- `READONLY`
  - currently falls back to `NORMAL`

This avoids hardcoded special cases such as tying new-record behavior to a specific `caption_key` or form name.

### Accordion Behavior

Displayed structure follows this pattern:

- expanded group: `- Adressen`
- collapsed group: `+ Dokumente`
- child row: `    Adressliste`

`is_expanded` controls whether child rows are visible.
`is_visible` controls whether a row participates in navigation output at all.

## Workspace Loading

Workspace loading is handled by `modAppWorkspaceService`.

Preferred pattern:

- forms are embedded through shell subform control `subWorkspaceHost`
- reports are opened in preview mode

Current public entry points:

- `OpenWorkspaceForm()`
- `PushWorkspaceState()`
- `GoBack()`
- `CanGoBack()`
- `PreviewWorkspaceReport()`
- `LoadDashboard()`
- `ClearWorkspace()`

### Workspace History

The shell now supports an in-memory workspace history stack.

Each history item can capture:

- `form_name`
- `where_condition`
- `filter`
- `order_by`
- `current_record_id`
- `workspace_state`
- `open_args`

History is not persisted to a table in this phase.

### Form State Contract

Workspace-aware forms may optionally expose:

- `Public Function GetWorkspaceState() As String`
- `Public Sub RestoreWorkspaceState(ByVal stateText As String)`
- `Public Function CanLeaveWorkspace() As Boolean`

If these members are available, the workspace service will use them during back navigation.

If they are not available:

- navigation still works
- basic reopen and generic filter/order restore still work where possible
- no hard failure should occur

## Status Bar Fields

`modAppShell.RefreshShellStatus()` fills these shell values when matching controls exist:

- app version
- current user
- current tenant
- current role
- backend status
- environment

Expected shell controls:

- `txtStatusAppVersion`
- `txtStatusCurrentUser`
- `txtStatusCurrentTenant`
- `txtStatusCurrentRole`
- `txtStatusBackend`
- `txtStatusEnvironment`

## Dashboard Fields

`frmAppDashboard` should use ordinary Access controls only:

- rectangle or box controls for visual card backgrounds
- labels for card titles
- unbound textboxes for card values

There is no custom or native MS Access `CardControl`.

`modAppDashboardService.RefreshDashboard()` fills these value controls when they exist:

- `txtCardTenant`
- `txtCardUser`
- `txtCardBackend`
- `txtCardFramework`

Optional title labels:

- `lblCardTenantTitle`
- `lblCardUserTitle`
- `lblCardBackendTitle`
- `lblCardFrameworkTitle`

Optional background boxes:

- `boxCardTenant`
- `boxCardUser`
- `boxCardBackend`
- `boxCardFramework`

Fallback text is always safe, usually `n/a`.

## Adding Navigation Entries

Recommended approach:

1. ensure `fw_navigation` exists with `EnsureNavigationTables()`
2. create or update a `GROUP` row first
3. create or update child rows afterwards with `parent_navigation_id`
3. use a stable `caption_key`
4. keep `fallback_caption` readable even if no translation exists
5. use supported `object_type` values:
   - `FORM`
   - `REPORT`
   - `ACTION`
   - `GROUP`

For this phase, navigation display is translation-aware through `caption_key`, but it still remains fallback-first through `fallback_caption` so the shell compiles and runs safely even when translations are missing.

### Navigation Maintenance

Framework navigation can be maintained through:

- `frmFwNavigationAdmin`

Default placement:

- `System`
  - `Navigation verwalten`

Maintenance rules:

- do not physically delete seeded navigation rows
- use `is_active=False` to deactivate entries
- use `is_visible=False` to hide entries
- setup may create missing default rows and update structural fields
- setup should not reactivate or re-show rows that were manually disabled or hidden

### Navigation Click Flow

- clicking a `GROUP` row toggles `is_expanded`
- the navigation form requeries itself
- clicking a `FORM` row calls `modAppWorkspaceService.OpenWorkspaceForm`
- `open_mode=ADD` opens a blank new record through the same workspace service
- clicking a `REPORT` row calls `modAppWorkspaceService.PreviewWorkspaceReport`
- `ACTION` rows are only logged for now

### Workspace Loading Flow

- `frmAppShell` hosts left navigation in `subNavigationHost`
- `frmAppShell` hosts workspace content in `subWorkspaceHost`
- forms are loaded into `subWorkspaceHost` through `SourceObject`
- reports open in preview mode

## Shell Translation Strategy

The shell layer is translation-aware, but intentionally conservative.

Current strategy:

- shell labels use translation keys where available
- navigation captions come from `fw_navigation.caption_key`
- fallback always uses `fallback_caption`
- dashboard card titles use translation keys
- dynamic values remain technical and are not fully translated

The shell uses a safe wrapper approach:

- `ResolveShellText(translation_key, fallback_text)`

Behavior:

- return translated text if available
- return fallback text if translation is missing
- never break shell loading because of a missing translation

### Navigation Caption Source

Navigation captions are resolved from:

- `caption_key`
- `fallback_caption`

This means a navigation entry can remain usable even if:

- the current language has no translation yet
- translation seeding has not been run

### Translation Key Naming

Current shell-related naming convention:

- `NAV.*`
- `FORM.<FORMNAME>.*`
- `STATUS.*`

For the broader framework translation policy, namespace rules, and cleanup phases, see:

- [TRANSLATION_RULES.md](./TRANSLATION_RULES.md)

Examples:

- `NAV.GROUP.ADDRESSES`
- `NAV.ADDRESS_LIST`
- `FORM.FRMAPPSHELL.USER`
- `FORM.FRMAPPSHELL.TENANT`
- `FORM.FRMAPPDASHBOARD.BACKEND`

### Form Translation Marker Rule

For shell forms and other translatable UI controls, the official rule is:

- `Caption`
  - contains the readable fallback or design text
- `Tag`
  - contains the translation marker in the form `TR:<translation_key>`

Example:

- `Caption = Access Framework`
- `Tag = TR:FORM.FRMAPPSHELL.APP_SUBTITLE`

Important:

- `fw_translation.translation_key` stores only the pure key
- the `TR:` prefix is a runtime and designer marker only
- new translation-tag writes should update `Tag`, not `Caption`
- legacy `Caption` values that begin with `TR:` are tolerated for backward compatibility, but they are no longer the target pattern

### Adding Translated Navigation Entries

When adding a new shell navigation entry:

1. define a stable `caption_key`
2. keep a readable `fallback_caption`
3. set `open_mode` explicitly when behavior differs from the default
4. add translation rows for at least the supported shell languages
5. do not rely on translation availability for functional navigation

### Workspace-Safe Form Navigation

Shell-aware workspace forms should prefer `modAppWorkspaceService.OpenWorkspaceForm()` instead of direct `DoCmd.OpenForm` calls when navigating to another form.

Recommended pattern:

- keep search, validation, and business actions inside the form or service modules
- route form-to-form navigation through the workspace service
- pass a `where_condition` when opening an existing record
- use add mode only when a genuine new-record workflow is intended
- rely on workspace history instead of reopening previous lists manually
- keep save and cancel behavior inside the workspace form instead of closing forms aggressively

This helps preserve a single-shell workflow:

- no unnecessary floating forms
- no stacked list/detail windows
- reusable workspace host behavior across modules

### Back Navigation Flow

- before a workspace form is replaced, the current workspace state is captured
- opening a detail form from a list form pushes the list state onto the history stack
- `GoBack()` restores the previous form into `subWorkspaceHost`
- if the previous form supports `RestoreWorkspaceState(...)`, its own search and selection context can be restored

Example:

- `frmAddressList`
- search for `meier`
- open `frmAddressDetail`
- execute `GoBack()`
- return to `frmAddressList` with restored list context when supported by the form

For shell-aware detail forms, `CanLeaveWorkspace()` can be used to prevent accidental loss of unsaved changes during shell navigation.

## Manual Layout Notes

### frmAppNavigation

Recommended control setup:

- continuous form
- `txtDisplayCaption`
  - bound to `display_text`
  - wide enough for group and child rows
- `txtNavigationId`
  - bound to `navigation_id`
  - hidden
- `txtObjectType`
  - bound to `object_type`
  - hidden
- optional:
  - `txtDisplayLevel`
  - `txtIsExpanded`
  - `cmdOpen`

Suggested visual treatment:

- no report-style grouping
- group rows may use bold conditional formatting
- child rows rely on text indentation from `display_text`
- keep record selectors and navigation buttons off where practical

## Role-Ready Design

The first shell version is already prepared for role-specific navigation.

Current design choices:

- `fw_navigation_role` stores optional role assignments
- navigation filtering accepts optional `role_code`
- no business permissions are hard-coded in the form classes

This keeps the shell extensible for later module/role policies without rewriting the UI host.

## Startup Integration

The repository already has bootstrap logic, but it does not yet force a startup form.

Current recommendation:

- keep startup settings manual in Access for now
- set `frmAppShell` as startup form when the layout is created and verified

This avoids changing production startup behavior prematurely.

### Optional Shell Back Button

If shell layout includes a command button named `cmdBack`:

- it can call `modAppWorkspaceService.GoBack(Me)`
- it should be enabled only when `modAppWorkspaceService.CanGoBack()` returns `True`
