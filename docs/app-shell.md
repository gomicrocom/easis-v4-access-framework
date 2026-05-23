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

For this phase, navigation display is intentionally based on `fallback_caption` so the shell compiles and runs even without any translation dependency.

### Navigation Click Flow

- clicking a `GROUP` row toggles `is_expanded`
- the navigation form requeries itself
- clicking a `FORM` row calls `modAppWorkspaceService.OpenWorkspaceForm`
- clicking a `REPORT` row calls `modAppWorkspaceService.PreviewWorkspaceReport`
- `ACTION` rows are only logged for now

### Workspace Loading Flow

- `frmAppShell` hosts left navigation in `subNavigationHost`
- `frmAppShell` hosts workspace content in `subWorkspaceHost`
- forms are loaded into `subWorkspaceHost` through `SourceObject`
- reports open in preview mode

### Workspace-Safe Form Navigation

Shell-aware workspace forms should prefer `modAppWorkspaceService.OpenWorkspaceForm()` instead of direct `DoCmd.OpenForm` calls when navigating to another form.

Recommended pattern:

- keep search, validation, and business actions inside the form or service modules
- route form-to-form navigation through the workspace service
- pass a `where_condition` when opening an existing record
- use add mode only when a genuine new-record workflow is intended
- rely on workspace history instead of reopening previous lists manually

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
