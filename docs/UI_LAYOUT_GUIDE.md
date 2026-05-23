# UI Layout Guide

## Goal

The Easis v4 Access Framework uses a fixed shell and workspace layout so future forms can be designed consistently.

This keeps the frontend predictable across modules:

- the shell remains stable
- navigation stays in one place
- workspace forms share one standard canvas
- status information stays compact and visible

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

## Navigation

Navigation conventions:

- `frmAppNavigation` is hosted in `subNavigationHost`
- accordion navigation is used
- only one group should be open at a time
- `GROUP` rows toggle expand and collapse
- `FORM` and `REPORT` rows open workspace content
- grouping is data-driven through `fw_navigation`
- no Access report-style grouping is used

This keeps navigation behavior simple, predictable, and easy to extend through data.

## Toolbar

Toolbar conventions:

- toolbar should live in the detail section, not the form header
- preferred host name: `subToolbarHost`
- preferred form name: `frmAppToolbar`
- initial implementation may be static
- future implementation may become context-sensitive

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

## Implementation Note

Existing Access form layout may need manual setup in Access because form geometry is not always reliably represented in exported text files.

This is especially relevant for:

- section heights
- exact control coordinates
- stacking order
- visual spacing
- embedded subform placement
