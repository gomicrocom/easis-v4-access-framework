# frmAddressList

## Purpose

`frmAddressList` is the planned address overview and navigation form for BasicModule v1.

It is intended to provide:

- fast address lookup
- compact continuous-list browsing
- a single selected-address workflow
- dynamic navigation into related business areas

The form should become the primary starting point for address-centric work such as order creation, account review, communication history, and reporting.

## Design Goals

The updated concept replaces row-level action buttons with a single dynamic action menu for the currently selected address.

This keeps the list compact and avoids:

- crowded row layouts
- duplicated button logic per record
- repeated future maintenance when new actions are introduced

The form therefore separates responsibilities into:

- shell-driven search and selection
- address list display
- selected-address action execution

## Search Model

The address list should expose one technical search field:

- preferred standard
  - `address_search_text`
- current legacy alias
  - `AddressSearchText`

This field is intended for broad free-text filtering and should concatenate relevant address data without spaces.

Required expression:

```text
AddressSearchText =
Nz([CompanyName],"") &
Nz([FirstName],"") &
Nz([LastName],"") &
Nz([PostalCode],"") &
Nz([City],"") &
Nz([Email],"") &
Nz([Phone],"")
```

Important rule:

- no inserted separator spaces inside the concatenation

Rationale:

- search terms should remain compact
- partial matching should work consistently across combined values
- the expression stays easy to extend later

## Future Search Extension

`AddressSearchText` is intentionally a technical aggregation field and should be treated as extensible.

Later additions may include:

- invoice numbers
- invoice amounts
- order numbers
- subscription references
- customer account numbers

The design should therefore avoid hard-coding the search concept to only company/person fields.

Recommended direction:

- keep one expandable search source only
- allow future query/service logic to append business-context fields without changing the form concept

## UI Structure

The preferred layout is:

### Header Area

- shell-owned command bar search
- optional result summary
  - count of matching addresses
- action menu host
  - side or header placement

### Main Area

- continuous form list of addresses
- one selected address at a time
- no row-level action button cluster

### Action Area

- one dynamic action menu bound to the selected address
- actions displayed in configured order
- execution always targets the currently selected `AddressId`

## Address List Behavior

The address list should stay visually simple and optimized for scanning.

Recommended visible fields:

- company name
- first name / last name
- postal code
- city
- email
- phone

Recommended hidden technical fields:

- `AddressId`
- `AddressSearchText`

The current record selection should drive:

- action availability
- target navigation
- placeholder execution messaging

## Dynamic Action Menu

The action model should be configuration-driven instead of button-driven.

Initial actions:

- `EDIT`
- `ORDERS`
- `NEW_ORDER`
- `SUBSCRIPTIONS`
- `EMAILS`
- `ACCOUNT`
- `REPORTS`

Behavior:

- actions are loaded from a table if practical
- only active actions are shown
- the selected action executes against the current address
- missing target forms show a placeholder `MsgBox` for now

This keeps the form open for future extension without redesigning every list row.

## Config Table

Preferred table:

- `fw_list_action`

Fields:

| Field | Type | Purpose |
|---|---|---|
| `action_id` | `AUTOINCREMENT` | technical primary key |
| `list_code` | `TEXT(50)` | list runtime scope such as `ADDRESS` |
| `action_code` | `TEXT(50)` | stable action identifier |
| `action_label` | `TEXT(100)` | UI label shown in the menu |
| `target_form` | `TEXT(100)` | navigation target |
| `requires_selection` | `YESNO` | indicates whether a selected address is required |
| `module_code` | `TEXT(50)` | optional module scoping |
| `role_code` | `TEXT(50)` | optional role scoping |
| `sort_order` | `LONG` | display order |
| `is_active` | `YESNO` | activation flag |
| `created_at` | `DATETIME` | creation timestamp |
| `created_by` | `TEXT(100)` | creator identity |
| `updated_at` | `DATETIME` | last update timestamp |
| `updated_by` | `TEXT(100)` | last updater identity |

Usage notes:

- `action_code` should remain stable
- `action_label` may later become translation-driven
- `target_form` may remain empty for placeholder or service-driven actions
- `requires_selection = False` can later support generic actions such as global reports or address creation

## Preferred Runtime Flow

### Form Load

- load active actions from `fw_list_action`
- default to current sort order
- load address list

### Search Change

- apply filter using the shell command bar search text against the technical search field
- keep current row selection stable where possible

### Address Selection Change

- update current context
- enable or disable address-dependent actions

### Action Execution

- read selected action from the dynamic menu
- resolve current `AddressId`
- open target form if available
- otherwise show placeholder message

## Placeholder Handling

Until all target forms exist, the concept explicitly allows placeholder behavior.

Expected placeholder pattern:

- if `target_form` is empty or does not exist
- show `MsgBox` indicating the selected action and target are not implemented yet

This allows the action architecture to be built before all downstream forms are ready.

## Rationale for Removing Row Buttons

The former row-button approach does not scale well once multiple business areas are connected to addresses.

Problems avoided by the new design:

- too many repeated controls in continuous forms
- inconsistent button placement and widths
- difficult extension when new business actions appear
- noisy visual hierarchy

The action menu approach is preferred because it:

- keeps the list readable
- centralizes navigation behavior
- makes action setup configurable
- supports future module growth

## Recommended Next Technical Steps

- create `fw_list_action`
- seed initial actions
- define the address list query with `AddressSearchText`
- build `frmAddressList` as a continuous form
- add a lightweight action execution service
- route missing forms to placeholder handling
