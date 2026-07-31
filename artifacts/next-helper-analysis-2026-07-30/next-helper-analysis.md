# Next Helper Analysis

Date: 2026-07-30

## Counts
- HasControl: 4 definition(s), 23 actual call site(s)
- ResolveCreatedBy: 3 definition(s), 3 actual call site(s)
- ResolveFieldValue: 2 definition(s), 6 actual call site(s)

## HasControl
- Definitions: 4
- Actual call sites: 23
- Modules: modAppDashboardService, modAppShell, modAppWorkspaceService, modFwTranslationEditService
- Global groups: one exact implementation family reused in 4 modules.
- Functional differences: none found in behavior; only parameter name differs (`formInstance` vs `targetForm`).
- Hidden dependencies: hard dependency on `Access.Form` and `.Controls`; no report or generic object support.
- Possible canonical signature: `Private Function HasControl(ByVal formInstance As Access.Form, ByVal ControlName As String) As Boolean`.
- Suggested target module: none yet; keep local to each UI/service module.
- Public/Private recommendation: Private.
- Better inline?: No. The helper is still more readable than repeating the loop, but not worth a shared public API.
- Risks: centralizing into shell/form infrastructure would spread a UI helper across unrelated modules for little gain.
- Reduced-to-the-Max answer: No. It would be shorter, but not simultaneously clearer and more maintainable.

## ResolveCreatedBy
- Definitions: 3
- Actual call sites: 3
- Modules: modAddressRepository, modContactRepository, modDocumentRepository
- Global groups: one exact implementation family reused in 3 modules.
- Functional differences: none found.
- Hidden dependencies: `IsSessionInitialized()` and `currentUserId` ambient session state.
- Possible canonical signature: `Public Function ResolveCreatedBy() As String`.
- Suggested target module: existing session/user-context owner.
- Public/Private recommendation: Public candidate if moved to that owning module.
- Better inline?: No. Repeating the session fallback rule is worse than one shared implementation.
- Risks: naming must remain scoped to audit/created-by semantics, not generic current-user display semantics.
- Reduced-to-the-Max answer: Yes. One central implementation would be shorter, clearer and more maintainable.

## ResolveFieldValue
- Definitions: 2
- Actual call sites: 6
- Modules: modAddressRepository, modContactRepository
- Global groups: one exact implementation family reused in 2 modules.
- Functional differences: none found.
- Hidden dependencies: DAO recordset input, `modDaoHelper.RecordsetHasField`, and `modDaoHelper.NzString`.
- Possible canonical signature: `Public Function ResolveFieldValue(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal defaultValue As String) As String`.
- Suggested target module: `modDaoHelper`.
- Public/Private recommendation: Public candidate in DAO helper layer.
- Better inline?: No. The helper captures a repeated low-level DAO rule succinctly.
- Risks: keep it read-only; do not blend with mutating helpers such as `SetRecordsetValue`.
- Reduced-to-the-Max answer: Yes. One central DAO helper would be shorter, clearer and more maintainable.

## Recommended Implementation Order
1. ResolveCreatedBy
2. ResolveFieldValue
3. Leave HasControl private unless the UI infrastructure gains a clearly responsible shared form-runtime module.
