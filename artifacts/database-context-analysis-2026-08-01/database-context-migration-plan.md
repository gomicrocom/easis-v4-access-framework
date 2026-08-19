# Database Context Migration Plan

## Scope

This is **step 1 analysis only**. No source-module migration is performed in this step.

## Current findings

### Existing database contexts

1. Frontend database
   - Source today: `CurrentDb` / `currentDb`
   - Typical local tables: `tmp_order`, `tmp_order_line`, `tmp_fw_translation_audit`, `tmp_fw_translation_tag_generator`, `fw_tmp_tag_composer`

2. System backend
   - Strong evidence: `sys_be.accdb`
   - Representative tables: `fw_translation`, `ref_*`, framework/reference tables such as `fw_navigation`, `fw_tag_help`, `fw_list_action`

3. Tenant backend
   - Runtime source today: `modTenantContext.CurrentTenantBackendPath`
   - Representative tables: `ten_*`, `adr_*`, `ord_*`, `doc_*`

### Existing resolver ownership

- `modDb`
  - already owns `GetCurrentDatabase`, `GetBackendPath`, `ValidateBackendConfiguration`
  - best candidate for the canonical public API

- `modTenantContext`
  - already owns active tenant code and tenant backend path construction
  - should remain the source of truth for the active tenant backend path

- `modConfigIni`
  - already owns INI path/config lookup
  - likely source of the system backend path if not already centralized elsewhere

- `modBasicModuleSchema`
  - currently has a private schema resolver (`OpenSchemaDatabase`) with mixed responsibilities
  - contains fallback behavior to frontend that is not acceptable for a strict backend-DDL architecture

- `modMigrationPaymentTerms`
  - currently has its own tenant-backend resolution and relink helpers
  - this is confirmed duplication and should be removed in step 2

## Proposed canonical public API

Preferred target module: `modDb`

### Recommended public functions

```vba
Public Function GetFrontendDatabase() As DAO.Database
Public Function GetSystemDatabase() As DAO.Database
Public Function GetCurrentTenantDatabase() As DAO.Database
```

### Supporting path functions

If existing naming can be clarified without duplicate public APIs, also prefer:

```vba
Public Function GetSystemBackendPath() As String
Public Function GetCurrentTenantBackendPath() As String
```

`GetCurrentDatabase()` should then either:

- remain as a backward-compatible frontend alias and be documented as such, or
- be retired later after call-site migration.

## Ownership and lifetime

Recommended rule:

- `GetFrontendDatabase()` returns `CurrentDb`; caller sets local variable to `Nothing`.
- `GetSystemDatabase()` opens the physical system backend; caller is responsible for closing if needed and releasing the DAO object.
- `GetCurrentTenantDatabase()` opens the physical tenant backend; caller is responsible for closing if needed and releasing the DAO object.

No resolver may silently fall back from system/tenant backend to `CurrentDb`.

## Concrete step-2 migration order

1. Implement / consolidate canonical API in `modDb`.
2. Add explicit system backend path resolution using existing config/bootstrap infrastructure.
3. Add explicit tenant backend DAO opener using `modTenantContext.CurrentTenantBackendPath`.
4. Migrate duplicated path readers:
   - `modMigrationPaymentTerms.ResolveBusinessBackendPath`
   - `modMigrationPaymentTerms.GetBackendPathForLinkedTable`
   - `modBasicModuleSchema.OpenSchemaDatabase`
5. Migrate highest-risk call sites first:
   - `modTranslationService`
   - `modFwTranslationRuntime`
   - `modFwSetup` reference/translation seeding
   - `modOrderRepository`
   - `modDocumentRepository`
   - `modAddressRepository`
6. Split mixed-context routines:
   - especially `modFwSetup.NormalizeLanguageCodeData`
7. Revisit relinking:
   - tenant relink should not blindly imply system tables share the same backend
   - system and tenant relink/refresh responsibilities should be separated explicitly
8. Remove redundant private resolvers after callers are migrated.

## Payment-term test case

`modMigrationPaymentTerms.ApplyPaymentTermsMigration` is the strongest existing proof-of-concept:

- it now resolves `ten_payment_term` from FE linked-table metadata
- it opens the physical tenant backend with `DBEngine.OpenDatabase`
- logging shows DDL target DB path explicitly

For step 2, this should be simplified to:

- `Set backendDb = modDb.GetCurrentTenantDatabase()`
- pass `backendDb` to all schema/data routines
- remove domain-local backend resolution and relink duplication

## Risks

1. `modDb.GetCurrentDatabase()` is currently semantically misleading because many callers use it for backend tables while it only returns FE `CurrentDb`.
2. `modBasicModuleSchema.OpenSchemaDatabase` still has a frontend fallback when no path is resolved.
3. `modFwSetup` mixes tenant and system tables in single routines, making one-shot migration unsafe without splitting responsibilities.
4. `modBackendLinker` currently appears tenant-backend centric; system-linked tables may need a separate strategy.
5. Existing code often relies on linked-table transparency, which is unsafe for DDL and ambiguous for architecture.

## Test plan for step 2

1. `GetFrontendDatabase()` returns the FE file path.
2. `GetSystemDatabase()` opens physical `sys_be.accdb`.
3. `GetCurrentTenantDatabase()` opens the active tenant backend, e.g. `DEFAULT_be.accdb`.
4. Missing/invalid system backend path returns failure or `Nothing`, never FE fallback.
5. Missing/invalid tenant backend path returns failure or `Nothing`, never FE fallback.
6. Payment-term migration uses only `GetCurrentTenantDatabase()` for tenant DDL/data changes.
7. Translation runtime / setup use only `GetSystemDatabase()` for `fw_translation` and `ref_*`.
8. Tenant switch updates resolved tenant DB path without stale cache.

## Documentation to update in step 2

- architecture documentation
- backend/bootstrap documentation
- developer/contributing documentation
- data model notes for frontend vs system vs tenant table placement
