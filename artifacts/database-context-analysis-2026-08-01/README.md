# Database Context Analysis

Date: 2026-08-01

This folder contains **step 1 analysis only** for central frontend/system/tenant database resolution.

## Included files

- `database-context-inventory.csv`
- `database-resolver-inventory.csv`
- `database-open-call-sites.csv`
- `database-context-migration-plan.md`

## Summary

### Existing resolvers found

- `modDb.GetCurrentDatabase`
- `modDb.GetBackendPath`
- `modDb.ValidateBackendConfiguration`
- `modTenantContext.InitializeTenantContext`
- `modTenantContext.CurrentTenantBackendPath`
- `modBasicModuleSchema.OpenSchemaDatabase` (private, duplicated resolver logic)
- `modMigrationPaymentTerms.ResolveBusinessBackendPath` (private, duplicated resolver logic)
- `modMigrationPaymentTerms.GetBackendPathForLinkedTable` (private, duplicated resolver logic)
- `modBackendLinker.RelinkBackendTables`

### Direct physical database open calls found

- `modBasicModuleSchema.OpenSchemaDatabase`
- `modBasicModuleSchema.OpenOrCreateAccessDatabase`
- `modMigrationPaymentTerms.ApplyPaymentTermsMigration`

### Recognized database contexts

- Frontend database
- System backend
- Tenant backend

### Main architectural conclusion

The project already has enough infrastructure to centralize the public API in `modDb`, but the current implementation is incomplete:

- frontend access is centralized,
- tenant backend path resolution exists,
- system backend resolution is not yet exposed through a matching canonical public DAO API,
- several modules still contain local backend-resolution logic,
- many repositories and setup routines still use `CurrentDb` or `modDb.GetCurrentDatabase()` even when the physical table belongs to a backend file.

### Most important migration candidates

- `modFwSetup`
- `modTranslationService`
- `modFwTranslationRuntime`
- `modBasicModuleSchema`
- `modMigrationPaymentTerms`
- `modOrderRepository`
- `modDocumentRepository`
- `modAddressRepository`
- `modContactRepository`

### Key risk

`modDb.GetCurrentDatabase()` currently returns FE `CurrentDb`, but many callers use it as if it were the correct physical data context for backend-owned tables. This is the primary ambiguity to remove in step 2.

### No code changes in this step

No VBA source modules were intentionally changed for this analysis task. Only artifact files in this folder were added.
