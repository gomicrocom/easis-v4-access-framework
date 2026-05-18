# ADR-0001: Separation of Local Configuration, Tenant Data, System Data Pools and Logs

Status: Accepted

## Context

Easis v4 uses an Access frontend with tenant-specific backend databases. In addition to tenant business data, the platform also requires:

- local installation-specific configuration
- system-wide reference and framework data pools
- technical and operational logs

These data categories have different lifecycles, ownership, update paths, backup rules, and operational risks. They must therefore not be stored together in one database or mixed across the same storage area.

Typical lifecycle differences are:

- local configuration belongs to a specific frontend installation and workstation context
- tenant data belongs to one tenant and must remain isolated
- system reference data must be centrally replaceable and updateable
- logs grow independently and must be archivable, purgeable, or rotatable without affecting business data

Without a clear separation, updates become harder, support becomes riskier, and data boundaries become unclear.

## Decision

Data storage is separated into four areas.

### Naming Convention

The following naming convention is mandatory across the backend structure:

| Prefix | Scope | Location |
|---|---|---|
| `ref*` | system-wide reference data | `sys_be.accdb` |
| `sys*` | technical system tables if required later | `sys_be.accdb` |
| `log*` | log tables | `log_be.accdb` |
| `ten*` | tenant parameters | tenant backend |
| `doc*` | documents and document-related business data | tenant backend |
| `adr*` | addresses and address-related business data | tenant backend |

Important rule:

The former `tbl*` prefix is no longer used in the new architecture.

### 1. Local Configuration

Path:

```text
<AppPath>\Cfg\easis.ini
```

Purpose:

- local installation-specific settings
- paths to backend files
- `TenantCode`
- `Environment`
- local defaults

Not contained here:

- no tenant business data
- no logs
- no reference data pools

### 2. Tenant Backend

File:

```text
<AppPath>\Data\<TenantCode>_be.accdb
```

Purpose:

- tenant-specific business data
- addresses
- documents
- document positions
- tenant parameters
- `BASIC_DOC_PATH`
- `ADDRESS_WINDOW_POSITION`
- `DEFAULT_LANGUAGE`

### 3. System Backend / Data Pool

File:

```text
<AppPath>\Data\sys_be.accdb
```

Purpose:

- system-wide reference data
- centrally exchangeable data pools
- translation tables
- ISO-3166-1 countries
- ISO-3166-2 regions
- ISO-4217 currencies
- postal codes
- phone prefixes
- document type definitions
- VAT code definitions
- framework reference data

Important rule:

`sys_be.accdb` must remain updateable by replacing the file. Therefore it must not contain local settings, logs, or tenant business data.

Current system-wide reference tables include:

- `refCountries`
- `refCountryTimezones`
- `ref_postal_code`
- `refCurrencies`

### 4. Log Backend

File:

```text
<AppPath>\Data\log_be.accdb
```

Purpose:

- technical system logs
- framework logs
- error logs
- export, import, and runtime logs

Example fields:

- `log_timestamp`
- `log_level`
- `module_name`
- `procedure_name`
- `message`
- `err_number`
- `tenant_code`
- `user_name`
- `machine_name`
- `frontend_version`

Not contained here:

- no tenant business data
- no system-wide reference data

## Rationale

This separation is intentional because each data area serves a different architectural purpose and follows a different operational lifecycle.

Local configuration belongs to the frontend installation, not to the tenant backend and not to shared framework databases. This keeps workstation-specific settings independent from tenant data and avoids unnecessary coupling during rollout, support, or workstation replacement.

Tenant business data must remain isolated per tenant. This is important for ownership, backup and restore, troubleshooting, security boundaries, and controlled rollout of tenant-specific changes. Mixing tenant data with shared pools or technical logs would weaken that isolation.

System reference data must be replaceable through a controlled update of `sys_be.accdb`. That only works if the file contains purely shared reference content. Once local settings, logs, or tenant state are mixed into that database, simple file replacement is no longer safe.

Logs have their own lifecycle. They may grow continuously, require retention rules, need archival, or be cleared independently from business data. A dedicated `log_be.accdb` keeps operational concerns separate from transactional and reference data.

The chosen structure therefore creates clear boundaries between:

- frontend-local configuration
- tenant-local business data
- system-wide reference data pools
- operational logging

## Data Area Overview

| Data Area | Location | Contains | Does Not Contain |
|---|---|---|---|
| Local Configuration | `<AppPath>\Cfg\easis.ini` | local installation settings, backend paths, `TenantCode`, `Environment`, local defaults | tenant business data, logs, reference data pools |
| Tenant Backend | `<AppPath>\Data\<TenantCode>_be.accdb` | tenant business data, addresses, documents, document positions, tenant parameters such as `BASIC_DOC_PATH`, `ADDRESS_WINDOW_POSITION`, `DEFAULT_LANGUAGE` | system-wide reference data, technical logs, machine-local configuration |
| System Backend / Data Pool | `<AppPath>\Data\sys_be.accdb` | shared reference data, translations, countries, regions, currencies, postal codes, phone prefixes, document type definitions, VAT code definitions, framework reference data | tenant business data, logs, local settings |
| Log Backend | `<AppPath>\Data\log_be.accdb` | technical logs, framework logs, error logs, export/import/runtime logs | tenant business data, system-wide reference data, local configuration |

## Consequences

- clear separation of lifecycles and responsibilities
- `sys_be.accdb` can be updated by file replacement
- `log_be.accdb` can be rotated, archived, or cleared separately
- tenant backends remain cleanly tenant-isolated
- the frontend does not hold productive business data
- `easis.ini` remains local installation configuration
- startup and linking must evolve to connect tenant backend, system backend, and log backend
- tenant databases should store only reference keys such as country codes, not full copies of shared system data

## Future Work

- implement startup linking for `sys_be.accdb` and `log_be.accdb`
- migrate existing framework tables into the correct backend area
- switch the logging service to `log_be.accdb`
- switch the translation service to `sys_be.accdb`
- extend postal-code coverage beyond the DACH region
- add ISO 3166-2 region structures in `sys_be.accdb`
- connect reports and document output to centralized currency formatting based on `refCurrencies`
- define versioning and update strategy for `sys_be.accdb`
- evaluate optional delta updates instead of full file replacement

## Open Questions

- how should log rotation and archival be implemented for `log_be.accdb`
- what is the update mechanism for `sys_be.accdb`
- how should customer-specific overrides of shared system reference data be handled
- should `sys_be.accdb` be read-only in production or maintained through a dedicated admin tool
- how should data quality updates in shared data pools be distributed and validated
