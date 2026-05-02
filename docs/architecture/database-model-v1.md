# Database Model v1

## Overview

Easis v4 separates its persisted data into distinct storage areas with different responsibilities and lifecycles.

| Area | Location | Purpose |
|---|---|---|
| Local configuration | `<AppPath>\Cfg\easis.ini` | installation-specific configuration |
| Tenant backend | `<AppPath>\Data\<TenantCode>_be.accdb` | tenant business data |
| System backend | `<AppPath>\Data\sys_be.accdb` | system-wide reference data pools |
| Log backend | `<AppPath>\Data\log_be.accdb` | technical and operational logs |

## Naming Convention

The following prefixes are mandatory:

| Prefix | Meaning | Backend |
|---|---|---|
| `ref*` | system-wide reference data | `sys_be.accdb` |
| `sys*` | technical system tables if needed later | `sys_be.accdb` |
| `log*` | logging tables | `log_be.accdb` |
| `ten*` | tenant parameters | tenant backend |
| `doc*` | document data | tenant backend |
| `adr*` | address data | tenant backend |

Important rule:

The former `tbl*` prefix is no longer part of the active naming convention.

## Backend Responsibilities

### Local Configuration

Location:

```text
<AppPath>\Cfg\easis.ini
```

Contains:

- local backend paths
- `TenantCode`
- `Environment`
- installation-local defaults

Does not contain:

- tenant business data
- shared system data
- logs

### Tenant Backend

Location:

```text
<AppPath>\Data\<TenantCode>_be.accdb
```

Contains:

- `adr*` address tables
- `doc*` document tables
- `ten*` tenant parameters
- tenant-specific operational and business data

Examples of tenant parameters:

- `BASIC_DOC_PATH`
- `ADDRESS_WINDOW_POSITION`
- `DEFAULT_LANGUAGE`

Does not contain:

- shared system reference tables
- local machine configuration
- technical log tables

### System Backend

Location:

```text
<AppPath>\Data\sys_be.accdb
```

Contains only:

- `ref*` system-wide reference data
- optionally later `sys*` technical system tables

Important rules:

- no tenant data
- no local settings
- no logs
- the database must remain updateable by exchanging the file

Current reference tables:

- `refCountries`
- `refCountryTimezones`
- `refPostalCodes_DACH`
- `refCurrencies`

### Log Backend

Location:

```text
<AppPath>\Data\log_be.accdb
```

Contains:

- `log*` technical logs
- framework logs
- error logs
- runtime, export, and import logs

Does not contain:

- tenant business data
- shared reference data
- local configuration

## System Data Pool Model

Tenant databases should store only reference keys into shared pools where appropriate, for example:

- country code instead of a copied country master record
- currency code instead of a copied currency master record

This keeps tenant backends smaller, avoids duplication, and allows central maintenance of shared standards.

### `refCountries`

Purpose:

- central country master based on ISO 3166-1 plus extended metadata

Used for:

- address validation
- language defaults
- currency defaults
- regional classification
- internationalization groundwork

Key fields:

| Field | Meaning |
|---|---|
| `ALPHA-2` | primary identifier |
| `ALPHA-3` | ISO alpha-3 code |
| `CountryName` | display name |
| `OfficialName` | official country name |
| `MainCurrency` | primary currency code |
| `MainLanguageCode` | primary language code |
| `PhonePrefix` | main phone prefix |
| `Continent` | continent |
| `Subregion` | subregion |
| `EU Member` | EU membership indicator |
| `Capital` | capital city |

Extended fields:

- `AllCurrencies`
- `AllLanguages`
- `AllPhonePrefixes`
- `Timezones`
- DST and UTC offset fields
- `DataSources`
- `DataQualityNotes`

### `refCountryTimezones`

Purpose:

- timezone definitions per country

Key fields:

| Field | Meaning |
|---|---|
| `ALPHA-2` | foreign key to `refCountries` |
| `Timezone IANA` | IANA timezone identifier |
| `UTC Offset Standard` | standard UTC offset |
| `UTC Offset DST` | daylight-saving UTC offset |
| `DST Detail` | daylight-saving detail |

### `refPostalCodes_DACH`

Purpose:

- structured postal-code data pool for DACH address handling

Key fields:

| Field | Meaning |
|---|---|
| `CountryCodeISO2` | foreign key to `refCountries` |
| `PostalCode` | postal code |
| `PlaceName` | place or locality |
| `Admin1Name` | region, canton, or state |
| `MunicipalityID` | municipality identifier |
| `Language` | language of the place name |
| `IsPrimary` | preferred row indicator |

Additional fields:

- `Latitude`
- `Longitude`
- `SourceFile`
- `SourceQuality`

### `refCurrencies`

Purpose:

- central ISO 4217 currency definition table
- basis for document currencies, reports, and centralized amount formatting

Key rules:

- `refCurrencies` belongs to `sys_be.accdb`
- tenant backends store only `CurrencyCode` references
- examples of tenant references are `doc_document.currency_code` and `ten_parameter.CURRENCY_CODE`
- no translated names are stored in `refCurrencies`
- translated names belong to the translation service

Fields:

| Field | Type | Meaning |
|---|---|---|
| `CurrencyCode` | `TEXT(3)` | primary key |
| `NumericCode` | `TEXT(3)` | ISO 4217 numeric code |
| `CurrencyName` | `TEXT(100)` | neutral name |
| `MinorUnit` | `BYTE` | number of decimal places |
| `Symbol` | `TEXT(10)` | display symbol |
| `SymbolPosition` | `TEXT(10)` | `PREFIX` or `SUFFIX` |
| `DecimalSeparator` | `TEXT(5)` | decimal separator |
| `ThousandSeparator` | `TEXT(5)` | thousand separator |
| `IsActive` | `YESNO` | active flag |
| `DataSource` | `TEXT(100)` | source provenance |
| `DataQualityNotes` | `TEXT(255)` | data quality notes |
| `timestamp` | `DATETIME` | last maintenance timestamp |

Initial values:

| CurrencyCode | NumericCode | CurrencyName | MinorUnit | Symbol | SymbolPosition | DecimalSeparator | ThousandSeparator | IsActive |
|---|---|---|---:|---|---|---|---|---|
| `CHF` | `756` | `Swiss Franc` | 2 | `CHF` | `PREFIX` | `.` | `'` | Yes |
| `EUR` | `978` | `Euro` | 2 | `€` | `SUFFIX` | `,` | `.` | Yes |
| `USD` | `840` | `US Dollar` | 2 | `$` | `PREFIX` | `.` | `,` | Yes |
| `GBP` | `826` | `Pound Sterling` | 2 | `£` | `PREFIX` | `.` | `,` | Yes |
| `JPY` | `392` | `Yen` | 0 | `¥` | `PREFIX` | `.` | `,` | Yes |

## Referential Integrity

The intended shared reference relationships are:

| Parent | Child | Relation |
|---|---|---|
| `refCountries.ALPHA-2` | `refCountryTimezones.ALPHA-2` | country to timezone |
| `refCountries.ALPHA-2` | `refPostalCodes_DACH.CountryCodeISO2` | country to postal-code pool |

These relations should be documented and enforced where practical in the system backend.

## Usage in the Frontend

The system data pools are intended for:

- structured data entry in the frontend
- address validation
- standardized country and region handling
- standardized currency formatting rules
- groundwork for multilingual behavior
- groundwork for broader internationalization

Reports and business logic should consume stored tenant references and resolve shared reference data through `sys_be.accdb`, not through copied tenant-local duplicates.

For currencies this means:

- business tables store `CurrencyCode`
- formatting metadata is resolved centrally from `refCurrencies`
- `MinorUnit` determines decimal places for display

## Currency Formatting Strategy

Currency formatting should be centralized in a dedicated service:

`modCurrencyFormatService`

Planned functions:

- `GetCurrencyFormat(CurrencyCode)`
- `FormatCurrencyAmount(Amount, CurrencyCode)`
- `NormalizeCurrencyCode(CurrencyCode)`

Report rule:

- reports should not use hardcoded Access currency format strings as the primary business rule
- reports should display formatted amounts through centralized currency formatting logic
- `qry_document_report_header` should later provide `currency_code`

## Future Work

- extend postal-code data beyond the DACH region
- add ISO 3166-2 region tables
- connect reports to centralized currency formatting
- define versioning and update strategy for `sys_be.accdb`
- evaluate delta update support instead of full file replacement

## Open Questions

- how should `sys_be.accdb` updates be delivered at customer sites
- how should customer-specific overrides be modeled and governed
- should system data be read-only in production or maintained through a dedicated tool
- how should data quality updates be reviewed and rolled out
