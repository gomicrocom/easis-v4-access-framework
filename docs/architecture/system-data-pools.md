# System Data Pools

## Purpose

`sys_be.accdb` is the shared backend for tenant-independent reference data in Easis v4.

Location:

```text
<AppPath>\Data\sys_be.accdb
```

It contains centrally maintained data pools that are valid for all tenants and can be updated independently from tenant backends.

## Architectural Rules

- `sys_be.accdb` contains only reference data
- no tenant business data
- no local installation settings
- no technical logs
- the database must remain updateable by replacing the file

Related storage areas:

| Area | Location | Responsibility |
|---|---|---|
| Local configuration | `<AppPath>\Cfg\easis.ini` | installation-specific settings |
| Tenant backend | `<AppPath>\Data\<TenantCode>_be.accdb` | tenant business data and tenant parameters |
| System backend | `<AppPath>\Data\sys_be.accdb` | shared reference data pools |
| Log backend | `<AppPath>\Data\log_be.accdb` | technical and operational logs |

## Naming Convention

| Prefix | Meaning |
|---|---|
| `ref*` | system-wide reference data |
| `sys*` | technical system tables if added later |

Important rule:

The former `tbl*` prefix is no longer used.

## Current Reference Tables

### `refCountries`

Purpose:

- central country reference based on ISO 3166-1 with extended master data

Used for:

- address validation
- language defaults
- currency defaults
- regional grouping
- future internationalization features

Key fields:

| Field | Meaning |
|---|---|
| `ALPHA-2` | primary identifier |
| `ALPHA-3` | ISO alpha-3 code |
| `CountryName` | common country name |
| `OfficialName` | official country name |
| `MainCurrency` | primary currency |
| `MainLanguageCode` | primary language |
| `PhonePrefix` | main dialing prefix |
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

- timezone data per country

Key fields:

| Field | Meaning |
|---|---|
| `ALPHA-2` | foreign key to `refCountries` |
| `Timezone IANA` | IANA timezone identifier |
| `UTC Offset Standard` | standard UTC offset |
| `UTC Offset DST` | daylight-saving UTC offset |
| `DST Detail` | DST detail |

### `ref_postal_code`

Purpose:

- postal-code pool for structured address capture in CH, DE, and AT

Key fields:

| Field | Meaning |
|---|---|
| `country_code` | foreign key to `ref_country` |
| `postal_code` | postal code |
| `place_name` | place name |
| `state_name` | state, canton, or Bundesland name |
| `state_code` | state, canton, or Bundesland code |
| `province_name` | province or secondary region name |
| `province_code` | province or secondary region code |
| `community_name` | municipality or community name |
| `community_code` | municipality or community code |
| `is_active` | preferred record indicator |

Additional fields:

- `postal_code_id`
- `latitude`
- `longitude`
- `created_at`
- `created_by`
- `updated_at`
- `updated_by`

### `refCurrencies`

Purpose:

- central ISO 4217 currency reference for all tenants
- basis for document currency handling
- basis for report formatting and centralized amount display rules

Architectural rules:

- `refCurrencies` belongs to `sys_be.accdb`
- tenant backends store only `CurrencyCode` references, for example in `doc_document.currency_code` or `ten_parameter.CURRENCY_CODE`
- `refCurrencies` does not contain translated currency names
- translated labels are handled through the translation service
- `MinorUnit` controls the number of decimal places to display
- formatting logic should be centralized in `modCurrencyFormatService`

Fields:

| Field | Type | Meaning |
|---|---|---|
| `CurrencyCode` | `TEXT(3)` | primary key, ISO 4217 alpha code |
| `NumericCode` | `TEXT(3)` | ISO 4217 numeric code |
| `CurrencyName` | `TEXT(100)` | neutral currency name |
| `MinorUnit` | `BYTE` | decimal places used for display |
| `Symbol` | `TEXT(10)` | display symbol or currency code |
| `SymbolPosition` | `TEXT(10)` | `PREFIX` or `SUFFIX` |
| `DecimalSeparator` | `TEXT(5)` | decimal separator for formatting |
| `ThousandSeparator` | `TEXT(5)` | thousand separator for formatting |
| `IsActive` | `YESNO` | active flag |
| `DataSource` | `TEXT(100)` | source provenance |
| `DataQualityNotes` | `TEXT(255)` | quality or review notes |
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

The intended core relations are:

| Parent | Child | Meaning |
|---|---|---|
| `refCountries.ALPHA-2` | `refCountryTimezones.ALPHA-2` | country to timezone |
| `ref_country.country_code` | `ref_postal_code.country_code` | country to postal-code set |

`refCurrencies` is a standalone shared reference table keyed by `CurrencyCode`.

## Usage Rules

The reference data pools are used for:

- structured data entry in the frontend
- address validation
- standardization of country and region data
- standardization of currency handling and amount formatting
- multilingual groundwork
- future internationalization

Tenant databases should store only references such as country codes and not full copies of shared system master data.

For currencies this means:

- store `CurrencyCode` in tenant data
- resolve display and formatting rules from `refCurrencies`
- avoid tenant-local duplication of symbol, separator, or minor-unit metadata

## Currency Formatting Service

The currency formatting logic should be centralized in a dedicated service module:

`modCurrencyFormatService`

Planned functions:

- `GetCurrencyFormat(CurrencyCode)`
- `FormatCurrencyAmount(Amount, CurrencyCode)`
- `NormalizeCurrencyCode(CurrencyCode)`

Report rule:

- reports should not rely on raw Access currency format strings as the primary business rule
- reports should use centrally defined currency formatting behavior
- `qry_document_report_header` should later provide `currency_code` so reports can resolve the correct currency format context

## Future Work

- extend postal-code coverage globally
- add ISO 3166-2 region tables
- connect reports and document output to centralized currency formatting
- define versioning and update strategy for `sys_be.accdb`
- evaluate optional delta updates instead of full file replacement

## Open Questions

- how should `sys_be.accdb` updates be delivered at customer installations
- how should customer-specific overrides be handled
- should `sys_be.accdb` be read-only in production or maintained through a dedicated tool
- how should data quality updates be distributed and validated
- how should exceptional customer-specific currency formatting overrides be governed, if allowed at all
