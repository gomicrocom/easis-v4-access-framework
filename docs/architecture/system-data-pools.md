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

### `refPostalCodes_DACH`

Purpose:

- postal-code pool for structured address capture in the DACH region

Key fields:

| Field | Meaning |
|---|---|
| `CountryCodeISO2` | foreign key to `refCountries` |
| `PostalCode` | postal code |
| `PlaceName` | place name |
| `Admin1Name` | region, canton, or state |
| `MunicipalityID` | municipality identifier |
| `Language` | language code |
| `IsPrimary` | preferred record indicator |

Additional fields:

- `Latitude`
- `Longitude`
- `SourceFile`
- `SourceQuality`

## Referential Integrity

The intended core relations are:

| Parent | Child | Meaning |
|---|---|---|
| `refCountries.ALPHA-2` | `refCountryTimezones.ALPHA-2` | country to timezone |
| `refCountries.ALPHA-2` | `refPostalCodes_DACH.CountryCodeISO2` | country to postal-code set |

## Usage Rules

The reference data pools are used for:

- structured data entry in the frontend
- address validation
- standardization of country and region data
- multilingual groundwork
- future internationalization

Tenant databases should store only references such as country codes and not full copies of shared system master data.

## Future Work

- introduce `refCurrencies` for ISO 4217 currencies
- extend postal-code coverage globally
- add ISO 3166-2 region tables
- define versioning and update strategy for `sys_be.accdb`
- evaluate optional delta updates instead of full file replacement

## Open Questions

- how should `sys_be.accdb` updates be delivered at customer installations
- how should customer-specific overrides be handled
- should `sys_be.accdb` be read-only in production or maintained through a dedicated tool
- how should data quality updates be distributed and validated
