# Data Model

## Overview

The current Easis v4 data model is split across distinct backend responsibilities:

- `FE.accdb`
  - frontend application, forms, reports, modules, and linked tables
- `Data\<TENANT>_be.accdb`
  - tenant-specific business and operational data
- `Data\sys_be.accdb`
  - system-wide reference data pools
- `Data\log_be.accdb`
  - technical and operational logs

This document focuses on the tenant backend structures now actively used by BasicModule v1.

## Logical Grouping

### Reference Tables

Reference tables provide controlled reusable business values for the tenant application layer.

- `refPaymentTerms`
  - payment term definitions used by orders and documents
- `refVatCodes`
  - VAT code and VAT-rate definitions
- `refUnits`
  - quantity and unit definitions
- `tblListActions`
  - configurable list navigation and action source for business UI flows

### Master Tables

Master tables hold reusable business entities.

- `tblAddresses`
  - address and business partner master data
- `tblProductGroups`
  - product and service grouping definitions
- `tblArticles`
  - article master data used in order entry

### Transaction Tables

Transaction tables represent operational business flow.

- `tblOrders`
  - order header records
- `tblOrderLines`
  - order position records

## Primary Keys

The exact physical field names may evolve by implementation detail, but the intended primary-key structure is:

| Table | Primary Key |
|---|---|
| `tblAddresses` | `AddressID` |
| `tblProductGroups` | `ProductGroupID` |
| `tblArticles` | `ArticleID` |
| `tblOrders` | `OrderID` |
| `tblOrderLines` | `OrderLineID` |
| `refPaymentTerms` | `PaymentTermID` or stable business code |
| `refVatCodes` | `VatCodeID` or stable VAT code |
| `refUnits` | `UnitID` or stable unit code |
| `tblListActions` | `ActionId` |

If a table uses a business code as a technical primary identifier, that code must remain stable and unique within the tenant backend.

## Important Relationships

The following business relationships are expected to be central:

| Parent | Child | Purpose |
|---|---|---|
| `tblProductGroups` | `tblArticles` | article classification |
| `tblAddresses` | `tblOrders` | customer / invoice / delivery linkage |
| `tblOrders` | `tblOrderLines` | order header to line items |
| `tblArticles` | `tblOrderLines` | line-level article reference |
| `refUnits` | `tblArticles` / `tblOrderLines` | unit standardization |
| `refVatCodes` | `tblArticles` / `tblOrderLines` | VAT assignment |
| `refPaymentTerms` | `tblOrders` | commercial payment handling |
| `tblListActions` | `frm<Entity>List` | configurable navigation/action menu |

Typical tenant-business relationships include:

- one address used across many orders
- one order with many order lines
- one article reused across many order lines
- one VAT code reused across articles and lines
- one payment term reused across many orders

## Naming Conventions

The repository currently contains two naming histories:

### Historical / Transitional Tables

- `tbl*`
  - existing business tables in the first business application layer

### Framework and Backend Naming Direction

- `ref*`
  - reference tables
- `sys*`
  - system tables in `sys_be.accdb`
- `log*`
  - log tables in `log_be.accdb`
- `ten*`
  - tenant settings and parameters
- `doc*`
  - document-related tables
- `adr*`
  - address-related tables

Current rule:

Business tables already introduced as `tbl*` remain valid as part of BasicModule v1. New design work should be explicit about whether a table is:

- part of the established business module naming
- or part of the newer framework/backend naming direction

## Business Module Interaction

The business layer interacts with the framework as follows:

- translations
  - labels, captions, and report texts are not hard-coded per form where a translation key exists
- tags
  - forms and controls use tag-driven runtime behavior
- validation
  - data-entry rules are executed centrally through framework validation services
- module access
  - business forms can be enabled or restricted by framework-managed access logic
- reporting
  - reports consume business data but reuse translation and runtime conventions
- logging
  - business services should use centralized logging rather than custom ad hoc tracing

## BasicModule v1 Table Roles

### `tblAddresses`

Stores customer and address master data used in business documents and order processing.

### `tblProductGroups`

Stores the grouping structure used to classify articles and support article organization.

### `tblArticles`

Stores sellable products and services, including business defaults such as unit and VAT context.

### `tblOrders`

Stores the commercial order header, customer linkage, status, date information, and later document-generation context.

### `tblOrderLines`

Stores quantity, article, pricing, and VAT-relevant transactional detail for each order position.

### `refPaymentTerms`

Stores reusable payment-term definitions used for order and document communication.

### `refVatCodes`

Stores VAT reference definitions used in calculations and document presentation.

### `refUnits`

Stores standardized unit definitions used by articles and order lines.

### `tblListActions`

Stores configurable action definitions for list-form driven navigation and workflow menus.

The initial concept includes:

- `ActionId`
- `ListCode`
- `ActionCode`
- `ActionLabel`
- `TargetForm`
- `RequiresSelection`
- `ModuleCode`
- `RoleCode`
- `SortOrder`
- `IsActive`

The intended use is framework-oriented rather than address-specific:

- one list form such as `frmAddressList` or `frmOrderList` can load its actions dynamically
- `TargetForm` points to standardized UI targets such as `frmAddressDetail`
- `ModuleCode` and `RoleCode` allow later module- and permission-aware action filtering

## Roadmap Context

The current data model is intended to support the next business-layer work:

- `frmAddresses`
- `frmArticles`
- `frmOrders`
- order line handling
- calculation services
- document generation
- mail handling
