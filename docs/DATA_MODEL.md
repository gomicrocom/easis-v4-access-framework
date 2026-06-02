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

- `ten_payment_term`
  - tenant-level payment term definitions used by orders and documents
- `ref_vat_code`
  - translated VAT code and VAT-rate definitions
- `ref_unit`
  - translated quantity and unit definitions
- `ref_article_type_code`
  - translated article-type behavior codes for extensibility and runtime specialization
- `fw_list_action`
  - configurable list navigation and action source for business UI flows

### Master Tables

Master tables hold reusable business entities.

- `adr_address`
  - authoritative address and business partner master data
- `art_product_group`
  - product and service grouping definitions
- `art_article`
  - article master data used in order entry

### Transaction Tables

Transaction tables represent operational business flow.

- `ord_order`
  - order header records
- `ord_order_line`
  - order position records

## Primary Keys

The exact physical field names may evolve by implementation detail, but the intended primary-key structure is:

| Table | Primary Key |
|---|---|
| `adr_address` | `address_id` |
| `art_product_group` | `product_group_id` |
| `art_article` | `article_id` |
| `ord_order` | `order_id` |
| `ord_order_line` | `order_line_id` |
| `ten_payment_term` | `payment_term_id` with unique `payment_term_code` + `language_code` |
| `ref_vat_code` | `vat_code` |
| `ref_unit` | `unit_code` |
| `ref_article_type_code` | `article_type_code` |
| `fw_list_action` | `action_id` |

If a table uses a business code as a technical primary identifier, that code must remain stable and unique within the tenant backend.

## Important Relationships

The following business relationships are expected to be central:

| Parent | Child | Purpose |
|---|---|---|
| `art_product_group` | `art_article` | article classification |
| `ref_article_type_code` | `art_article.article_type_code` | article behavior and future specialization |
| `adr_address` | `ord_order.address_id` | customer / invoice / delivery linkage |
| `ord_order` | `ord_order_line` | order header to line items |
| `art_article` | `ord_order_line` | line-level article reference |
| `ref_unit` | `art_article.unit_code` / `ord_order_line.unit_code` | unit standardization |
| `ref_vat_code` | `art_article.vat_code` / `ord_order_line.vat_code` | VAT assignment |
| `ten_payment_term` | `ord_order` | commercial payment handling via `payment_term_code` |
| `fw_list_action` | `frm<Entity>List` | configurable navigation/action menu |

Typical tenant-business relationships include:

- one address used across many orders
- one order with many order lines
- one article reused across many order lines
- one VAT code reused across articles and lines
- one payment term reused across many orders

## Naming Conventions

The repository currently contains two naming histories:

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

- `adr_address` is the authoritative address master table
- `tblAddresses` is deprecated and must not be recreated by current setup paths
- new design work must use the current physical tables and fields directly

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

### `adr_address`

Stores customer and address master data used in business documents and order processing.

### `art_product_group`

Stores the grouping structure used to classify articles and support article organization.

In the UI, this object is presented as `Artikelgruppe` / `Article Group`, but
the existing physical tenant table and key naming stay aligned with the
established Easis business model:

- table
  - `art_product_group`
- key fields
  - `product_group_id`
  - `product_group_code`
  - `product_group_name`

### `art_article`

Stores sellable products and services, including business defaults such as unit and VAT context.

Current physical field direction:

- `article_id`
- `article_no`
- `article_name`
- `product_group_id`
- `article_type_code`
- `unit_code`
- `vat_code`
- `purchase_price`
- `sales_price`
- `barcode`
- `description_text`
- `is_active`
- `created_at`
- `created_by`
- `updated_at`
- `updated_by`

`product_group_id` links to `art_product_group.product_group_id`.
`article_type_code` links to `ref_article_type_code.article_type_code`.
Business article values remain language-neutral; only UI and output layers are localized.

#### Product Group vs Article Type

These are intentionally different concepts:

- `product_group_id`
  - business grouping for reporting, organization, and categorization
- `article_type_code`
  - behavior and specialization hook for future article extensions

Examples:

- product group
  - Wine
  - Furniture
  - Services
- article type
  - `PRODUCT`
  - `SERVICE`
  - `SUBSCRIPTION`
  - `FEE`
  - `DISCOUNT`
  - `WINE`
  - `CUSTOM_SIZE`
  - `APPAREL_SIZE`

Future extension-table strategy:

- `art_article`
  - shared base article record
- future extension tables
  - `art_article_wine`
  - `art_article_dimension`
  - `art_article_apparel`
  - `art_article_subscription`

These extension tables are architectural targets only and are not implemented yet.

### `ord_order`

Stores the commercial order header, customer linkage, status, date information, and later document-generation context.

The current physical field set is:

- `order_id`
- `order_no`
- `address_id`
- `order_date`
- `document_status_code`
- `created_at`
- `created_by`
- `updated_at`
- `updated_by`

### `ord_order_line`

Stores quantity, article, pricing, and VAT-relevant transactional detail for each order position.

The current physical field set is:

- `order_line_id`
- `order_id`
- `article_id`
- `line_no`
- `quantity`
- `unit_code`
- `unit_price`
- `vat_code`
- `line_total`
- `created_at`
- `created_by`
- `updated_at`
- `updated_by`

### `ten_payment_term`

Stores tenant-level payment-term definitions used for order and document communication.

### `ref_vat_code`

Stores VAT reference definitions used in calculations and document presentation.

The current intended structure includes:

- `vat_code`
- `translation_key`
- `vat_rate`
- `country_code`
- `valid_from`
- `valid_to`
- `sort_order`
- `is_active`
- `created_at`
- `created_by`
- `updated_at`
- `updated_by`

### `ref_unit`

Stores translated standardized unit definitions used by articles and order lines.

The current intended structure includes:

- `unit_code`
- `translation_key`
- `sort_order`
- `is_active`
- `created_at`
- `created_by`
- `updated_at`
- `updated_by`

### `ref_article_type_code`

Stores the translated article-type behavior definitions used to classify base
article behavior independently from business grouping.

Current intended structure includes:

- `article_type_code`
- `article_type_name`
- `translation_key`
- `description_text`
- `sort_order`
- `is_active`
- `created_at`
- `created_by`
- `updated_at`
- `updated_by`

### `fw_list_action`

Stores configurable action definitions for list-form driven navigation and workflow menus.

The initial concept includes:

- `action_id`
- `list_code`
- `action_code`
- `action_label`
- `target_form`
- `requires_selection`
- `module_code`
- `role_code`
- `sort_order`
- `is_active`
- `created_at`
- `created_by`
- `updated_at`
- `updated_by`

The intended use is framework-oriented rather than address-specific:

- one list form such as `frmAddressList` or `frmOrderList` can load its actions dynamically
- `target_form` points to standardized UI targets such as `frmAddressDetail`
- `module_code` and `role_code` allow later module- and permission-aware action filtering

## Roadmap Context

The current data model is intended to support the next business-layer work:

- `frmAddresses`
- `frmArticles`
- `frmOrders`
- order line handling
- calculation services
- document generation
- mail handling
