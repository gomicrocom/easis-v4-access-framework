# Database Model V1

## Purpose

Version 1 defines the baseline conventions for tenant-specific backend databases used by the Access frontend.

## Principles

- each tenant has its own backend database file
- framework metadata is stored separately from transactional business data where practical
- optional modules extend the model in a controlled way
- naming should remain stable across languages; translations belong in UI or metadata resources

## Baseline Areas

- system metadata: tenant identity, schema version, language defaults, and configuration markers
- licensing metadata: enabled features, entitlement dates, and validation status
- shared reference data: reusable lookup tables required by the framework
- module data: tables introduced by optional features like `CAMT054`, `PROPERTY_MGMT`, and `WINE_MGMT`

## Versioning

Schema changes should be introduced through explicit framework versioning so the frontend can detect compatibility and apply controlled migrations when needed.

## Current Boundary

Detailed table definitions are intentionally deferred until the framework services, licensing rules, and first module boundaries are finalized.
