# Architecture

## Overview

Easis Version 4 follows an MS Access VBA architecture with a tenant-facing frontend and a dedicated backend database for each tenant. The design aims to keep framework concerns separated from tenant data and optional business modules.

## Core Principles

- one backend per tenant to simplify data isolation and operational ownership
- shared framework patterns for forms, reports, classes, queries, and modules
- INI-based configuration for startup, environment selection, and backend binding
- multilingual support through centralized text resources and language-aware UI loading
- feature-based licensing to enable or disable modules at runtime

## Logical Layers

- startup and bootstrap: initializes configuration, licensing, language, and tenant context
- framework services: logging, configuration, translation, licensing, and navigation
- application modules: optional feature packages such as `CAMT054`, `PROPERTY_MGMT`, and `WINE_MGMT`
- data access: linked tables, queries, and import/export routines against the tenant backend

## Configuration Direction

System configuration is expected to come from INI files. Typical settings include frontend environment, backend location, tenant identifier, default language, enabled modules, and licensing parameters.

## Deployment Direction

The Access frontend can be distributed as a controlled application package, while each tenant backend remains separately managed. This supports module rollout, customer-specific operations, and safer maintenance boundaries.
