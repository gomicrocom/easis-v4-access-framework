# Architecture

## Overview

Easis Version 4 uses an MS Access VBA frontend with a dedicated backend database for each tenant. The framework is designed to separate shared application services from tenant data and optional domain modules.

## Core Principles

- one backend per tenant to simplify data isolation and operational ownership
- shared framework patterns for forms, reports, classes, queries, and modules
- INI-based configuration for startup, environment selection, and backend binding
- multilingual support through centralized text resources and language-aware UI loading
- feature-based licensing to enable or disable modules at runtime

## Logical Layers

- startup and bootstrap: initializes configuration, licensing, language, and tenant context
- framework services: configuration, translation, licensing, navigation, and common utilities
- application modules: optional feature packages such as `CAMT054`, `PROPERTY_MGMT`, and `WINE_MGMT`
- data access: linked tables, queries, and import/export routines against the tenant backend

## Configuration

System behavior is expected to be driven by INI files. Typical settings include environment, backend location, tenant identifier, default language, enabled modules, and licensing parameters.

## Deployment

The Access frontend can be distributed as a controlled application package while tenant backends remain separately managed. This supports safer rollout, customer-specific operations, and clearer maintenance boundaries.
