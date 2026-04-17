# Licensing

## Goal

Licensing in Easis Version 4 is feature-based. A tenant can run the core framework plus only the modules and capabilities that are licensed for that installation.

## Licensing Model

- core framework license enables application startup and shared services
- feature flags control access to optional capabilities
- module licenses may cover packages such as `CAMT054`, `PROPERTY_MGMT`, and `WINE_MGMT`
- licensing decisions should be available to VBA code, forms, menus, and reports

## Runtime Expectations

- license state is loaded during application startup
- unavailable features remain hidden or blocked in the user interface
- module startup routines should check entitlement before activation
- tenant-specific configuration may define license source and validation settings

## Configuration

INI files can define license lookup parameters, local validation settings, and environment-specific behavior. Final license storage and validation mechanics will be documented once implementation choices are fixed.
