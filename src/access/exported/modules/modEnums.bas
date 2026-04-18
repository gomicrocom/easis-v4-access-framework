Option Compare Database
Option Explicit

'===============================================================================
' Module    : modEnums
' Purpose   : Shared enumeration definitions for framework semantics.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Public Enum AppEnvironment
    AppEnvironmentUnknown = 0
    AppEnvironmentDevelopment = 1
    AppEnvironmentTest = 2
    AppEnvironmentProduction = 3
End Enum

Public Enum VatMode
    VatModeUnknown = 0
    VatModeExcluded = 1
    VatModeIncluded = 2
    VatModeExempt = 3
End Enum

Public Enum DocumentStatus
    DocumentStatusUnknown = 0
    DocumentStatusDraft = 1
    DocumentStatusOpen = 2
    DocumentStatusPosted = 3
    DocumentStatusCancelled = 4
    DocumentStatusArchived = 5
End Enum

Public Enum UserRoleType
    UserRoleTypeUnknown = 0
    UserRoleTypeGuest = 1
    UserRoleTypeUser = 2
    UserRoleTypeManager = 3
    UserRoleTypeAdministrator = 4
    UserRoleTypeAuditor = 5
End Enum