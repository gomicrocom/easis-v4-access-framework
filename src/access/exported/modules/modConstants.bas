Attribute VB_Name = "modConstants"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modConstants
' Purpose   : Central constant definitions for configuration and runtime values.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

' Application parameter names
Public Const APP_PARAM_TENANT_ID As String = "CurrentTenantId"
Public Const APP_PARAM_TENANT_CODE As String = "CurrentTenantCode"
Public Const APP_PARAM_TENANT_NAME As String = "CurrentTenantName"
Public Const APP_PARAM_BACKEND_PATH As String = "CurrentBackendPath"
Public Const APP_PARAM_USER_ID As String = "CurrentUserId"
Public Const APP_PARAM_USER_NAME As String = "CurrentUserName"
Public Const APP_PARAM_ROLE_CODE As String = "CurrentRoleCode"
Public Const APP_PARAM_SESSION_STARTED_AT As String = "SessionStartedAt"

' INI section names
Public Const INI_SECTION_APPLICATION As String = "Application"
Public Const INI_SECTION_PATHS As String = "Paths"
Public Const INI_SECTION_DATABASE As String = "Database"
Public Const INI_SECTION_TENANT As String = "Tenant"
Public Const INI_SECTION_DEBUG As String = "Debug"
Public Const INI_SECTION_LICENSE As String = "License"
Public Const INI_SECTION_SESSION As String = "Session"

' Tenant parameter keys
Public Const TENANT_KEY_ID As String = "TenantId"
Public Const TENANT_KEY_CODE As String = "TenantCode"
Public Const TENANT_KEY_NAME As String = "TenantName"
Public Const TENANT_KEY_BACKEND_PATH As String = "BackendPath"

' Role codes
Public Const ROLE_CODE_ADMIN As String = "ADMIN"
Public Const ROLE_CODE_MANAGER As String = "MANAGER"
Public Const ROLE_CODE_USER As String = "USER"
Public Const ROLE_CODE_AUDITOR As String = "AUDITOR"
Public Const ROLE_CODE_GUEST As String = "GUEST"

' Document type codes
Public Const DOC_TYPE_INVOICE As String = "INVOICE"
Public Const DOC_TYPE_CREDIT_NOTE As String = "CREDIT_NOTE"
Public Const DOC_TYPE_PAYMENT As String = "PAYMENT"
Public Const DOC_TYPE_RECEIPT As String = "RECEIPT"
Public Const DOC_TYPE_ORDER As String = "ORDER"
