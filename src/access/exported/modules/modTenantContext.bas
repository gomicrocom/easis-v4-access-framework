Attribute VB_Name = "modTenantContext"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modTenantContext
' Purpose   : Manages the current tenant runtime context.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modTenantContext"

Private mCurrentTenantId As String
Private mCurrentTenantCode As String
Private mCurrentTenantName As String
Private mCurrentBackendPath As String
Private mTenantLoaded As Boolean

Public Sub InitializeTenantContext(Optional ByVal IniPath As String = vbNullString)
    On Error GoTo ErrorHandler

    mCurrentTenantId = modConfigIni.GetConfigValue(INI_SECTION_TENANT, TENANT_KEY_ID, "TENANT-DEFAULT", IniPath)
    mCurrentTenantCode = modConfigIni.GetConfigValue(INI_SECTION_TENANT, TENANT_KEY_CODE, "DEFAULT", IniPath)
    mCurrentTenantName = modConfigIni.GetConfigValue(INI_SECTION_TENANT, TENANT_KEY_NAME, "Default Tenant", IniPath)
    mCurrentBackendPath = modConfigIni.GetConfigValue(INI_SECTION_TENANT, TENANT_KEY_BACKEND_PATH, vbNullString, IniPath)
    mTenantLoaded = True

    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeTenantContext", _
        "Tenant context initialized for tenant code '" & mCurrentTenantCode & "'."
    Exit Sub

ErrorHandler:
    mTenantLoaded = False
    modErrorHandler.HandleError MODULE_NAME, "InitializeTenantContext", Err
End Sub

Public Sub ClearTenantContext()
    On Error GoTo ErrorHandler

    mCurrentTenantId = vbNullString
    mCurrentTenantCode = vbNullString
    mCurrentTenantName = vbNullString
    mCurrentBackendPath = vbNullString
    mTenantLoaded = False

    modLoggingHandler.LogInfo MODULE_NAME & ".ClearTenantContext", "Tenant context cleared."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ClearTenantContext", Err
End Sub

Public Function IsTenantInitialized() As Boolean
    IsTenantInitialized = mTenantLoaded
End Function

Public Property Get CurrentTenantId() As String
    CurrentTenantId = mCurrentTenantId
End Property

Public Property Get CurrentTenantCode() As String
    CurrentTenantCode = mCurrentTenantCode
End Property

Public Property Get CurrentTenantName() As String
    CurrentTenantName = mCurrentTenantName
End Property

Public Property Get CurrentTenantBackendPath() As String
    CurrentTenantBackendPath = mCurrentBackendPath
End Property
