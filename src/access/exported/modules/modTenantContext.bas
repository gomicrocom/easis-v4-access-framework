Option Compare Database
Option Explicit

'===============================================================================
' Module    : modTenantContext
' Purpose   : Manages the current tenant runtime context.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modTenantContext"
Private Const DEFAULT_TENANT_CODE_KEY As String = "DefaultTenantCode"
Private Const BACKEND_ROOT_KEY As String = "BackendRoot"

Private mCurrentTenantId As String
Private mCurrentTenantCode As String
Private mCurrentTenantName As String
Private mCurrentBackendPath As String
Private mTenantLoaded As Boolean

Public Sub InitializeTenantContext(Optional ByVal IniPath As String = vbNullString)
    On Error GoTo ErrorHandler

    Dim BackendRoot As String

    mCurrentTenantId = modConfigIni.GetConfigValue(INI_SECTION_TENANT, TENANT_KEY_ID, "TENANT-DEFAULT", IniPath)
    mCurrentTenantCode = modConfigIni.GetConfigValue(INI_SECTION_DATABASE, DEFAULT_TENANT_CODE_KEY, "DEFAULT", IniPath)
    mCurrentTenantName = modConfigIni.GetConfigValue(INI_SECTION_TENANT, TENANT_KEY_NAME, "Default Tenant", IniPath)

    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeTenantContext", _
        "Tenant code initialized as '" & mCurrentTenantCode & "'."

    BackendRoot = Trim$(modConfigIni.GetConfigValue(INI_SECTION_DATABASE, BACKEND_ROOT_KEY, vbNullString, IniPath))
    If LenB(BackendRoot) = 0 Then
        mCurrentBackendPath = vbNullString
        modLoggingHandler.LogWarning MODULE_NAME & ".InitializeTenantContext", _
            "Backend root configuration is missing."
    Else
        mCurrentBackendPath = BuildBackendPath(BackendRoot, mCurrentTenantCode)
        modLoggingHandler.LogInfo MODULE_NAME & ".InitializeTenantContext", _
            "Backend path initialized as '" & mCurrentBackendPath & "'."
    End If

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

Public Property Get currentTenantCode() As String
    currentTenantCode = mCurrentTenantCode
End Property

Public Property Get CurrentTenantName() As String
    CurrentTenantName = mCurrentTenantName
End Property

Public Property Get CurrentTenantBackendPath() As String
    CurrentTenantBackendPath = mCurrentBackendPath
End Property

Public Property Get CurrentBackendPath() As String
    CurrentBackendPath = mCurrentBackendPath
End Property

Private Function BuildBackendPath(ByVal BackendRoot As String, ByVal tenantCode As String) As String
    Dim normalizedRoot As String

    normalizedRoot = Trim$(BackendRoot)
    If Right$(normalizedRoot, 1) = "\" Then
        normalizedRoot = Left$(normalizedRoot, Len(normalizedRoot) - 1)
    End If

    BuildBackendPath = normalizedRoot & "\" & Trim$(tenantCode) & "_be.accdb"
End Function