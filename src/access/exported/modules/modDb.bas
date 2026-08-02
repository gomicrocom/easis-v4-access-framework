Attribute VB_Name = "modDb"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modDb
' Purpose   : Database foundation helpers for Access frontend and backend setup.
' Author    : Codex
' Version   : 0.2.0
'===============================================================================

Private Const MODULE_NAME As String = "modDb"
Private Const DEFAULT_SYSTEM_BACKEND_PATH As String = "C:\easis\Data\sys_be.accdb"
Private Const SYSTEM_BACKEND_PATH_KEY As String = "SystemBackendPath"

Private mLastValidatedBackendPath As String
Private mLastValidationContext As String
Private mLastValidationSucceeded As Boolean

Public Function GetFrontendDatabase() As DAO.Database
    On Error GoTo ErrorHandler

    Set GetFrontendDatabase = CurrentDb
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetFrontendDatabase", Err
    Set GetFrontendDatabase = Nothing
End Function

Public Function GetSystemBackendPath() As String
    On Error GoTo ErrorHandler

    GetSystemBackendPath = Trim$(modConfigIni.GetConfigValue( _
        INI_SECTION_DATABASE, _
        SYSTEM_BACKEND_PATH_KEY, _
        vbNullString, _
        ConfigFilePath))

    If LenB(GetSystemBackendPath) = 0 Then
        GetSystemBackendPath = Trim$(modConfigIni.GetConfigValue( _
            INI_SECTION_PATHS, _
            SYSTEM_BACKEND_PATH_KEY, _
            DEFAULT_SYSTEM_BACKEND_PATH, _
            ConfigFilePath))
    End If

    If LenB(GetSystemBackendPath) = 0 Then
        modLoggingHandler.LogError MODULE_NAME & ".GetSystemBackendPath", _
            "System backend path could not be resolved from configuration."
    End If
    Exit Function

ErrorHandler:
    GetSystemBackendPath = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "GetSystemBackendPath", Err
End Function

Public Function GetCurrentTenantBackendPath() As String
    On Error GoTo ErrorHandler

    If modTenantContext.IsTenantInitialized Then
        GetCurrentTenantBackendPath = Trim$(modTenantContext.CurrentTenantBackendPath)
    End If

    If LenB(GetCurrentTenantBackendPath) = 0 Then
        GetCurrentTenantBackendPath = Trim$(modConfigIni.GetConfigValue(INI_SECTION_TENANT, TENANT_KEY_BACKEND_PATH, vbNullString, ConfigFilePath))
    End If

    If LenB(GetCurrentTenantBackendPath) = 0 Then
        modLoggingHandler.LogError MODULE_NAME & ".GetCurrentTenantBackendPath", _
            "Tenant backend path could not be resolved."
    End If
    Exit Function

ErrorHandler:
    GetCurrentTenantBackendPath = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "GetCurrentTenantBackendPath", Err
End Function

Public Function GetSystemDatabase() As DAO.Database
    On Error GoTo ErrorHandler

    Set GetSystemDatabase = OpenPhysicalDatabaseByPath(GetSystemBackendPath(), "GetSystemDatabase")
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetSystemDatabase", Err
    Set GetSystemDatabase = Nothing
End Function

Public Function GetCurrentTenantDatabase() As DAO.Database
    On Error GoTo ErrorHandler

    Set GetCurrentTenantDatabase = OpenPhysicalDatabaseByPath(GetCurrentTenantBackendPath(), "GetCurrentTenantDatabase")
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetCurrentTenantDatabase", Err
    Set GetCurrentTenantDatabase = Nothing
End Function

Public Function BackendExists() As Boolean
    On Error GoTo ErrorHandler

    Dim backendPath As String

    backendPath = GetCurrentTenantBackendPath()
    If LenB(backendPath) = 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".BackendExists", "Backend path is not configured."
        Exit Function
    End If

    BackendExists = (LenB(Dir$(backendPath, vbNormal)) > 0)
    Exit Function

ErrorHandler:
    BackendExists = False
    modErrorHandler.HandleError MODULE_NAME, "BackendExists", Err
End Function

Public Function ValidateBackendConfiguration() As Boolean
    On Error GoTo ErrorHandler

    Dim backendPath As String
    Dim logContext As String

    backendPath = GetCurrentTenantBackendPath()
    logContext = BuildValidationContext()

    If LenB(backendPath) = 0 Then
        ResetBackendValidationLogGuard
        modLoggingHandler.LogError MODULE_NAME & ".ValidateBackendConfiguration", _
            "Backend validation failed: no backend path configured. " & logContext
        Exit Function
    End If

    If Not BackendExists() Then
        ResetBackendValidationLogGuard
        modLoggingHandler.LogError MODULE_NAME & ".ValidateBackendConfiguration", _
            "Backend validation failed: file not found at '" & backendPath & "'. " & logContext
        Exit Function
    End If

    If ShouldLogSuccessfulValidation(backendPath, logContext) Then
        modLoggingHandler.LogInfo MODULE_NAME & ".ValidateBackendConfiguration", _
            "Backend configuration validated successfully for path '" & backendPath & "'. " & logContext
    End If

    ValidateBackendConfiguration = True
    Exit Function

ErrorHandler:
    ResetBackendValidationLogGuard
    ValidateBackendConfiguration = False
    modErrorHandler.HandleError MODULE_NAME, "ValidateBackendConfiguration", Err
End Function

Private Function OpenPhysicalDatabaseByPath( _
    ByVal databasePath As String, _
    ByVal callerName As String) As DAO.Database
    On Error GoTo ErrorHandler

    Dim normalizedPath As String

    normalizedPath = Trim$(databasePath)
    If LenB(normalizedPath) = 0 Then
        modLoggingHandler.LogError MODULE_NAME & "." & callerName, _
            "Database path is empty."
        Exit Function
    End If

    If LenB(Dir$(normalizedPath, vbNormal)) = 0 Then
        modLoggingHandler.LogError MODULE_NAME & "." & callerName, _
            "Database file not found at '" & normalizedPath & "'."
        Exit Function
    End If

    Set OpenPhysicalDatabaseByPath = DBEngine.OpenDatabase(normalizedPath)
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, callerName, Err
    Set OpenPhysicalDatabaseByPath = Nothing
End Function

Private Function BuildValidationContext() As String
    Dim contextParts As String

    If IsTenantInitialized Then
        contextParts = "TenantCode=" & currentTenantCode
    Else
        contextParts = "TenantCode=<uninitialized>"
    End If

    If IsSessionInitialized Then
        contextParts = contextParts & ", UserId=" & currentUserId
    Else
        contextParts = contextParts & ", UserId=<uninitialized>"
    End If

    BuildValidationContext = contextParts
End Function

Private Function ShouldLogSuccessfulValidation(ByVal backendPath As String, ByVal ValidationContext As String) As Boolean
    Dim normalizedPath As String
    Dim normalizedContext As String

    normalizedPath = Trim$(backendPath)
    normalizedContext = Trim$(ValidationContext)

    If mLastValidationSucceeded Then
        If StrComp(mLastValidatedBackendPath, normalizedPath, vbTextCompare) = 0 And _
           StrComp(mLastValidationContext, normalizedContext, vbTextCompare) = 0 Then
            Exit Function
        End If
    End If

    mLastValidatedBackendPath = normalizedPath
    mLastValidationContext = normalizedContext
    mLastValidationSucceeded = True
    ShouldLogSuccessfulValidation = True
End Function

Private Sub ResetBackendValidationLogGuard()
    mLastValidatedBackendPath = vbNullString
    mLastValidationContext = vbNullString
    mLastValidationSucceeded = False
End Sub
