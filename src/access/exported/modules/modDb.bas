Attribute VB_Name = "modDb"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modDb
' Purpose   : Database foundation helpers for Access frontend and backend setup.
' Author    : Codex
' Version   : 0.1.1
'===============================================================================

Private Const MODULE_NAME As String = "modDb"

Private mLastValidatedBackendPath As String
Private mLastValidationContext As String
Private mLastValidationSucceeded As Boolean

Public Function GetCurrentDatabase() As DAO.Database
    On Error GoTo ErrorHandler

    Set GetCurrentDatabase = CurrentDb
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetCurrentDatabase", Err
    Set GetCurrentDatabase = Nothing
End Function

Public Function GetBackendPath() As String
    On Error GoTo ErrorHandler

    If IsTenantInitialized Then
        GetBackendPath = Trim$(CurrentTenantBackendPath)
    End If

    If LenB(GetBackendPath) = 0 Then
        GetBackendPath = Trim$(modConfigIni.GetConfigValue(INI_SECTION_TENANT, TENANT_KEY_BACKEND_PATH, vbNullString, ConfigFilePath))
    End If
    Exit Function

ErrorHandler:
    GetBackendPath = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "GetBackendPath", Err
End Function

Public Function BackendExists() As Boolean
    On Error GoTo ErrorHandler

    Dim BackendPath As String

    BackendPath = GetBackendPath()
    If LenB(BackendPath) = 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".BackendExists", "Backend path is not configured."
        Exit Function
    End If

    BackendExists = (LenB(Dir$(BackendPath, vbNormal)) > 0)
    Exit Function

ErrorHandler:
    BackendExists = False
    modErrorHandler.HandleError MODULE_NAME, "BackendExists", Err
End Function

Public Function ValidateBackendConfiguration() As Boolean
    On Error GoTo ErrorHandler

    Dim BackendPath As String
    Dim logContext As String

    BackendPath = GetBackendPath()
    logContext = BuildValidationContext()

    If LenB(BackendPath) = 0 Then
        ResetBackendValidationLogGuard
        modLoggingHandler.LogError MODULE_NAME & ".ValidateBackendConfiguration", _
            "Backend validation failed: no backend path configured. " & logContext
        Exit Function
    End If

    If Not BackendExists() Then
        ResetBackendValidationLogGuard
        modLoggingHandler.LogError MODULE_NAME & ".ValidateBackendConfiguration", _
            "Backend validation failed: file not found at '" & BackendPath & "'. " & logContext
        Exit Function
    End If

    If ShouldLogSuccessfulValidation(BackendPath, logContext) Then
        modLoggingHandler.LogInfo MODULE_NAME & ".ValidateBackendConfiguration", _
            "Backend configuration validated successfully for path '" & BackendPath & "'. " & logContext
    End If

    ValidateBackendConfiguration = True
    Exit Function

ErrorHandler:
    ResetBackendValidationLogGuard
    ValidateBackendConfiguration = False
    modErrorHandler.HandleError MODULE_NAME, "ValidateBackendConfiguration", Err
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

Private Function ShouldLogSuccessfulValidation(ByVal BackendPath As String, ByVal ValidationContext As String) As Boolean
    Dim normalizedPath As String
    Dim normalizedContext As String

    normalizedPath = Trim$(BackendPath)
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
