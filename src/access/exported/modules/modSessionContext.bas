Option Compare Database
Option Explicit

'===============================================================================
' Module    : modSessionContext
' Purpose   : Manages the current application session context.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modSessionContext"

Private mCurrentUserId As String
Private mCurrentUserName As String
Private mCurrentRoleCode As String
Private mSessionStartedAt As Date
Private mSessionInitialized As Boolean

Public Sub InitializeSessionContext(Optional ByVal IniPath As String = vbNullString)
    On Error GoTo ErrorHandler

    mCurrentUserId = modConfigIni.GetConfigValue(INI_SECTION_SESSION, "UserId", "SYSTEM", IniPath)
    mCurrentUserName = modConfigIni.GetConfigValue(INI_SECTION_SESSION, "UserName", "System User", IniPath)
    mCurrentRoleCode = modConfigIni.GetConfigValue(INI_SECTION_SESSION, "RoleCode", ROLE_CODE_ADMIN, IniPath)
    mSessionStartedAt = Now
    mSessionInitialized = True

    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeSessionContext", _
        "Session context initialized for user '" & mCurrentUserName & "'."
    Exit Sub

ErrorHandler:
    mSessionInitialized = False
    modErrorHandler.HandleError MODULE_NAME, "InitializeSessionContext", Err
End Sub

Public Sub ClearSessionContext()
    On Error GoTo ErrorHandler

    mCurrentUserId = vbNullString
    mCurrentUserName = vbNullString
    mCurrentRoleCode = vbNullString
    mSessionStartedAt = 0
    mSessionInitialized = False

    modLoggingHandler.LogInfo MODULE_NAME & ".ClearSessionContext", "Session context cleared."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ClearSessionContext", Err
End Sub

Public Function IsSessionInitialized() As Boolean
    IsSessionInitialized = mSessionInitialized
End Function

Public Property Get currentUserId() As String
    currentUserId = mCurrentUserId
End Property

Public Property Get CurrentUserName() As String
    CurrentUserName = mCurrentUserName
End Property

Public Property Get CurrentRoleCode() As String
    CurrentRoleCode = mCurrentRoleCode
End Property

Public Property Get SessionStartedAt() As Date
    SessionStartedAt = mSessionStartedAt
End Property