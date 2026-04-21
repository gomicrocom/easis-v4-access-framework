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
Private mCurrentRoles As Object
Private mSessionStartedAt As Date
Private mSessionInitialized As Boolean

Public Sub InitializeSessionContext(Optional ByVal IniPath As String = vbNullString)
    On Error GoTo ErrorHandler

    Dim rawRoles As String

    mCurrentUserId = modConfigIni.GetConfigValue(INI_SECTION_SESSION, "UserId", "SYSTEM", IniPath)
    mCurrentUserName = modConfigIni.GetConfigValue(INI_SECTION_SESSION, "UserName", "System User", IniPath)
    rawRoles = modConfigIni.GetConfigValue(INI_SECTION_SESSION, "Roles", vbNullString, IniPath)
    If LenB(Trim$(rawRoles)) = 0 Then
        rawRoles = modConfigIni.GetConfigValue(INI_SECTION_SESSION, "RoleCode", ROLE_CODE_ADMIN, IniPath)
    End If

    InitializeCurrentRoles rawRoles
    mCurrentRoleCode = GetPrimaryRoleCode(ROLE_CODE_ADMIN)
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

    If Not mCurrentRoles Is Nothing Then
        mCurrentRoles.RemoveAll
    End If

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
    If LenB(mCurrentRoleCode) = 0 Then
        mCurrentRoleCode = GetPrimaryRoleCode()
    End If
    CurrentRoleCode = mCurrentRoleCode
End Property

Public Function GetCurrentUserRoles() As Collection
    On Error GoTo ErrorHandler

    Dim userRoles As Collection
    Dim roleKey As Variant

    EnsureRoleStore
    Set userRoles = New Collection

    For Each roleKey In mCurrentRoles.Keys
        userRoles.Add CStr(roleKey)
    Next roleKey

    Set GetCurrentUserRoles = userRoles
    Exit Function

ErrorHandler:
    Set GetCurrentUserRoles = New Collection
    modErrorHandler.HandleError MODULE_NAME, "GetCurrentUserRoles", Err
End Function

Public Function IsCurrentUserInRole(ByVal RoleCode As String) As Boolean
    On Error GoTo ErrorHandler

    Dim normalizedRole As String

    normalizedRole = NormalizeRoleCode(RoleCode)
    If LenB(normalizedRole) = 0 Then
        Exit Function
    End If

    EnsureRoleStore
    IsCurrentUserInRole = mCurrentRoles.Exists(normalizedRole)
    Exit Function

ErrorHandler:
    IsCurrentUserInRole = False
    modErrorHandler.HandleError MODULE_NAME, "IsCurrentUserInRole", Err
End Function

Public Property Get SessionStartedAt() As Date
    SessionStartedAt = mSessionStartedAt
End Property

Private Sub InitializeCurrentRoles(ByVal RoleList As String)
    On Error GoTo ErrorHandler

    Dim roleParts() As String
    Dim roleItem As Variant
    Dim normalizedRole As String

    EnsureRoleStore
    mCurrentRoles.RemoveAll

    roleParts = Split(Replace(RoleList, ";", ","), ",")
    For Each roleItem In roleParts
        normalizedRole = NormalizeRoleCode(CStr(roleItem))
        If LenB(normalizedRole) > 0 Then
            If Not mCurrentRoles.Exists(normalizedRole) Then
                mCurrentRoles.Add normalizedRole, True
            End If
        End If
    Next roleItem

    If mCurrentRoles.Count = 0 Then
        mCurrentRoles.Add NormalizeRoleCode(ROLE_CODE_ADMIN), True
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "InitializeCurrentRoles", Err
End Sub

Private Sub EnsureRoleStore()
    If mCurrentRoles Is Nothing Then
        Set mCurrentRoles = CreateObject("Scripting.Dictionary")
        mCurrentRoles.CompareMode = vbTextCompare
    End If
End Sub

Private Function NormalizeRoleCode(ByVal RoleCode As String) As String
    NormalizeRoleCode = UCase$(Trim$(RoleCode))
End Function

Private Function GetPrimaryRoleCode(Optional ByVal DefaultValue As String = vbNullString) As String
    On Error GoTo ErrorHandler

    Dim roleKey As Variant

    EnsureRoleStore

    For Each roleKey In mCurrentRoles.Keys
        GetPrimaryRoleCode = CStr(roleKey)
        Exit Function
    Next roleKey

    GetPrimaryRoleCode = NormalizeRoleCode(DefaultValue)
    Exit Function

ErrorHandler:
    GetPrimaryRoleCode = NormalizeRoleCode(DefaultValue)
    modErrorHandler.HandleError MODULE_NAME, "GetPrimaryRoleCode", Err
End Function
