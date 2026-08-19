Attribute VB_Name = "modUserRepository"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modUserRepository
' Purpose   : Reads user metadata from backend user tables.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modUserRepository"
Private Const TABLE_USR_USER As String = "usr_user"
Private Const FIELD_USER_ID As String = "user_id"
Private Const FIELD_USER_NAME As String = "user_name"
Private Const FIELD_ROLE_CODE As String = "role_code"
Private Const FIELD_LANGUAGE_CODE As String = "language_code"
Private Const FIELD_IS_ACTIVE As String = "is_active"

Public Function UserExists(ByVal UserId As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If LenB(Trim$(UserId)) = 0 Then
        Exit Function
    End If

    If Not CanReadUsers() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentTenantDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_USR_USER & "];", dbOpenSnapshot)

    UserExists = FindActiveUser(rs, UserId)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    UserExists = False
    modErrorHandler.HandleError MODULE_NAME, "UserExists", Err
    Resume CleanExit
End Function

Public Function GetUserDisplayName(ByVal UserId As String, Optional ByVal defaultValue As String = "") As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If LenB(Trim$(UserId)) = 0 Then
        GetUserDisplayName = defaultValue
        Exit Function
    End If

    If Not CanReadUsers() Then
        GetUserDisplayName = defaultValue
        Exit Function
    End If

    Set db = modDb.GetCurrentTenantDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_USR_USER & "];", dbOpenSnapshot)

    GetUserDisplayName = ResolveUserFieldValue(rs, UserId, FIELD_USER_NAME, defaultValue)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    GetUserDisplayName = defaultValue
    modErrorHandler.HandleError MODULE_NAME, "GetUserDisplayName", Err
    Resume CleanExit
End Function

Public Function GetUserRoleCode(ByVal UserId As String, Optional ByVal defaultValue As String = "USER") As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If LenB(Trim$(UserId)) = 0 Then
        GetUserRoleCode = defaultValue
        Exit Function
    End If

    If Not CanReadUsers() Then
        GetUserRoleCode = defaultValue
        Exit Function
    End If

    Set db = modDb.GetCurrentTenantDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_USR_USER & "];", dbOpenSnapshot)

    GetUserRoleCode = ResolveUserFieldValue(rs, UserId, FIELD_ROLE_CODE, defaultValue)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    GetUserRoleCode = defaultValue
    modErrorHandler.HandleError MODULE_NAME, "GetUserRoleCode", Err
    Resume CleanExit
End Function

Private Function CanReadUsers() As Boolean
    Dim db As DAO.Database

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CanReadUsers", _
            "Backend configuration is not ready for user lookup."
        Exit Function
    End If

    Set db = modDb.GetCurrentTenantDatabase()

    If Not modDbSchema.TableExists(db, TABLE_USR_USER) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CanReadUsers", _
            "Table '" & TABLE_USR_USER & "' is not available yet for tenant '" & ResolveTenantCode() & "'."
        Exit Function
    End If

    CanReadUsers = True
End Function

Private Function ResolveTenantCode() As String
    If IsTenantInitialized() Then
        ResolveTenantCode = currentTenantCode
    Else
        ResolveTenantCode = vbNullString
    End If
End Function

Private Function FindActiveUser(ByVal rs As DAO.Recordset, ByVal UserId As String) As Boolean
    On Error GoTo ErrorHandler

    FindActiveUser = (LenB(ResolveUserFieldValue(rs, UserId, FIELD_USER_ID, vbNullString)) > 0)
    Exit Function

ErrorHandler:
    FindActiveUser = False
    modErrorHandler.HandleError MODULE_NAME, "FindActiveUser", Err
End Function

Private Function ResolveUserFieldValue(ByVal rs As DAO.Recordset, ByVal UserId As String, ByVal TargetField As String, ByVal defaultValue As String) As String
    On Error GoTo ErrorHandler

    Dim targetUserId As String
    Dim hasUserIdField As Boolean
    Dim hasTargetField As Boolean
    Dim hasActiveField As Boolean
    Dim currentUserId As String

    targetUserId = UCase$(Trim$(UserId))
    hasUserIdField = modDaoHelper.RecordsetHasField(rs, FIELD_USER_ID)
    hasTargetField = modDaoHelper.RecordsetHasField(rs, TargetField)
    hasActiveField = modDaoHelper.RecordsetHasField(rs, FIELD_IS_ACTIVE)

    If Not hasUserIdField Or Not hasTargetField Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ResolveUserFieldValue", _
            "Required fields are not available in table '" & TABLE_USR_USER & "'."
        ResolveUserFieldValue = defaultValue
        Exit Function
    End If

    If rs.BOF And rs.EOF Then
        ResolveUserFieldValue = defaultValue
        Exit Function
    End If

    rs.MoveFirst
    Do Until rs.EOF
        currentUserId = UCase$(Trim$(modDaoHelper.NzString(rs.Fields(FIELD_USER_ID).Value)))

        If currentUserId = targetUserId Then
            If hasActiveField Then
                If Not modDaoHelper.NzBoolean(rs.Fields(FIELD_IS_ACTIVE).Value, False) Then
                    ResolveUserFieldValue = defaultValue
                    Exit Function
                End If
            End If

            ResolveUserFieldValue = modDaoHelper.NzString(rs.Fields(TargetField).Value, defaultValue)
            Exit Function
        End If

        rs.MoveNext
    Loop

    ResolveUserFieldValue = defaultValue
    Exit Function

ErrorHandler:
    ResolveUserFieldValue = defaultValue
    modErrorHandler.HandleError MODULE_NAME, "ResolveUserFieldValue", Err
End Function

Public Function GetUserLanguageCode(ByVal UserId As String, Optional ByVal defaultValue As String = "") As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If LenB(Trim$(UserId)) = 0 Then
        GetUserLanguageCode = defaultValue
        Exit Function
    End If

    If Not CanReadUsers() Then
        GetUserLanguageCode = defaultValue
        Exit Function
    End If

    Set db = modDb.GetCurrentTenantDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_USR_USER & "];", dbOpenSnapshot)

    If Not modDaoHelper.RecordsetHasField(rs, FIELD_LANGUAGE_CODE) Then
        GetUserLanguageCode = defaultValue
        GoTo CleanExit
    End If

    GetUserLanguageCode = ResolveUserFieldValue(rs, UserId, FIELD_LANGUAGE_CODE, defaultValue)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    GetUserLanguageCode = defaultValue
    modErrorHandler.HandleError MODULE_NAME, "GetUserLanguageCode", Err
    Resume CleanExit
End Function
