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

    Set db = modDb.GetCurrentDatabase()
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

Public Function GetUserDisplayName(ByVal UserId As String, Optional ByVal DefaultValue As String = "") As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If LenB(Trim$(UserId)) = 0 Then
        GetUserDisplayName = DefaultValue
        Exit Function
    End If

    If Not CanReadUsers() Then
        GetUserDisplayName = DefaultValue
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_USR_USER & "];", dbOpenSnapshot)

    GetUserDisplayName = ResolveUserFieldValue(rs, UserId, FIELD_USER_NAME, DefaultValue)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    GetUserDisplayName = DefaultValue
    modErrorHandler.HandleError MODULE_NAME, "GetUserDisplayName", Err
    Resume CleanExit
End Function

Public Function GetUserRoleCode(ByVal UserId As String, Optional ByVal DefaultValue As String = "USER") As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If LenB(Trim$(UserId)) = 0 Then
        GetUserRoleCode = DefaultValue
        Exit Function
    End If

    If Not CanReadUsers() Then
        GetUserRoleCode = DefaultValue
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_USR_USER & "];", dbOpenSnapshot)

    GetUserRoleCode = ResolveUserFieldValue(rs, UserId, FIELD_ROLE_CODE, DefaultValue)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    GetUserRoleCode = DefaultValue
    modErrorHandler.HandleError MODULE_NAME, "GetUserRoleCode", Err
    Resume CleanExit
End Function

Private Function CanReadUsers() As Boolean
    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CanReadUsers", _
            "Backend configuration is not ready for user lookup."
        Exit Function
    End If

    If Not TableExists(TABLE_USR_USER) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CanReadUsers", _
            "Table '" & TABLE_USR_USER & "' is not available yet for tenant '" & ResolveTenantCode() & "'."
        Exit Function
    End If

    CanReadUsers = True
End Function

Private Function TableExists(ByVal TableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = modDb.GetCurrentDatabase()

    For Each tdf In db.TableDefs
        If UCase$(Trim$(tdf.Name)) = UCase$(Trim$(TableName)) Then
            TableExists = True
            Exit For
        End If
    Next tdf

CleanExit:
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    TableExists = False
    modErrorHandler.HandleError MODULE_NAME, "TableExists", Err
    Resume CleanExit
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

Private Function ResolveUserFieldValue(ByVal rs As DAO.Recordset, ByVal UserId As String, ByVal TargetField As String, ByVal DefaultValue As String) As String
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
        ResolveUserFieldValue = DefaultValue
        Exit Function
    End If

    If rs.BOF And rs.EOF Then
        ResolveUserFieldValue = DefaultValue
        Exit Function
    End If

    rs.MoveFirst
    Do Until rs.EOF
        currentUserId = UCase$(Trim$(modDaoHelper.NzString(rs.Fields(FIELD_USER_ID).Value)))

        If currentUserId = targetUserId Then
            If hasActiveField Then
                If Not modDaoHelper.NzBoolean(rs.Fields(FIELD_IS_ACTIVE).Value, False) Then
                    ResolveUserFieldValue = DefaultValue
                    Exit Function
                End If
            End If

            ResolveUserFieldValue = modDaoHelper.NzString(rs.Fields(TargetField).Value, DefaultValue)
            Exit Function
        End If

        rs.MoveNext
    Loop

    ResolveUserFieldValue = DefaultValue
    Exit Function

ErrorHandler:
    ResolveUserFieldValue = DefaultValue
    modErrorHandler.HandleError MODULE_NAME, "ResolveUserFieldValue", Err
End Function
