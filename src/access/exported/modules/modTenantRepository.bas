Attribute VB_Name = "modTenantRepository"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modTenantRepository
' Purpose   : Reads tenant-related configuration values from backend tables.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modTenantRepository"
Private Const TABLE_TEN_PARAMETER As String = "ten_parameter"
Private Const FIELD_PARAMETER_KEY As String = "param_key"
Private Const FIELD_PARAMETER_VALUE As String = "param_value"
Private Const FIELD_TENANT_CODE As String = "tenant_code"

Public Function GetTenantParameter(ByVal ParameterKey As String, Optional ByVal defaultValue As String = "") As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If LenB(Trim$(ParameterKey)) = 0 Then
        GetTenantParameter = defaultValue
        Exit Function
    End If

    If Not CanReadTenantParameters() Then
        GetTenantParameter = defaultValue
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_TEN_PARAMETER & "];", dbOpenSnapshot)

    GetTenantParameter = ResolveTenantParameterValue(rs, ParameterKey, defaultValue)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    GetTenantParameter = defaultValue
    modErrorHandler.HandleError MODULE_NAME, "GetTenantParameter", Err
    Resume CleanExit
End Function

Public Function HasTenantParameter(ByVal ParameterKey As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If LenB(Trim$(ParameterKey)) = 0 Then
        HasTenantParameter = False
        Exit Function
    End If

    If Not CanReadTenantParameters() Then
        HasTenantParameter = False
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_TEN_PARAMETER & "];", dbOpenSnapshot)

    HasTenantParameter = (LenB(ResolveTenantParameterValue(rs, ParameterKey, vbNullString)) > 0)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    HasTenantParameter = False
    modErrorHandler.HandleError MODULE_NAME, "HasTenantParameter", Err
    Resume CleanExit
End Function

Private Function CanReadTenantParameters() As Boolean
    Dim db As DAO.Database

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CanReadTenantParameters", _
            "Backend configuration is not ready for tenant parameter lookup."
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()

    If Not modDbSchema.TableExists(db, TABLE_TEN_PARAMETER) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CanReadTenantParameters", _
            "Table '" & TABLE_TEN_PARAMETER & "' is not available yet for tenant '" & ResolveTenantCode() & "'."
        Exit Function
    End If

    CanReadTenantParameters = True
End Function

Private Function ResolveTenantCode() As String
    If IsTenantInitialized() Then
        ResolveTenantCode = currentTenantCode
    Else
        ResolveTenantCode = vbNullString
    End If
End Function

Private Function ResolveTenantParameterValue(ByVal rs As DAO.Recordset, ByVal ParameterKey As String, ByVal defaultValue As String) As String
    On Error GoTo ErrorHandler

    Dim targetKey As String
    Dim TenantCode As String
    Dim hasKeyField As Boolean
    Dim hasValueField As Boolean
    Dim hasTenantField As Boolean
    Dim currentKey As String
    Dim currentTenantCode As String

    targetKey = UCase$(Trim$(ParameterKey))
    TenantCode = UCase$(Trim$(ResolveTenantCode()))

    hasKeyField = modDaoHelper.RecordsetHasField(rs, FIELD_PARAMETER_KEY)
    hasValueField = modDaoHelper.RecordsetHasField(rs, FIELD_PARAMETER_VALUE)
    hasTenantField = modDaoHelper.RecordsetHasField(rs, FIELD_TENANT_CODE)

    If Not hasKeyField Or Not hasValueField Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ResolveTenantParameterValue", _
            "Required fields are not available in table '" & TABLE_TEN_PARAMETER & _
            "'. Expected fields: '" & FIELD_PARAMETER_KEY & "', '" & FIELD_PARAMETER_VALUE & "'."
        ResolveTenantParameterValue = defaultValue
        Exit Function
    End If

    If rs.BOF And rs.EOF Then
        ResolveTenantParameterValue = defaultValue
        Exit Function
    End If

    rs.MoveFirst
    Do Until rs.EOF
        currentKey = UCase$(Trim$(modDaoHelper.NzString(rs.Fields(FIELD_PARAMETER_KEY).Value)))

        If currentKey = targetKey Then
            If hasTenantField Then
                currentTenantCode = UCase$(Trim$(modDaoHelper.NzString(rs.Fields(FIELD_TENANT_CODE).Value)))

                If LenB(currentTenantCode) = 0 Then
                    ResolveTenantParameterValue = modDaoHelper.NzString(rs.Fields(FIELD_PARAMETER_VALUE).Value, defaultValue)
                    Exit Function
                End If

                If LenB(TenantCode) > 0 And currentTenantCode = TenantCode Then
                    ResolveTenantParameterValue = modDaoHelper.NzString(rs.Fields(FIELD_PARAMETER_VALUE).Value, defaultValue)
                    Exit Function
                End If
            Else
                ResolveTenantParameterValue = modDaoHelper.NzString(rs.Fields(FIELD_PARAMETER_VALUE).Value, defaultValue)
                Exit Function
            End If
        End If

        rs.MoveNext
    Loop

    ResolveTenantParameterValue = defaultValue
    Exit Function

ErrorHandler:
    ResolveTenantParameterValue = defaultValue
    modErrorHandler.HandleError MODULE_NAME, "ResolveTenantParameterValue", Err
End Function
