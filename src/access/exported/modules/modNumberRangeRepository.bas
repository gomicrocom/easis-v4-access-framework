Attribute VB_Name = "modNumberRangeRepository"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modNumberRangeRepository
' Purpose   : DAO access helpers for tenant number range records.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modNumberRangeRepository"
Private Const TABLE_TEN_NUMBERRANGE As String = "ten_numberrange"

Private Const FIELD_DOCUMENT_TYPE_CODE As String = "document_type_code"
Private Const FIELD_FISCAL_YEAR As String = "fiscal_year"
Private Const FIELD_PREFIX As String = "prefix"
Private Const FIELD_CURRENT_VALUE As String = "current_value"
Private Const FIELD_FORMAT_MASK As String = "format_mask"
Private Const FIELD_IS_ACTIVE As String = "is_active"

Public Function GetCurrentNumberValue(ByVal DocumentTypeCode As String, ByVal FiscalYear As Long) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If Not CanReadNumberRanges() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_TEN_NUMBERRANGE & "];", dbOpenDynaset)

    GetCurrentNumberValue = ResolveCurrentValue(rs, DocumentTypeCode, FiscalYear)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    GetCurrentNumberValue = 0
    modErrorHandler.HandleError MODULE_NAME, "GetCurrentNumberValue", Err
    Resume CleanExit
End Function

Public Function IncrementNumberValue(ByVal DocumentTypeCode As String, ByVal FiscalYear As Long) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim nextValue As Long

    If Not CanReadNumberRanges() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_TEN_NUMBERRANGE & "];", dbOpenDynaset)

    If Not FindNumberRangeRow(rs, DocumentTypeCode, FiscalYear) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".IncrementNumberValue", _
            "No active number range found for DocumentType='" & UCase$(Trim$(DocumentTypeCode)) & "', FiscalYear=" & CStr(FiscalYear) & "."
        GoTo CleanExit
    End If

    nextValue = modDaoHelper.NzLong(rs.Fields(FIELD_CURRENT_VALUE).Value, 0) + 1
    rs.Edit
    rs.Fields(FIELD_CURRENT_VALUE).Value = nextValue
    rs.Update

    IncrementNumberValue = nextValue

    modLoggingHandler.LogInfo MODULE_NAME & ".IncrementNumberValue", _
        "Incremented number range for DocumentType='" & UCase$(Trim$(DocumentTypeCode)) & "', FiscalYear=" & CStr(FiscalYear) & " to " & CStr(nextValue) & "."

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    IncrementNumberValue = 0
    modErrorHandler.HandleError MODULE_NAME, "IncrementNumberValue", Err
    Resume CleanExit
End Function

Public Function NumberRangeExists(ByVal DocumentTypeCode As String, ByVal FiscalYear As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If Not CanReadNumberRanges() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_TEN_NUMBERRANGE & "];", dbOpenSnapshot)

    NumberRangeExists = FindNumberRangeRow(rs, DocumentTypeCode, FiscalYear)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    NumberRangeExists = False
    modErrorHandler.HandleError MODULE_NAME, "NumberRangeExists", Err
    Resume CleanExit
End Function

Private Function CanReadNumberRanges() As Boolean
    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CanReadNumberRanges", _
            "Backend configuration is not ready for number range lookup."
        Exit Function
    End If

    If Not TableExists(TABLE_TEN_NUMBERRANGE) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CanReadNumberRanges", _
            "Table '" & TABLE_TEN_NUMBERRANGE & "' is not available yet."
        Exit Function
    End If

    CanReadNumberRanges = True
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

Private Function FindNumberRangeRow(ByVal rs As DAO.Recordset, ByVal DocumentTypeCode As String, ByVal FiscalYear As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim targetType As String

    targetType = UCase$(Trim$(DocumentTypeCode))
    If LenB(targetType) = 0 Then
        Exit Function
    End If

    If rs.BOF And rs.EOF Then
        Exit Function
    End If

    If Not modDaoHelper.RecordsetHasField(rs, FIELD_DOCUMENT_TYPE_CODE) _
        Or Not modDaoHelper.RecordsetHasField(rs, FIELD_FISCAL_YEAR) _
        Or Not modDaoHelper.RecordsetHasField(rs, FIELD_CURRENT_VALUE) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".FindNumberRangeRow", _
            "Required number range fields are not available in table '" & TABLE_TEN_NUMBERRANGE & "'."
        Exit Function
    End If

    rs.MoveFirst
    Do Until rs.EOF
        If UCase$(Trim$(modDaoHelper.NzString(rs.Fields(FIELD_DOCUMENT_TYPE_CODE).Value))) = targetType _
            And modDaoHelper.NzLong(rs.Fields(FIELD_FISCAL_YEAR).Value, 0) = FiscalYear Then

            If modDaoHelper.RecordsetHasField(rs, FIELD_IS_ACTIVE) Then
                If Not modDaoHelper.NzBoolean(rs.Fields(FIELD_IS_ACTIVE).Value, False) Then
                    rs.MoveNext
                    GoTo ContinueLoop
                End If
            End If

            FindNumberRangeRow = True
            Exit Function
        End If

ContinueLoop:
        rs.MoveNext
    Loop
    Exit Function

ErrorHandler:
    FindNumberRangeRow = False
    modErrorHandler.HandleError MODULE_NAME, "FindNumberRangeRow", Err
End Function

Private Function ResolveCurrentValue(ByVal rs As DAO.Recordset, ByVal DocumentTypeCode As String, ByVal FiscalYear As Long) As Long
    On Error GoTo ErrorHandler

    If FindNumberRangeRow(rs, DocumentTypeCode, FiscalYear) Then
        ResolveCurrentValue = modDaoHelper.NzLong(rs.Fields(FIELD_CURRENT_VALUE).Value, 0)
    End If
    Exit Function

ErrorHandler:
    ResolveCurrentValue = 0
    modErrorHandler.HandleError MODULE_NAME, "ResolveCurrentValue", Err
End Function
