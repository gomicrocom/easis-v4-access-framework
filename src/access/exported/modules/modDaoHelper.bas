Option Compare Database
Option Explicit

'===============================================================================
' Module    : modDaoHelper
' Purpose   : Reusable DAO helper functions for safe value handling.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modDaoHelper"

Public Function NzString(ByVal Value As Variant, Optional ByVal DefaultValue As String = "") As String
    On Error GoTo ErrorHandler

    If IsNull(Value) Or IsEmpty(Value) Then
        NzString = DefaultValue
    Else
        NzString = CStr(Value)
    End If
    Exit Function

ErrorHandler:
    NzString = DefaultValue
    modErrorHandler.HandleError MODULE_NAME, "NzString", Err
End Function

Public Function NzLong(ByVal Value As Variant, Optional ByVal DefaultValue As Long = 0) As Long
    On Error GoTo ErrorHandler

    If IsNull(Value) Or IsEmpty(Value) Then
        NzLong = DefaultValue
    ElseIf IsNumeric(Value) Then
        NzLong = CLng(Value)
    Else
        NzLong = DefaultValue
    End If
    Exit Function

ErrorHandler:
    NzLong = DefaultValue
    modErrorHandler.HandleError MODULE_NAME, "NzLong", Err
End Function

Public Function NzBoolean(ByVal Value As Variant, Optional ByVal DefaultValue As Boolean = False) As Boolean
    On Error GoTo ErrorHandler

    If IsNull(Value) Or IsEmpty(Value) Then
        NzBoolean = DefaultValue
    Else
        Select Case UCase$(Trim$(CStr(Value)))
            Case "1", "TRUE", "YES", "Y", "ON"
                NzBoolean = True
            Case "0", "FALSE", "NO", "N", "OFF"
                NzBoolean = False
            Case Else
                NzBoolean = CBool(Value)
        End Select
    End If
    Exit Function

ErrorHandler:
    NzBoolean = DefaultValue
    modErrorHandler.HandleError MODULE_NAME, "NzBoolean", Err
End Function

Public Function RecordsetHasField(ByVal rs As DAO.Recordset, ByVal fieldName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim fld As DAO.Field
    Dim normalizedFieldName As String

    normalizedFieldName = UCase$(Trim$(fieldName))
    If rs Is Nothing Or LenB(normalizedFieldName) = 0 Then
        Exit Function
    End If

    For Each fld In rs.Fields
        If UCase$(fld.Name) = normalizedFieldName Then
            RecordsetHasField = True
            Exit Function
        End If
    Next fld
    Exit Function

ErrorHandler:
    RecordsetHasField = False
    modErrorHandler.HandleError MODULE_NAME, "RecordsetHasField", Err
End Function