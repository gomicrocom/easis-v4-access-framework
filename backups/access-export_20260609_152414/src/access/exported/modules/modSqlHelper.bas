Attribute VB_Name = "modSqlHelper"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modSqlHelper
' Purpose   : Centralized Access SQL literal formatting helpers.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Public Function SqlText(ByVal Value As Variant) As String
    Dim textValue As String

    If IsNull(Value) Or IsEmpty(Value) Then
        SqlText = "Null"
        Exit Function
    End If

    textValue = Trim$(CStr(Value))
    SqlText = "'" & Replace(textValue, "'", "''") & "'"
End Function

Public Function SqlTextNav(ByVal Value As Variant) As String
    Dim textValue As String

    If IsNull(Value) Or IsEmpty(Value) Then
        SqlTextNav = "Null"
        Exit Function
    End If

    ' Important: do not trim.
    ' Navigation display prefixes may intentionally contain leading spaces.
    textValue = CStr(Value)

    SqlTextNav = "'" & Replace(textValue, "'", "''") & "'"
End Function

Public Function SqlNullableText(ByVal Value As Variant) As String
    Dim textValue As String

    If IsNull(Value) Or IsEmpty(Value) Then
        SqlNullableText = "Null"
        Exit Function
    End If

    textValue = Trim$(CStr(Value))
    If LenB(textValue) = 0 Then
        SqlNullableText = "Null"
    Else
        SqlNullableText = SqlText(textValue)
    End If
End Function

Public Function SqlBoolean(ByVal Value As Boolean) As String
    If Value Then
        SqlBoolean = "True"
    Else
        SqlBoolean = "False"
    End If
End Function

Public Function SqlLongOrNull(ByVal Value As Variant) As String
    If IsNull(Value) Or IsEmpty(Value) Then
        SqlLongOrNull = "Null"
    ElseIf Not IsNumeric(Value) Then
        SqlLongOrNull = "Null"
    ElseIf CLng(Value) > 0 Then
        SqlLongOrNull = CStr(CLng(Value))
    Else
        SqlLongOrNull = "Null"
    End If
End Function

Public Function SqlDateTime(ByVal Value As Variant) As String
    If IsNull(Value) Or IsEmpty(Value) Then
        SqlDateTime = "Null"
    ElseIf Not IsDate(Value) Then
        SqlDateTime = "Null"
    Else
        SqlDateTime = "#" & Format$(CDate(Value), "yyyy-mm-dd hh:nn:ss") & "#"
    End If
End Function

Public Function SqlDateOrNull(ByVal Value As Variant) As String
    SqlDateOrNull = SqlDateTime(Value)
End Function
