Attribute VB_Name = "modDbSchema"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modDbSchema
' Purpose   : Shared schema inspection helpers for DAO databases.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modDbSchema"

Public Function TableExists( _
    ByVal db As DAO.Database, _
    ByVal tableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim normalizedTableName As String
    Dim tableDefinition As DAO.TableDef

    If db Is Nothing Then
        Exit Function
    End If

    normalizedTableName = Trim$(tableName)
    If LenB(normalizedTableName) = 0 Then
        Exit Function
    End If

    db.TableDefs.Refresh

    For Each tableDefinition In db.TableDefs
        If StrComp(Trim$(tableDefinition.Name), normalizedTableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit For
        End If
    Next tableDefinition

CleanExit:
    Set tableDefinition = Nothing
    Exit Function

ErrorHandler:
    TableExists = False
    modErrorHandler.HandleError MODULE_NAME, "TableExists", Err
    Resume CleanExit
End Function

Public Function FieldExists( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal fieldName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim normalizedTableName As String
    Dim normalizedFieldName As String
    Dim tableDefinition As DAO.TableDef
    Dim fieldDefinition As DAO.Field

    If db Is Nothing Then
        Exit Function
    End If

    normalizedTableName = Trim$(tableName)
    normalizedFieldName = Trim$(fieldName)

    If LenB(normalizedTableName) = 0 Then
        Exit Function
    End If

    If LenB(normalizedFieldName) = 0 Then
        Exit Function
    End If

    db.TableDefs.Refresh

    For Each tableDefinition In db.TableDefs
        If StrComp(Trim$(tableDefinition.Name), normalizedTableName, vbTextCompare) = 0 Then
            For Each fieldDefinition In tableDefinition.Fields
                If StrComp(Trim$(fieldDefinition.Name), normalizedFieldName, vbTextCompare) = 0 Then
                    FieldExists = True
                    Exit For
                End If
            Next fieldDefinition
            Exit For
        End If
    Next tableDefinition

CleanExit:
    Set fieldDefinition = Nothing
    Set tableDefinition = Nothing
    Exit Function

ErrorHandler:
    FieldExists = False
    modErrorHandler.HandleError MODULE_NAME, "FieldExists", Err
    Resume CleanExit
End Function
