Attribute VB_Name = "modBackendLinker"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modBackendLinker
' Purpose   : Relinks linked Access backend tables to the active tenant backend.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modBackendLinker"
Private Const ACCESS_CONNECT_PREFIX As String = ";DATABASE="

Public Function RelinkBackendTables() As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim BackendPath As String
    Dim relinkedCount As Long

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogError MODULE_NAME & ".RelinkBackendTables", _
            "Backend relink aborted because backend configuration is invalid."
        Exit Function
    End If

    BackendPath = GetBackendPath()
    Set db = GetCurrentDatabase()

    For Each tdf In db.TableDefs
        If ShouldSkipTable(tdf.Name) Then
            GoTo NextTable
        End If

        If IsLinkedAccessTable(tdf) Then
            If RelinkTable(tdf, BackendPath) Then
                relinkedCount = relinkedCount + 1
            End If
        End If

NextTable:
    Next tdf

    modLoggingHandler.LogInfo MODULE_NAME & ".RelinkBackendTables", _
        "Backend relink completed. Relinked tables: " & CStr(relinkedCount) & "."

    RelinkBackendTables = True
    Exit Function

ErrorHandler:
    RelinkBackendTables = False
    modErrorHandler.HandleError MODULE_NAME, "RelinkBackendTables", Err
End Function

Public Function GetLinkedTableCount() As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = GetCurrentDatabase()

    For Each tdf In db.TableDefs
        If Not ShouldSkipTable(tdf.Name) Then
            If IsLinkedAccessTable(tdf) Then
                GetLinkedTableCount = GetLinkedTableCount + 1
            End If
        End If
    Next tdf

    Exit Function

ErrorHandler:
    GetLinkedTableCount = 0
    modErrorHandler.HandleError MODULE_NAME, "GetLinkedTableCount", Err
End Function

Private Function IsLinkedAccessTable(ByVal TableDef As DAO.TableDef) As Boolean
    Dim connectText As String

    connectText = Trim$(Nz(TableDef.Connect, vbNullString))
    If LenB(connectText) = 0 Then
        Exit Function
    End If

    IsLinkedAccessTable = (InStr(1, connectText, ACCESS_CONNECT_PREFIX, vbTextCompare) > 0)
End Function

Private Function RelinkTable(ByVal TableDef As DAO.TableDef, ByVal BackendPath As String) As Boolean
    On Error GoTo ErrorHandler

    TableDef.Connect = ACCESS_CONNECT_PREFIX & BackendPath
    TableDef.RefreshLink

    modLoggingHandler.LogInfo MODULE_NAME & ".RelinkTable", _
        "Relinked table '" & TableDef.Name & "' to '" & BackendPath & "'."

    RelinkTable = True
    Exit Function

ErrorHandler:
    RelinkTable = False
    modLoggingHandler.LogError MODULE_NAME & ".RelinkTable", _
        "Failed to relink table '" & TableDef.Name & "' to '" & BackendPath & "'.", Err.Number
End Function

Private Function ShouldSkipTable(ByVal TableName As String) As Boolean
    Dim normalizedName As String

    normalizedName = UCase$(Trim$(TableName))

    If LenB(normalizedName) = 0 Then
        ShouldSkipTable = True
        Exit Function
    End If

    If Left$(normalizedName, 4) = "MSYS" Then
        ShouldSkipTable = True
        Exit Function
    End If

    If Left$(normalizedName, 1) = "~" Then
        ShouldSkipTable = True
        Exit Function
    End If

    If Left$(normalizedName, 4) = "TMP_" Or Left$(normalizedName, 5) = "TEMP_" Then
        ShouldSkipTable = True
    End If
End Function
