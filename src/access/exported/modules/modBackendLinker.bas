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
    Dim tdf As DAO.tableDef
    Dim systemBackendPath As String
    Dim tenantBackendPath As String
    Dim relinkedCount As Long

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogError MODULE_NAME & ".RelinkBackendTables", _
            "Backend relink aborted because backend configuration is invalid."
        Exit Function
    End If

    systemBackendPath = modDb.GetSystemBackendPath()
    tenantBackendPath = modDb.GetCurrentTenantBackendPath()
    Set db = modDb.GetFrontendDatabase()

    For Each tdf In db.TableDefs
        If ShouldSkipTable(tdf.Name) Then
            GoTo NextTable
        End If

        If IsLinkedAccessTable(tdf) Then
            If RelinkTable(tdf, ResolveBackendPathForTable(tdf.Name, systemBackendPath, tenantBackendPath)) Then
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
    Dim tdf As DAO.tableDef

    Set db = modDb.GetFrontendDatabase()

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

Private Function ResolveBackendPathForTable( _
    ByVal tableName As String, _
    ByVal systemBackendPath As String, _
    ByVal tenantBackendPath As String) As String
    Dim normalizedName As String

    normalizedName = UCase$(Trim$(tableName))
    If LenB(normalizedName) = 0 Then
        Exit Function
    End If

    If IsSystemTableName(normalizedName) Then
        ResolveBackendPathForTable = Trim$(systemBackendPath)
    Else
        ResolveBackendPathForTable = Trim$(tenantBackendPath)
    End If
End Function

Private Function IsSystemTableName(ByVal normalizedTableName As String) As Boolean
    If Left$(normalizedTableName, 3) = "FW_" Then
        IsSystemTableName = True
        Exit Function
    End If

    If Left$(normalizedTableName, 4) = "REF_" Then
        IsSystemTableName = True
        Exit Function
    End If

    Select Case normalizedTableName
        Case "REF_LANGUAGE"
            IsSystemTableName = True
    End Select
End Function

Private Function IsLinkedAccessTable(ByVal tableDef As DAO.tableDef) As Boolean
    Dim connectText As String

    connectText = Trim$(Nz(tableDef.Connect, vbNullString))
    If LenB(connectText) = 0 Then
        Exit Function
    End If

    IsLinkedAccessTable = (InStr(1, connectText, ACCESS_CONNECT_PREFIX, vbTextCompare) > 0)
End Function

Private Function RelinkTable(ByVal tableDef As DAO.tableDef, ByVal backendPath As String) As Boolean
    On Error GoTo ErrorHandler

    If LenB(Trim$(backendPath)) = 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".RelinkTable", _
            "Skipped relink for table '" & tableDef.Name & "' because no backend path was resolved."
        Exit Function
    End If

    tableDef.Connect = ACCESS_CONNECT_PREFIX & backendPath
    tableDef.RefreshLink

    modLoggingHandler.LogInfo MODULE_NAME & ".RelinkTable", _
        "Relinked table '" & tableDef.Name & "' to '" & backendPath & "'."

    RelinkTable = True
    Exit Function

ErrorHandler:
    RelinkTable = False
    modLoggingHandler.LogError MODULE_NAME & ".RelinkTable", _
        "Failed to relink table '" & tableDef.Name & "' to '" & backendPath & "'.", Err.Number
End Function

Private Function ShouldSkipTable(ByVal tableName As String) As Boolean
    Dim normalizedName As String

    normalizedName = UCase$(Trim$(tableName))

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
        Exit Function
    End If

    If normalizedName = "REF_LANGUAGE" Then
        ShouldSkipTable = True
    End If
End Function
