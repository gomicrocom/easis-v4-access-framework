Option Compare Database
Option Explicit

'===============================================================================
' Module    : modAppNavigationService
' Purpose   : Hierarchical shell navigation setup, seeding, and click handling.
' Author    : Codex
' Version   : 0.2.1
'===============================================================================

Private Const MODULE_NAME As String = "modAppNavigationService"
Private Const TABLE_NAVIGATION As String = "fw_navigation"
Private Const TABLE_NAVIGATION_ROLE As String = "fw_navigation_role"
Private Const TYPE_FORM As String = "FORM"
Private Const TYPE_REPORT As String = "REPORT"
Private Const TYPE_ACTION As String = "ACTION"
Private Const TYPE_GROUP As String = "GROUP"
Private Const PREFIX_GROUP_EXPANDED As String = "- "
Private Const PREFIX_GROUP_COLLAPSED As String = "+ "
Private Const PREFIX_CHILD As String = "    "

Public Function EnsureNavigationTables() As Boolean
    On Error GoTo ErrorHandler

    EnsureNavigationTable
    EnsureNavigationRoleTable

    EnsureNavigationTables = True

    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureNavigationTables", _
        "Navigation tables ensured successfully."
    Exit Function

ErrorHandler:
    EnsureNavigationTables = False
    modErrorHandler.HandleError MODULE_NAME, "EnsureNavigationTables", Err
End Function

Public Function SeedDefaultNavigation() As Boolean
    On Error GoTo ErrorHandler

    Dim addressGroupId As Long
    Dim documentGroupId As Long
    Dim frameworkGroupId As Long

    If Not EnsureNavigationTables() Then
        Exit Function
    End If

    addressGroupId = EnsureNavigationEntry( _
        0, "Adressen", "NAV.GROUP.ADDRESSES", "Adressen", vbNullString, TYPE_GROUP, 10, True, True, True)

    EnsureNavigationEntry _
        addressGroupId, "Adressen", "NAV.ADDRESS_LIST", "Adressliste", "frmAddressList", TYPE_FORM, 10, False, True, True

    EnsureNavigationEntry _
        addressGroupId, "Adressen", "NAV.NEW_ADDRESS", "Neue Adresse", "frmAddressDetail", TYPE_FORM, 20, False, True, True

    documentGroupId = EnsureNavigationEntry( _
        0, "Dokumente", "NAV.GROUP.DOCUMENTS", "Dokumente", vbNullString, TYPE_GROUP, 20, False, True, True)

    EnsureNavigationEntry _
        documentGroupId, "Dokumente", "NAV.DOCUMENT_PREVIEW", "Dokumentvorschau", "rpt_document", TYPE_REPORT, 10, False, True, True

    frameworkGroupId = EnsureNavigationEntry( _
        0, "Framework", "NAV.GROUP.FRAMEWORK", "Framework", vbNullString, TYPE_GROUP, 90, False, True, True)

    EnsureNavigationEntry _
        frameworkGroupId, "Framework", "NAV.TRANSLATIONS", "Uebersetzungen", "frmFwTranslationList", TYPE_FORM, 10, False, True, True

    EnsureNavigationEntry _
        frameworkGroupId, "Framework", "NAV.COMPOSER", "Composer", "frmFwComposer", TYPE_FORM, 20, False, True, True

    EnsureNavigationEntry _
        frameworkGroupId, "Framework", "NAV.LOCALISATION", "Lokalisierung", "frmLocalisation", TYPE_FORM, 30, False, True, True

    EnsureNavigationEntry _
        frameworkGroupId, "Framework", "NAV.TAGS", "Tags", "frmTagComposer", TYPE_FORM, 40, False, True, True

    EnsureNavigationEntry _
        frameworkGroupId, "Framework", "NAV.TAG_HELP", "Tag-Hilfe", "frmTagHelp", TYPE_FORM, 50, False, True, True

    SeedDefaultNavigation = True

    modLoggingHandler.LogInfo MODULE_NAME & ".SeedDefaultNavigation", _
        "Default navigation ensured successfully."
    Exit Function

ErrorHandler:
    SeedDefaultNavigation = False
    modErrorHandler.HandleError MODULE_NAME, "SeedDefaultNavigation", Err
End Function

Public Function GetNavigationRowSource(Optional ByVal role_code As String = "") As String
    On Error GoTo ErrorHandler

    GetNavigationRowSource = BuildNavigationRowSource(False, role_code)
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetNavigationRowSource", Err
End Function

Public Function GetVisibleNavigationRowSource(Optional ByVal role_code As String = "") As String
    On Error GoTo ErrorHandler

    GetVisibleNavigationRowSource = BuildNavigationRowSource(True, role_code)
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetVisibleNavigationRowSource", Err
End Function

Public Function ToggleNavigationGroup(ByVal navigation_id As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim currentlyExpanded As Boolean

    If Not IsGroupNavigation(navigation_id) Then
        Exit Function
    End If

    currentlyExpanded = GetNavigationExpandedState(navigation_id)
    Set db = CurrentDb

    ' Wenn die Gruppe bereits offen ist: schließen.
    If currentlyExpanded Then
        db.Execute "UPDATE " & TABLE_NAVIGATION & " " & _
                   "SET is_expanded = False, " & _
                   "updated_at = Now(), " & _
                   "updated_by = 'SYSTEM' " & _
                   "WHERE navigation_id = " & CStr(navigation_id) & ";", dbFailOnError
    Else
        ' Sonst: alle Gruppen schließen und nur diese öffnen.
        db.Execute "UPDATE " & TABLE_NAVIGATION & " " & _
                   "SET is_expanded = False, " & _
                   "updated_at = Now(), " & _
                   "updated_by = 'SYSTEM' " & _
                   "WHERE UCase(Nz(object_type,'')) = 'GROUP';", dbFailOnError

        db.Execute "UPDATE " & TABLE_NAVIGATION & " " & _
                   "SET is_expanded = True, " & _
                   "updated_at = Now(), " & _
                   "updated_by = 'SYSTEM' " & _
                   "WHERE navigation_id = " & CStr(navigation_id) & ";", dbFailOnError
    End If

    ToggleNavigationGroup = True
    Exit Function

ErrorHandler:
    ToggleNavigationGroup = False
    modErrorHandler.HandleError MODULE_NAME, "ToggleNavigationGroup", Err
End Function
Public Function ExpandNavigationGroup(ByVal navigation_id As Long) As Boolean
    On Error GoTo ErrorHandler

    ExpandNavigationGroup = SetNavigationExpandedState(navigation_id, True)
    Exit Function

ErrorHandler:
    ExpandNavigationGroup = False
    modErrorHandler.HandleError MODULE_NAME, "ExpandNavigationGroup", Err
End Function

Public Function CollapseNavigationGroup(ByVal navigation_id As Long) As Boolean
    On Error GoTo ErrorHandler

    CollapseNavigationGroup = SetNavigationExpandedState(navigation_id, False)
    Exit Function

ErrorHandler:
    CollapseNavigationGroup = False
    modErrorHandler.HandleError MODULE_NAME, "CollapseNavigationGroup", Err
End Function

Public Function CollapseAllNavigationGroups() As Boolean
    On Error GoTo ErrorHandler

    CurrentDb.Execute "UPDATE " & TABLE_NAVIGATION & " " & _
                      "SET is_expanded = False, " & _
                      "updated_at = Now(), " & _
                      "updated_by = 'SYSTEM' " & _
                      "WHERE UCase(Nz(object_type,'')) = 'GROUP';", dbFailOnError

    CollapseAllNavigationGroups = True

    modLoggingHandler.LogInfo MODULE_NAME & ".CollapseAllNavigationGroups", _
        "All navigation groups collapsed."
    Exit Function

ErrorHandler:
    CollapseAllNavigationGroups = False
    modErrorHandler.HandleError MODULE_NAME, "CollapseAllNavigationGroups", Err
End Function

Public Function HandleNavigationClick( _
    ByVal shellForm As Access.Form, _
    ByVal navigation_id As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlStatement As String
    Dim objectName As String
    Dim objectType As String

    If navigation_id <= 0 Then
        Exit Function
    End If

    Set db = CurrentDb
    sqlStatement = "SELECT TOP 1 object_name, object_type " & _
                   "FROM " & TABLE_NAVIGATION & " " & _
                   "WHERE navigation_id = " & CStr(navigation_id) & " " & _
                   "AND Nz(is_active, True) = True " & _
                   "AND Nz(is_visible, True) = True;"

    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)

    If rs.BOF And rs.EOF Then
        GoTo CleanExit
    End If

    objectName = Trim$(modDaoHelper.NzString(rs.Fields("object_name").Value))
    objectType = UCase$(Trim$(modDaoHelper.NzString(rs.Fields("object_type").Value)))

    Select Case objectType
        Case TYPE_GROUP
            HandleNavigationClick = ToggleNavigationGroup(navigation_id)

        Case TYPE_FORM
            HandleNavigationClick = modAppWorkspaceService.OpenWorkspaceForm(shellForm, objectName)

        Case TYPE_REPORT
            HandleNavigationClick = modAppWorkspaceService.PreviewWorkspaceReport(shellForm, objectName)

        Case TYPE_ACTION
            modLoggingHandler.LogWarning MODULE_NAME & ".HandleNavigationClick", _
                "ACTION navigation is not implemented yet for '" & objectName & "'."

        Case Else
            modLoggingHandler.LogWarning MODULE_NAME & ".HandleNavigationClick", _
                "Unsupported navigation object_type '" & objectType & "'."
    End Select

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    HandleNavigationClick = False
    modErrorHandler.HandleError MODULE_NAME, "HandleNavigationClick", Err
    Resume CleanExit
End Function

Private Function BuildNavigationRowSource(ByVal onlyVisibleRows As Boolean, ByVal role_code As String) As String
    Dim normalizedRoleCode As String
    Dim sqlStatement As String

    normalizedRoleCode = UCase$(Trim$(role_code))

    sqlStatement = ""
    sqlStatement = sqlStatement & "SELECT "
    sqlStatement = sqlStatement & "n.navigation_id, "
    sqlStatement = sqlStatement & "n.parent_navigation_id, "
    sqlStatement = sqlStatement & "n.fallback_caption AS display_caption, "
    sqlStatement = sqlStatement & "n.fallback_caption, "
    sqlStatement = sqlStatement & "n.object_name, "
    sqlStatement = sqlStatement & "n.object_type, "
    sqlStatement = sqlStatement & "n.icon_key, "
    sqlStatement = sqlStatement & "n.sort_order, "
    sqlStatement = sqlStatement & "n.is_expanded, "
    sqlStatement = sqlStatement & "IIf(UCase(Nz(n.object_type,''))='GROUP', True, False) AS is_group, "
    sqlStatement = sqlStatement & "IIf(n.parent_navigation_id Is Null, 0, 1) AS display_level, "
    sqlStatement = sqlStatement & "IIf(UCase(Nz(n.object_type,''))='GROUP', "
    sqlStatement = sqlStatement & "IIf(Nz(n.is_expanded,False)=True, " & SqlText(PREFIX_GROUP_EXPANDED) & ", " & SqlText(PREFIX_GROUP_COLLAPSED) & ") & n.fallback_caption, "
    sqlStatement = sqlStatement & SqlText(PREFIX_CHILD) & " & n.fallback_caption) AS display_text "
    sqlStatement = sqlStatement & "FROM " & TABLE_NAVIGATION & " AS n "
    sqlStatement = sqlStatement & "LEFT JOIN " & TABLE_NAVIGATION & " AS p "
    sqlStatement = sqlStatement & "ON n.parent_navigation_id = p.navigation_id "

    If LenB(normalizedRoleCode) > 0 Then
        sqlStatement = sqlStatement & "LEFT JOIN " & TABLE_NAVIGATION_ROLE & " AS nr "
        sqlStatement = sqlStatement & "ON n.navigation_id = nr.navigation_id "
    End If

    sqlStatement = sqlStatement & "WHERE Nz(n.is_active, True) = True "
    sqlStatement = sqlStatement & "AND Nz(n.is_visible, True) = True "

    If LenB(normalizedRoleCode) > 0 Then
        sqlStatement = sqlStatement & "AND (nr.role_code = " & SqlText(normalizedRoleCode) & " "
        sqlStatement = sqlStatement & "OR nr.navigation_role_id Is Null) "
        sqlStatement = sqlStatement & "AND (Nz(nr.is_active, True) = True "
        sqlStatement = sqlStatement & "OR nr.navigation_role_id Is Null) "
    End If

    If onlyVisibleRows Then
        sqlStatement = sqlStatement & "AND (n.parent_navigation_id Is Null "
        sqlStatement = sqlStatement & "OR Nz(p.is_expanded, False) = True) "
    End If

    sqlStatement = sqlStatement & "ORDER BY "
    sqlStatement = sqlStatement & "IIf(n.parent_navigation_id Is Null, n.sort_order, p.sort_order), "
    sqlStatement = sqlStatement & "IIf(n.parent_navigation_id Is Null, 0, 1), "
    sqlStatement = sqlStatement & "n.sort_order, "
    sqlStatement = sqlStatement & "n.navigation_id;"

    BuildNavigationRowSource = sqlStatement
End Function

Private Sub EnsureNavigationTable()
    If Not TableExists(TABLE_NAVIGATION) Then
        ExecuteSql "CREATE TABLE " & TABLE_NAVIGATION & " (" & _
                   "navigation_id AUTOINCREMENT CONSTRAINT pk_fw_navigation PRIMARY KEY, " & _
                   "parent_navigation_id LONG, " & _
                   "navigation_group TEXT(100), " & _
                   "caption_key TEXT(150), " & _
                   "fallback_caption TEXT(150), " & _
                   "object_name TEXT(150), " & _
                   "object_type TEXT(30), " & _
                   "icon_key TEXT(100), " & _
                   "sort_order LONG, " & _
                   "is_active YESNO, " & _
                   "is_expanded YESNO, " & _
                   "is_visible YESNO, " & _
                   "created_at DATETIME, " & _
                   "created_by TEXT(100), " & _
                   "updated_at DATETIME, " & _
                   "updated_by TEXT(100));"
    End If

    EnsureFieldExists TABLE_NAVIGATION, "parent_navigation_id", "LONG"
    EnsureFieldExists TABLE_NAVIGATION, "navigation_group", "TEXT(100)"
    EnsureFieldExists TABLE_NAVIGATION, "caption_key", "TEXT(150)"
    EnsureFieldExists TABLE_NAVIGATION, "fallback_caption", "TEXT(150)"
    EnsureFieldExists TABLE_NAVIGATION, "object_name", "TEXT(150)"
    EnsureFieldExists TABLE_NAVIGATION, "object_type", "TEXT(30)"
    EnsureFieldExists TABLE_NAVIGATION, "icon_key", "TEXT(100)"
    EnsureFieldExists TABLE_NAVIGATION, "sort_order", "LONG"
    EnsureFieldExists TABLE_NAVIGATION, "is_active", "YESNO"
    EnsureFieldExists TABLE_NAVIGATION, "is_expanded", "YESNO"
    EnsureFieldExists TABLE_NAVIGATION, "is_visible", "YESNO"
    EnsureFieldExists TABLE_NAVIGATION, "created_at", "DATETIME"
    EnsureFieldExists TABLE_NAVIGATION, "created_by", "TEXT(100)"
    EnsureFieldExists TABLE_NAVIGATION, "updated_at", "DATETIME"
    EnsureFieldExists TABLE_NAVIGATION, "updated_by", "TEXT(100)"

    EnsureIndexExists TABLE_NAVIGATION, "ix_fw_navigation_parent_navigation_id", _
        "CREATE INDEX ix_fw_navigation_parent_navigation_id ON fw_navigation (parent_navigation_id);"

    EnsureIndexExists TABLE_NAVIGATION, "ix_fw_navigation_navigation_group", _
        "CREATE INDEX ix_fw_navigation_navigation_group ON fw_navigation (navigation_group);"

    EnsureIndexExists TABLE_NAVIGATION, "ix_fw_navigation_sort_order", _
        "CREATE INDEX ix_fw_navigation_sort_order ON fw_navigation (sort_order);"

    EnsureIndexExists TABLE_NAVIGATION, "ix_fw_navigation_is_visible", _
        "CREATE INDEX ix_fw_navigation_is_visible ON fw_navigation (is_visible);"

    EnsureIndexExists TABLE_NAVIGATION, "ix_fw_navigation_object_type", _
        "CREATE INDEX ix_fw_navigation_object_type ON fw_navigation (object_type);"
End Sub

Private Sub EnsureNavigationRoleTable()
    If Not TableExists(TABLE_NAVIGATION_ROLE) Then
        ExecuteSql "CREATE TABLE " & TABLE_NAVIGATION_ROLE & " (" & _
                   "navigation_role_id AUTOINCREMENT CONSTRAINT pk_fw_navigation_role PRIMARY KEY, " & _
                   "navigation_id LONG, " & _
                   "role_code TEXT(50), " & _
                   "is_active YESNO, " & _
                   "created_at DATETIME, " & _
                   "created_by TEXT(100), " & _
                   "updated_at DATETIME, " & _
                   "updated_by TEXT(100));"
    End If

    EnsureFieldExists TABLE_NAVIGATION_ROLE, "navigation_id", "LONG"
    EnsureFieldExists TABLE_NAVIGATION_ROLE, "role_code", "TEXT(50)"
    EnsureFieldExists TABLE_NAVIGATION_ROLE, "is_active", "YESNO"
    EnsureFieldExists TABLE_NAVIGATION_ROLE, "created_at", "DATETIME"
    EnsureFieldExists TABLE_NAVIGATION_ROLE, "created_by", "TEXT(100)"
    EnsureFieldExists TABLE_NAVIGATION_ROLE, "updated_at", "DATETIME"
    EnsureFieldExists TABLE_NAVIGATION_ROLE, "updated_by", "TEXT(100)"

    EnsureIndexExists TABLE_NAVIGATION_ROLE, "ix_fw_navigation_role_navigation_id", _
        "CREATE INDEX ix_fw_navigation_role_navigation_id ON fw_navigation_role (navigation_id);"

    EnsureIndexExists TABLE_NAVIGATION_ROLE, "ix_fw_navigation_role_role_code", _
        "CREATE INDEX ix_fw_navigation_role_role_code ON fw_navigation_role (role_code);"

    EnsureIndexExists TABLE_NAVIGATION_ROLE, "ux_fw_navigation_role_nav_role", _
        "CREATE UNIQUE INDEX ux_fw_navigation_role_nav_role ON fw_navigation_role (navigation_id, role_code);"
End Sub

Private Function EnsureNavigationEntry( _
    ByVal parentNavigationId As Long, _
    ByVal navigationGroup As String, _
    ByVal captionKey As String, _
    ByVal fallbackCaption As String, _
    ByVal objectName As String, _
    ByVal objectType As String, _
    ByVal sortOrder As Long, _
    ByVal isExpanded As Boolean, _
    ByVal isVisible As Boolean, _
    ByVal isActive As Boolean) As Long
    On Error GoTo ErrorHandler

    Dim sqlStatement As String
    Dim existingId As Long

    existingId = LookupNavigationId(captionKey, objectType, objectName, fallbackCaption)

    If existingId <= 0 Then
        sqlStatement = "INSERT INTO " & TABLE_NAVIGATION & " (" & _
                       "parent_navigation_id, navigation_group, caption_key, fallback_caption, " & _
                       "object_name, object_type, icon_key, sort_order, is_active, is_expanded, is_visible, " & _
                       "created_at, created_by, updated_at, updated_by) VALUES (" & _
                       SqlLongOrNull(parentNavigationId) & ", " & _
                       SqlText(navigationGroup) & ", " & _
                       SqlText(captionKey) & ", " & _
                       SqlText(fallbackCaption) & ", " & _
                       SqlNullableText(objectName) & ", " & _
                       SqlText(UCase$(Trim$(objectType))) & ", Null, " & _
                       CStr(sortOrder) & ", " & _
                       SqlBoolean(isActive) & ", " & _
                       SqlBoolean(isExpanded) & ", " & _
                       SqlBoolean(isVisible) & ", " & _
                       "Now(), 'SYSTEM', Now(), 'SYSTEM');"
        ExecuteSql sqlStatement
    Else
        sqlStatement = "UPDATE " & TABLE_NAVIGATION & " SET " & _
                       "parent_navigation_id = " & SqlLongOrNull(parentNavigationId) & ", " & _
                       "navigation_group = " & SqlText(navigationGroup) & ", " & _
                       "caption_key = " & SqlText(captionKey) & ", " & _
                       "fallback_caption = " & SqlText(fallbackCaption) & ", " & _
                       "object_name = " & SqlNullableText(objectName) & ", " & _
                       "object_type = " & SqlText(UCase$(Trim$(objectType))) & ", " & _
                       "sort_order = " & CStr(sortOrder) & ", " & _
                       "is_active = " & SqlBoolean(isActive) & ", " & _
                       "is_expanded = " & SqlBoolean(isExpanded) & ", " & _
                       "is_visible = " & SqlBoolean(isVisible) & ", " & _
                       "updated_at = Now(), " & _
                       "updated_by = 'SYSTEM' " & _
                       "WHERE navigation_id = " & CStr(existingId) & ";"
        ExecuteSql sqlStatement
    End If

    EnsureNavigationEntry = LookupNavigationId(captionKey, objectType, objectName, fallbackCaption)
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureNavigationEntry", Err
End Function

Private Function LookupNavigationId( _
    ByVal captionKey As String, _
    ByVal objectType As String, _
    ByVal objectName As String, _
    ByVal fallbackCaption As String) As Long
    On Error GoTo ErrorHandler

    Dim lookupValue As Variant
    Dim criteria As String

    criteria = BuildNavigationLookupCriteria(captionKey, objectType, objectName)
    lookupValue = DLookup("navigation_id", TABLE_NAVIGATION, criteria)
    LookupNavigationId = modDaoHelper.NzLong(lookupValue, 0)

    If LookupNavigationId > 0 Then
        Exit Function
    End If

    If LenB(Trim$(objectName)) > 0 Then
        criteria = "object_type = " & SqlText(UCase$(Trim$(objectType))) & " " & _
                   "AND object_name = " & SqlText(objectName)
        lookupValue = DLookup("navigation_id", TABLE_NAVIGATION, criteria)
        LookupNavigationId = modDaoHelper.NzLong(lookupValue, 0)
    ElseIf UCase$(Trim$(objectType)) = TYPE_GROUP Then
        criteria = "object_type = " & SqlText(TYPE_GROUP) & " " & _
                   "AND fallback_caption = " & SqlText(fallbackCaption)
        lookupValue = DLookup("navigation_id", TABLE_NAVIGATION, criteria)
        LookupNavigationId = modDaoHelper.NzLong(lookupValue, 0)
    End If

    Exit Function

ErrorHandler:
    LookupNavigationId = 0
    modErrorHandler.HandleError MODULE_NAME, "LookupNavigationId", Err
End Function

Private Function BuildNavigationLookupCriteria( _
    ByVal captionKey As String, _
    ByVal objectType As String, _
    ByVal objectName As String) As String

    BuildNavigationLookupCriteria = "caption_key = " & SqlText(captionKey) & " " & _
                                    "AND object_type = " & SqlText(UCase$(Trim$(objectType))) & " "

    If LenB(Trim$(objectName)) > 0 Then
        BuildNavigationLookupCriteria = BuildNavigationLookupCriteria & _
                                        "AND object_name = " & SqlText(objectName)
    Else
        BuildNavigationLookupCriteria = BuildNavigationLookupCriteria & _
                                        "AND (object_name Is Null OR object_name = '')"
    End If
End Function

Private Function SetNavigationExpandedState(ByVal navigation_id As Long, ByVal expandedState As Boolean) As Boolean
    On Error GoTo ErrorHandler

    If Not IsGroupNavigation(navigation_id) Then
        Exit Function
    End If

    CurrentDb.Execute "UPDATE " & TABLE_NAVIGATION & " " & _
                      "SET is_expanded = " & SqlBoolean(expandedState) & ", " & _
                      "updated_at = Now(), " & _
                      "updated_by = 'SYSTEM' " & _
                      "WHERE navigation_id = " & CStr(navigation_id) & ";", dbFailOnError

    SetNavigationExpandedState = True
    Exit Function

ErrorHandler:
    SetNavigationExpandedState = False
    modErrorHandler.HandleError MODULE_NAME, "SetNavigationExpandedState", Err
End Function

Private Function IsGroupNavigation(ByVal navigation_id As Long) As Boolean
    On Error GoTo ErrorHandler

    IsGroupNavigation = (StrComp( _
        modDaoHelper.NzString(DLookup("object_type", TABLE_NAVIGATION, "navigation_id = " & CStr(navigation_id))), _
        TYPE_GROUP, vbTextCompare) = 0)
    Exit Function

ErrorHandler:
    IsGroupNavigation = False
    modErrorHandler.HandleError MODULE_NAME, "IsGroupNavigation", Err
End Function

Private Function GetNavigationExpandedState(ByVal navigation_id As Long) As Boolean
    On Error GoTo ErrorHandler

    GetNavigationExpandedState = modDaoHelper.NzBoolean( _
        DLookup("is_expanded", TABLE_NAVIGATION, "navigation_id = " & CStr(navigation_id)), False)
    Exit Function

ErrorHandler:
    GetNavigationExpandedState = False
    modErrorHandler.HandleError MODULE_NAME, "GetNavigationExpandedState", Err
End Function

Private Sub EnsureFieldExists(ByVal table_name As String, ByVal field_name As String, ByVal ddlType As String)
    If Not FieldExists(table_name, field_name) Then
        ExecuteSql "ALTER TABLE " & table_name & " ADD COLUMN " & field_name & " " & ddlType & ";"
    End If
End Sub

Private Sub EnsureIndexExists(ByVal table_name As String, ByVal indexName As String, ByVal createSql As String)
    If Not IndexExists(table_name, indexName) Then
        ExecuteSql createSql
    End If
End Sub

Private Function TableExists(ByVal table_name As String) As Boolean
    On Error GoTo SafeExit

    Dim db As DAO.Database
    Dim tdf As DAO.tableDef

    Set db = CurrentDb

    For Each tdf In db.TableDefs
        If StrComp(tdf.Name, table_name, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next tdf

SafeExit:
    Set tdf = Nothing
    Set db = Nothing
End Function

Private Function FieldExists(ByVal table_name As String, ByVal field_name As String) As Boolean
    On Error GoTo SafeExit

    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field

    Set db = CurrentDb
    Set tdf = db.TableDefs(table_name)

    For Each fld In tdf.Fields
        If StrComp(fld.Name, field_name, vbTextCompare) = 0 Then
            FieldExists = True
            Exit Function
        End If
    Next fld

SafeExit:
    Set fld = Nothing
    Set tdf = Nothing
    Set db = Nothing
End Function

Private Function IndexExists(ByVal table_name As String, ByVal indexName As String) As Boolean
    On Error GoTo SafeExit

    Dim db As DAO.Database
    Dim tdf As DAO.tableDef
    Dim idx As DAO.index

    Set db = CurrentDb
    Set tdf = db.TableDefs(table_name)

    For Each idx In tdf.Indexes
        If StrComp(idx.Name, indexName, vbTextCompare) = 0 Then
            IndexExists = True
            Exit Function
        End If
    Next idx

SafeExit:
    Set idx = Nothing
    Set tdf = Nothing
    Set db = Nothing
End Function

Private Sub ExecuteSql(ByVal sqlStatement As String)
    CurrentDb.Execute sqlStatement, dbFailOnError
End Sub

Private Function SqlText(ByVal Value As String) As String
    SqlText = "'" & Replace(Trim$(Value), "'", "''") & "'"
End Function

Private Function SqlNullableText(ByVal Value As String) As String
    If LenB(Trim$(Value)) = 0 Then
        SqlNullableText = "Null"
    Else
        SqlNullableText = SqlText(Value)
    End If
End Function

Private Function SqlLongOrNull(ByVal Value As Long) As String
    If Value > 0 Then
        SqlLongOrNull = CStr(Value)
    Else
        SqlLongOrNull = "Null"
    End If
End Function

Private Function SqlBoolean(ByVal Value As Boolean) As String
    SqlBoolean = IIf(Value, "True", "False")
End Function
