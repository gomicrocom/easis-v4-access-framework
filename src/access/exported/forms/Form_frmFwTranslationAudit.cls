VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFwTranslationAudit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmFwTranslationAudit
' Purpose   : Workspace-aware administrator audit view for translation
'             completeness and registry-based audit follow-up.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmFwTranslationAudit"
Private Const AUDIT_TABLE As String = "tmp_fw_translation_audit"
Private Const TRANSLATION_FORM_NAME As String = "frmFwTranslationEdit"
Private Const SCOPE_ALL As String = "(ALL)"
Private Const LANGUAGE_ALL As String = "(ALL)"
Private Const STATUS_ALL As String = "(ALL)"
Private Const EMPTY_RESULTS_TEXT As String = "Keine Eintraege gefunden."
Private Const STATUS_MISSING_ROW As String = "MISSING_ROW"
Private Const STATUS_MISSING_ROWS_LEGACY As String = "MISSING_ROWS"

Private m_currentSearchText As String

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    modFormRuntime.InitializeForm Me
    modFwTranslationRuntime.ApplyTranslations Me

    PopulateFilterCombos
    ResetFilters
    ConfigureAuditDataSource
    RefreshAuditSummary

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "Translation audit workspace loaded."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.FW_TRANSLATION_AUDIT_LOAD_ERROR", "Fehler beim Laden des Uebersetzungs-Audits."), vbExclamation, MODULE_NAME
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Current", _
        "Entered Form_Current; HasCurrentAuditRow=" & CStr(HasCurrentAuditRow())

    If Not HasVisibleAuditRows() Then
        ClearAuditDetail
        UpdateActionState
        Exit Sub
    End If

    If Not HasCurrentAuditRow() Then
        ClearAuditDetail
        UpdateActionState
        Exit Sub
    End If

    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub cboScopeCode_AfterUpdate()
    On Error GoTo ErrorHandler

    ApplyAuditFilter
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboScopeCode_AfterUpdate", Err
End Sub

Private Sub cboLanguageCode_AfterUpdate()
    On Error GoTo ErrorHandler

    ApplyAuditFilter
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboLanguageCode_AfterUpdate", Err
End Sub

Private Sub cboAuditStatus_AfterUpdate()
    On Error GoTo ErrorHandler

    Dim hasRows As Boolean

    modLoggingHandler.LogInfo MODULE_NAME & ".cboAuditStatus_AfterUpdate", _
        "Selected audit status raw value='" & modDaoHelper.NzString(GetFilterValue("cboAuditStatus"), vbNullString) & "'."

    hasRows = ApplyAuditFilter()
    If Not hasRows Then
        Exit Sub
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboAuditStatus_AfterUpdate", Err
End Sub

Private Sub cmdRefreshAudit_Click()
    On Error GoTo ErrorHandler

    RefreshAuditData
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdRefreshAudit_Click", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.FW_TRANSLATION_AUDIT_REFRESH_ERROR", "Fehler beim Aktualisieren des Uebersetzungs-Audits."), vbExclamation, MODULE_NAME
End Sub

Private Sub cmdCreateMissingRows_Click()
    On Error GoTo ErrorHandler

    modFwTranslationAuditService.EnsureMissingTranslationRows
    RefreshAuditData
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdCreateMissingRows_Click", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.FW_TRANSLATION_AUDIT_CREATE_MISSING_ERROR", "Fehler beim Erzeugen fehlender Uebersetzungseintraege."), vbExclamation, MODULE_NAME
End Sub

Private Sub cmdOpenTranslation_Click()
    On Error GoTo ErrorHandler

    OpenSelectedTranslation
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdOpenTranslation_Click", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.FW_TRANSLATION_AUDIT_OPEN_ERROR", "Fehler beim Oeffnen der Uebersetzung."), vbExclamation, MODULE_NAME
End Sub

Private Sub cmdClearFilters_Click()
    On Error GoTo ErrorHandler

    ResetFilters
    ApplyAuditFilter
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdClearFilters_Click", Err
End Sub

Public Function SupportsListCommandBar() As Boolean
    SupportsListCommandBar = True
End Function

Public Function SupportsWorkspaceCommandBar() As Boolean
    SupportsWorkspaceCommandBar = True
End Function

Public Function GetWorkspaceCommandBarConfig() As Object
    On Error GoTo ErrorHandler

    Dim cfg As Object
    Dim shellForm As Access.Form
    Dim hasShellHost As Boolean

    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = vbTextCompare
    Set shellForm = ResolveHostingShellForm()
    hasShellHost = Not shellForm Is Nothing

    cfg("search.visible") = True
    cfg("search.enabled") = True
    cfg("search.value") = m_currentSearchText
    cfg("search.command_change") = modAppShell.WCMD_LIST_SEARCH
    cfg("search.command_clear") = modAppShell.WCMD_LIST_CLEAR_SEARCH

    cfg(modAppShell.WSLOT_L1 & ".visible") = True
    cfg(modAppShell.WSLOT_L1 & ".enabled") = hasShellHost
    cfg(modAppShell.WSLOT_L1 & ".caption") = modFwTranslationRuntime.ResolveText("COMMON.HOME", "Home")
    cfg(modAppShell.WSLOT_L1 & ".command") = modAppShell.WCMD_NAV_HOME

    cfg(modAppShell.WSLOT_L2 & ".visible") = True
    cfg(modAppShell.WSLOT_L2 & ".enabled") = hasShellHost And modAppWorkspaceService.CanGoBack()
    cfg(modAppShell.WSLOT_L2 & ".caption") = modFwTranslationRuntime.ResolveText("COMMON.BACK", "Zurueck")
    cfg(modAppShell.WSLOT_L2 & ".command") = modAppShell.WCMD_NAV_BACK

    cfg(modAppShell.WSLOT_L3 & ".visible") = False
    cfg(modAppShell.WSLOT_L3 & ".enabled") = False
    cfg(modAppShell.WSLOT_L3 & ".caption") = vbNullString
    cfg(modAppShell.WSLOT_L3 & ".command") = vbNullString

    cfg(modAppShell.WSLOT_R3 & ".visible") = True
    cfg(modAppShell.WSLOT_R3 & ".enabled") = True
    cfg(modAppShell.WSLOT_R3 & ".caption") = modFwTranslationRuntime.ResolveText("FORM.FRMFWTRANSLATIONAUDIT.CLEAR_FILTERS", "Filter loeschen")
    cfg(modAppShell.WSLOT_R3 & ".command") = modAppShell.WCMD_LIST_CLEAR_SEARCH

    cfg(modAppShell.WSLOT_R2 & ".visible") = True
    cfg(modAppShell.WSLOT_R2 & ".enabled") = HasCurrentAuditRow()
    cfg(modAppShell.WSLOT_R2 & ".caption") = modFwTranslationRuntime.ResolveText("COMMON.EDIT", "Bearbeiten")
    cfg(modAppShell.WSLOT_R2 & ".command") = modAppShell.WCMD_LIST_EDIT

    cfg(modAppShell.WSLOT_R1 & ".visible") = True
    cfg(modAppShell.WSLOT_R1 & ".enabled") = True
    cfg(modAppShell.WSLOT_R1 & ".caption") = modFwTranslationRuntime.ResolveText("COMMON.REFRESH", "Aktualisieren")
    cfg(modAppShell.WSLOT_R1 & ".command") = modAppShell.WCMD_LIST_REFRESH

    Set GetWorkspaceCommandBarConfig = cfg
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetWorkspaceCommandBarConfig", Err
End Function

Public Function CanExecuteWorkspaceCommand(ByVal commandKey As String) As Boolean
    On Error GoTo ErrorHandler

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()

    Select Case UCase$(Trim$(commandKey))
        Case modAppShell.WCMD_NAV_HOME
            CanExecuteWorkspaceCommand = Not shellForm Is Nothing
        Case modAppShell.WCMD_NAV_BACK
            CanExecuteWorkspaceCommand = (Not shellForm Is Nothing) And modAppWorkspaceService.CanGoBack()
        Case modAppShell.WCMD_LIST_SEARCH
            CanExecuteWorkspaceCommand = True
        Case modAppShell.WCMD_LIST_CLEAR_SEARCH
            CanExecuteWorkspaceCommand = True
        Case modAppShell.WCMD_LIST_EDIT
            CanExecuteWorkspaceCommand = HasCurrentAuditRow()
        Case modAppShell.WCMD_LIST_REFRESH
            CanExecuteWorkspaceCommand = True
        Case Else
            CanExecuteWorkspaceCommand = False
    End Select
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CanExecuteWorkspaceCommand", Err
End Function

Public Sub ExecuteWorkspaceCommand(ByVal commandKey As String, Optional ByVal commandValue As String = "")
    On Error GoTo ErrorHandler

    Select Case UCase$(Trim$(commandKey))
        Case modAppShell.WCMD_NAV_HOME
            NavigateHome
        Case modAppShell.WCMD_NAV_BACK
            NavigateBack
        Case modAppShell.WCMD_LIST_SEARCH
            ListSearch commandValue
        Case modAppShell.WCMD_LIST_CLEAR_SEARCH
            ClearAllFilters
        Case modAppShell.WCMD_LIST_EDIT
            OpenSelectedTranslation
        Case modAppShell.WCMD_LIST_REFRESH
            RefreshAuditData
    End Select
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ExecuteWorkspaceCommand", Err
End Sub

Public Function SupportsListNew() As Boolean
    SupportsListNew = False
End Function

Public Function SupportsListEdit() As Boolean
    SupportsListEdit = HasCurrentAuditRow()
End Function

Public Sub ListSearch(ByVal searchText As String)
    m_currentSearchText = Trim$(modDaoHelper.NzString(searchText))
    SyncLocalSearchControl
    ApplyAuditFilter
End Sub

Public Sub ListClearSearch()
    m_currentSearchText = vbNullString
    SyncLocalSearchControl
    ApplyAuditFilter
End Sub

Public Sub ListNew()
    modLoggingHandler.LogInfo MODULE_NAME & ".ListNew", _
        "Shell command bar New is not supported for translation audit."
End Sub

Public Sub ListEdit()
    OpenSelectedTranslation
End Sub

Public Sub ListRefresh()
    RefreshAuditData
End Sub

Public Function GetCurrentWorkspaceRecordId() As Long
    GetCurrentWorkspaceRecordId = GetCurrentAuditId()
End Function

Public Function GetWorkspaceState() As String
    Dim stateText As String

    stateText = AddWorkspaceStatePair(stateText, "scope", modDaoHelper.NzString(GetFilterValue("cboScopeCode"), SCOPE_ALL))
    stateText = AddWorkspaceStatePair(stateText, "language", modDaoHelper.NzString(GetFilterValue("cboLanguageCode"), LANGUAGE_ALL))
    stateText = AddWorkspaceStatePair(stateText, "status", modDaoHelper.NzString(GetFilterValue("cboAuditStatus"), STATUS_ALL))
    stateText = AddWorkspaceStatePair(stateText, "search", m_currentSearchText)
    stateText = AddWorkspaceStatePair(stateText, "audit_id", CStr(GetCurrentAuditId()))
    stateText = AddWorkspaceStatePair(stateText, "translation_key", GetCurrentTranslationKey())
    stateText = AddWorkspaceStatePair(stateText, "row_language", GetCurrentRowLanguageCode())

    GetWorkspaceState = stateText
End Function

Public Sub RestoreWorkspaceState(ByVal stateText As String)
    On Error GoTo ErrorHandler

    Dim auditId As Long
    Dim translationKey As String
    Dim rowLanguageCode As String

    SetFilterValue "cboScopeCode", GetWorkspaceStateValue(stateText, "scope"), SCOPE_ALL
    SetFilterValue "cboLanguageCode", GetWorkspaceStateValue(stateText, "language"), LANGUAGE_ALL
    SetFilterValue "cboAuditStatus", GetWorkspaceStateValue(stateText, "status"), STATUS_ALL

    m_currentSearchText = GetWorkspaceStateValue(stateText, "search")
    SyncLocalSearchControl

    If Not ApplyAuditFilter() Then
        Exit Sub
    End If

    translationKey = GetWorkspaceStateValue(stateText, "translation_key")
    rowLanguageCode = GetWorkspaceStateValue(stateText, "row_language")
    If LenB(translationKey) > 0 Then
        RestoreCurrentTranslationSelection translationKey, rowLanguageCode
        Exit Sub
    End If

    auditId = modDaoHelper.NzLong(GetWorkspaceStateValue(stateText, "audit_id"), 0)
    If auditId > 0 Then
        RestoreCurrentAuditSelection auditId
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreWorkspaceState", Err
End Sub

Private Sub ConfigureAuditDataSource()
    On Error GoTo ErrorHandler

    modFwTranslationAuditService.EnsureTranslationAuditTable
    Me.RecordSource = BuildAuditRecordSourceSql(vbNullString)
    Me.OrderBy = BuildAuditOrderBy()
    Me.OrderByOn = True
    Me.Filter = vbNullString
    Me.FilterOn = False
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureAuditDataSource", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub RefreshAuditData()
    On Error GoTo ErrorHandler

    Dim selectedTranslationKey As String
    Dim selectedLanguageCode As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    selectedTranslationKey = GetCurrentTranslationKey()
    selectedLanguageCode = GetCurrentRowLanguageCode()

    Me.FilterOn = False
    Me.Filter = vbNullString

    modFwTranslationAuditService.BuildTranslationAuditData
    If Not ApplyAuditFilter() Then
        RefreshAuditSummary
        UpdateActionState
        Exit Sub
    End If
    If LenB(selectedTranslationKey) > 0 And HasVisibleAuditRows() Then
        RestoreCurrentTranslationSelection selectedTranslationKey, selectedLanguageCode
    End If
    RefreshAuditSummary
    UpdateActionState

    modLoggingHandler.LogInfo MODULE_NAME & ".RefreshAuditData", _
        "Translation audit data rebuilt and requeried."
    Exit Sub

ErrorHandler:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    modErrorHandler.HandleError MODULE_NAME, "RefreshAuditData", Err
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Private Sub RefreshAuditSummary()
    On Error GoTo ErrorHandler

    If HasControl("txtCoverageSummary") Then
        If HasVisibleAuditRows() Then
            Me!txtCoverageSummary.Value = modFwTranslationAuditService.GetTranslationCoverageSummary()
        Else
            Me!txtCoverageSummary.Value = EMPTY_RESULTS_TEXT & " | " & modFwTranslationAuditService.GetTranslationCoverageSummary()
        End If
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RefreshAuditSummary", Err
End Sub

Private Sub PopulateFilterCombos()
    On Error GoTo ErrorHandler

    PopulateScopeCombo
    PopulateLanguageCombo
    PopulateStatusCombo
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "PopulateFilterCombos", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub PopulateScopeCombo()
    If Not HasControl("cboScopeCode") Then
        Exit Sub
    End If

    Me!cboScopeCode.RowSourceType = "Value List"
    Me!cboScopeCode.RowSource = BuildValueList( _
        SCOPE_ALL, modFwTranslationRuntime.ResolveText("COMMON.ALL", "Alle"), _
        "FORM", "FORM", _
        "NAV", "NAV", _
        "REF", "REF", _
        "MSG", "MSG", _
        "STATUS", "STATUS", _
        "COMMON", "COMMON")
    Me!cboScopeCode.BoundColumn = 1
    Me!cboScopeCode.ColumnCount = 2
    Me!cboScopeCode.ColumnWidths = "0cm;3cm"
End Sub

Private Sub PopulateLanguageCombo()
    If Not HasControl("cboLanguageCode") Then
        Exit Sub
    End If

    Me!cboLanguageCode.RowSourceType = "Value List"
    Me!cboLanguageCode.RowSource = BuildValueList( _
        LANGUAGE_ALL, modFwTranslationRuntime.ResolveText("COMMON.ALL", "Alle"), _
        "DE-CH", "DE-CH", _
        "EN-US", "EN-US", _
        "FR-FR", "FR-FR", _
        "DE", "DE", _
        "EN", "EN")
    Me!cboLanguageCode.BoundColumn = 1
    Me!cboLanguageCode.ColumnCount = 2
    Me!cboLanguageCode.ColumnWidths = "0cm;3cm"
End Sub

Private Sub PopulateStatusCombo()
    If Not HasControl("cboAuditStatus") Then
        Exit Sub
    End If

    Me!cboAuditStatus.RowSourceType = "Value List"
    Me!cboAuditStatus.RowSource = BuildValueList( _
        STATUS_ALL, modFwTranslationRuntime.ResolveText("COMMON.ALL", "Alle"), _
        "OK", "OK", _
        "MISSING_ROW", "MISSING_ROW", _
        "EMPTY_VALUE", "EMPTY_VALUE", _
        "ORPHAN", "ORPHAN", _
        "LEGACY_KEY", "LEGACY_KEY")
    Me!cboAuditStatus.BoundColumn = 1
    Me!cboAuditStatus.ColumnCount = 2
    Me!cboAuditStatus.ColumnWidths = "0cm;3.5cm"
End Sub

Private Sub ResetFilters()
    m_currentSearchText = vbNullString
    SetFilterValue "cboScopeCode", SCOPE_ALL, SCOPE_ALL
    SetFilterValue "cboLanguageCode", LANGUAGE_ALL, LANGUAGE_ALL
    SetFilterValue "cboAuditStatus", STATUS_ALL, STATUS_ALL
    SyncLocalSearchControl
End Sub

Private Sub ClearAllFilters()
    ResetFilters
    ApplyAuditFilter
End Sub

Private Function ApplyAuditFilter() As Boolean
    On Error GoTo ErrorHandler

    Dim whereClause As String
    Dim scopeValue As String
    Dim languageValue As String
    Dim statusValue As String
    Dim searchText As String
    Dim rowCount As Long
    Dim sqlStatement As String

    scopeValue = modDaoHelper.NzString(GetFilterValue("cboScopeCode"), SCOPE_ALL)
    languageValue = modDaoHelper.NzString(GetFilterValue("cboLanguageCode"), LANGUAGE_ALL)
    statusValue = NormalizeAuditStatusValue(modDaoHelper.NzString(GetFilterValue("cboAuditStatus"), STATUS_ALL))
    searchText = Trim$(m_currentSearchText)

    If HasControl("cboAuditStatus") Then
        If StrComp(modDaoHelper.NzString(Me!cboAuditStatus.Value, STATUS_ALL), statusValue, vbTextCompare) <> 0 Then
            Me!cboAuditStatus.Value = statusValue
        End If
    End If

    If StrComp(scopeValue, SCOPE_ALL, vbTextCompare) <> 0 Then
        whereClause = AppendFilterClause(whereClause, "[scope_code] = " & SqlText(scopeValue))
    End If

    If StrComp(languageValue, LANGUAGE_ALL, vbTextCompare) <> 0 Then
        whereClause = AppendFilterClause(whereClause, "[language_code] = " & SqlText(languageValue))
    End If

    If StrComp(statusValue, STATUS_ALL, vbTextCompare) <> 0 Then
        whereClause = AppendFilterClause(whereClause, "[audit_status] = " & SqlText(statusValue))
    End If

    If LenB(searchText) > 0 Then
        whereClause = AppendFilterClause(whereClause, BuildSearchFilter(searchText))
    End If

    Me.FilterOn = False
    Me.Filter = vbNullString

    rowCount = CountAuditRows(whereClause)
    sqlStatement = BuildAuditRecordSourceSql(whereClause)

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyAuditFilter", _
        "audit_status='" & statusValue & "'; where_clause='" & whereClause & "'; row_count=" & CStr(rowCount) & "; pending_record_source=" & sqlStatement

    If rowCount <= 0 Then
        MsgBox "Keine Eintraege fuer den gewaehlten Filter gefunden. Es werden wieder alle Audit-Eintraege angezeigt.", vbInformation, MODULE_NAME
        ResetFilters
        Me.FilterOn = False
        Me.Filter = vbNullString
        Me.RecordSource = BuildAuditRecordSourceSql(vbNullString)
        Me.Requery
        ClearAuditDetail
        RefreshAuditSummary
        UpdateActionState
        ApplyAuditFilter = False
        modLoggingHandler.LogInfo MODULE_NAME & ".ApplyAuditFilter", _
            "Zero-match filter result handled by resetting all filters and reloading the full audit list."
        Exit Function
    End If

    Me.FilterOn = False
    Me.Filter = vbNullString
    Me.RecordSource = sqlStatement
    Me.Requery
    Me.OrderBy = BuildAuditOrderBy()
    Me.OrderByOn = True

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyAuditFilter", _
        "Post-recordsource visible_row_count=" & CStr(GetVisibleAuditRowCount()) & "; HasCurrentAuditRow=" & CStr(HasCurrentAuditRow())

    If Not HasCurrentAuditRow() Then
        ClearAuditDetail
        UpdateActionState
        Exit Function
    End If

    RefreshAuditSummary
    UpdateActionState
    ApplyAuditFilter = True
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyAuditFilter", Err
End Function

Private Sub OpenSelectedTranslation()
    On Error GoTo ErrorHandler

    Dim translationKey As String

    translationKey = GetCurrentTranslationKey()
    If LenB(translationKey) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.FW_TRANSLATION_AUDIT_SELECT_FIRST", "Bitte zuerst einen Uebersetzungseintrag auswaehlen."), vbExclamation, MODULE_NAME
        Exit Sub
    End If

    If Not modAppWorkspaceService.OpenWorkspaceForm(Me, TRANSLATION_FORM_NAME, , acFormPropertySettings, translationKey) Then
        Err.Raise vbObjectError + 5400, MODULE_NAME & ".OpenSelectedTranslation", "Workspace navigation failed."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "OpenSelectedTranslation", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub UpdateActionState()
    On Error GoTo SafeExit

    If HasControl("cmdOpenTranslation") Then
        Me!cmdOpenTranslation.Enabled = HasCurrentAuditRow()
    End If

    If HasControl("cmdCreateMissingRows") Then
        Me!cmdCreateMissingRows.Enabled = HasCurrentAuditRow()
    End If

    If HasControl("cmdRefreshAudit") Then
        Me!cmdRefreshAudit.Enabled = True
    End If

    RefreshShellStatusIfHosted

SafeExit:
End Sub

Private Function GetCurrentAuditId() As Long
    On Error GoTo SafeExit

    If Not HasCurrentAuditRow() Then
        Exit Function
    End If

    If HasControl("audit_id") Then
        GetCurrentAuditId = modDaoHelper.NzLong(Me!audit_id.Value, 0)
    End If

SafeExit:
End Function

Private Function GetCurrentTranslationKey() As String
    On Error GoTo SafeExit

    If Not HasCurrentAuditRow() Then
        Exit Function
    End If

    If HasControl("translation_key") Then
        GetCurrentTranslationKey = Trim$(modDaoHelper.NzString(Me!translation_key.Value))
    End If

SafeExit:
End Function

Private Function GetCurrentRowLanguageCode() As String
    On Error GoTo SafeExit

    If Not HasCurrentAuditRow() Then
        Exit Function
    End If

    If HasControl("language_code") Then
        GetCurrentRowLanguageCode = Trim$(modDaoHelper.NzString(Me!language_code.Value))
    End If

SafeExit:
End Function

Private Sub SyncLocalSearchControl()
    On Error GoTo SafeExit

    If HasControl("txtSearch") Then
        Me!txtSearch.Value = NullIfEmpty(m_currentSearchText)
    End If

SafeExit:
End Sub

Private Sub RefreshShellStatusIfHosted()
    On Error GoTo SafeExit

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()
    If Not shellForm Is Nothing Then
        modAppShell.RefreshShellStatus shellForm
    End If

SafeExit:
End Sub

Private Sub NavigateHome()
    On Error GoTo ErrorHandler

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()
    If shellForm Is Nothing Then
        Exit Sub
    End If

    modAppShell.LoadDefaultWorkspace shellForm
    modAppShell.RefreshShellStatus shellForm
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NavigateHome", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub NavigateBack()
    On Error GoTo ErrorHandler

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()
    If shellForm Is Nothing Then
        Exit Sub
    End If

    If modAppWorkspaceService.GoBack(shellForm) Then
        modAppShell.RefreshShellStatus shellForm
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NavigateBack", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function ResolveHostingShellForm() As Access.Form
    On Error GoTo SafeExit

    Dim parentForm As Access.Form

    Set parentForm = Me.Parent
    If parentForm Is Nothing Then
        Exit Function
    End If

    If HasControlOnForm(parentForm, "subWorkspaceHost") Then
        Set ResolveHostingShellForm = parentForm
    End If

SafeExit:
End Function

Private Function GetFilterValue(ByVal controlName As String) As Variant
    On Error GoTo SafeExit

    If HasControl(controlName) Then
        GetFilterValue = Me.Controls(controlName).Value
    End If

SafeExit:
End Function

Private Sub SetFilterValue(ByVal controlName As String, ByVal filterValue As String, ByVal defaultValue As String)
    On Error GoTo SafeExit

    If Not HasControl(controlName) Then
        Exit Sub
    End If

    If StrComp(controlName, "cboAuditStatus", vbTextCompare) = 0 Then
        filterValue = NormalizeAuditStatusValue(filterValue)
    End If

    If LenB(Trim$(modDaoHelper.NzString(filterValue))) = 0 Then
        Me.Controls(controlName).Value = defaultValue
    Else
        Me.Controls(controlName).Value = filterValue
    End If

SafeExit:
End Sub

Private Sub RestoreCurrentAuditSelection(ByVal auditId As Long)
    On Error GoTo ErrorHandler

    Dim recordsetClone As DAO.Recordset

    If auditId <= 0 Then
        Exit Sub
    End If

    Set recordsetClone = Me.RecordsetClone
    If Not RecordsetHasRows(recordsetClone) Then
        GoTo CleanExit
    End If

    recordsetClone.FindFirst "[audit_id]=" & CStr(auditId)
    If Not recordsetClone.NoMatch Then
        Me.Bookmark = recordsetClone.Bookmark
    End If

CleanExit:
    On Error Resume Next
    If Not recordsetClone Is Nothing Then recordsetClone.Close
    Set recordsetClone = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreCurrentAuditSelection", Err
    Resume CleanExit
End Sub

Private Sub RestoreCurrentTranslationSelection(ByVal translationKey As String, ByVal languageCode As String)
    On Error GoTo ErrorHandler

    Dim recordsetClone As DAO.Recordset
    Dim criteria As String

    translationKey = Replace(Trim$(modDaoHelper.NzString(translationKey)), "'", "''")
    languageCode = Replace(Trim$(modDaoHelper.NzString(languageCode)), "'", "''")
    If LenB(translationKey) = 0 Then
        Exit Sub
    End If

    If Not HasVisibleAuditRows() Then
        Exit Sub
    End If

    criteria = "[translation_key]='" & translationKey & "'"
    If LenB(languageCode) > 0 Then
        criteria = criteria & " AND [language_code]='" & languageCode & "'"
    End If

    Set recordsetClone = Me.RecordsetClone
    If Not RecordsetHasRows(recordsetClone) Then
        GoTo CleanExit
    End If

    recordsetClone.FindFirst criteria
    If Not recordsetClone.NoMatch Then
        Me.Bookmark = recordsetClone.Bookmark
    End If

CleanExit:
    On Error Resume Next
    If Not recordsetClone Is Nothing Then recordsetClone.Close
    Set recordsetClone = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreCurrentTranslationSelection", Err
    Resume CleanExit
End Sub

Private Function BuildAuditOrderBy() As String
    BuildAuditOrderBy = RemoveTrailingComma( _
        "IIf([audit_status]='ORPHAN',1," & _
        "IIf([audit_status]='LEGACY_KEY',2," & _
        "IIf([audit_status]='MISSING_ROW',3," & _
        "IIf([audit_status]='EMPTY_VALUE',4,5))))," & _
        " [scope_code], [translation_key], [language_code]")
End Function

Private Function BuildAuditRecordSourceSql(ByVal whereClause As String) As String
    BuildAuditRecordSourceSql = _
        "SELECT audit_id, audit_status, language_code, scope_code, translation_key, " & _
        "source_type, source_object, source_control, translation_value, fallback_text, " & _
        "exists_in_fw_translation, created_at " & _
        "FROM " & AUDIT_TABLE

    If LenB(Trim$(whereClause)) > 0 Then
        BuildAuditRecordSourceSql = BuildAuditRecordSourceSql & " WHERE " & whereClause
    End If

    BuildAuditRecordSourceSql = BuildAuditRecordSourceSql & " ORDER BY " & BuildAuditOrderBy() & ";"
End Function

Private Function BuildEmptyAuditRecordSourceSql() As String
    BuildEmptyAuditRecordSourceSql = vbNullString
End Function

Private Function RemoveTrailingComma(ByVal sqlFragment As String) As String
    sqlFragment = RTrim$(sqlFragment)
    If Right$(sqlFragment, 1) = "," Then
        sqlFragment = Left$(sqlFragment, Len(sqlFragment) - 1)
    End If
    RemoveTrailingComma = sqlFragment
End Function

Private Function CountAuditRows(ByVal whereClause As String) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlStatement As String

    sqlStatement = "SELECT Count(*) AS row_count FROM " & AUDIT_TABLE
    If LenB(Trim$(whereClause)) > 0 Then
        sqlStatement = sqlStatement & " WHERE " & whereClause
    End If
    sqlStatement = sqlStatement & ";"

    Set db = CurrentDb
    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then
        CountAuditRows = modDaoHelper.NzLong(rs.Fields("row_count").Value, 0)
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CountAuditRows", Err
    Resume CleanExit
End Function

Private Function BuildValueList(ParamArray items() As Variant) As String
    Dim itemIndex As Long
    Dim valueList As String

    For itemIndex = LBound(items) To UBound(items)
        If LenB(valueList) > 0 Then
            valueList = valueList & ";"
        End If
        valueList = valueList & Replace(modDaoHelper.NzString(items(itemIndex)), ";", ",")
    Next itemIndex

    BuildValueList = valueList
End Function

Private Function AppendFilterClause(ByVal filterExpression As String, ByVal clauseText As String) As String
    If LenB(filterExpression) > 0 Then
        filterExpression = filterExpression & " AND "
    End If

    AppendFilterClause = filterExpression & clauseText
End Function

Private Function BuildSearchFilter(ByVal searchText As String) As String
    Dim likeValue As String

    likeValue = SqlText("*" & Replace(searchText, "'", "''") & "*")

    BuildSearchFilter = "(" & _
        "Nz([translation_key],'') Like " & likeValue & " OR " & _
        "Nz([translation_value],'') Like " & likeValue & " OR " & _
        "Nz([fallback_text],'') Like " & likeValue & " OR " & _
        "Nz([source_object],'') Like " & likeValue & " OR " & _
        "Nz([source_control],'') Like " & likeValue & ")"
End Function

Private Function HasControl(ByVal controlName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentControl As Control

    For Each currentControl In Me.Controls
        If StrComp(currentControl.Name, controlName, vbTextCompare) = 0 Then
            HasControl = True
            Exit Function
        End If
    Next currentControl

SafeExit:
End Function

Private Function HasControlOnForm(ByVal formInstance As Access.Form, ByVal controlName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentControl As Control

    If formInstance Is Nothing Then
        Exit Function
    End If

    For Each currentControl In formInstance.Controls
        If StrComp(currentControl.Name, controlName, vbTextCompare) = 0 Then
            HasControlOnForm = True
            Exit Function
        End If
    Next currentControl

SafeExit:
End Function

Private Function AddWorkspaceStatePair(ByVal stateText As String, ByVal keyName As String, ByVal valueText As String) As String
    If LenB(stateText) > 0 Then
        stateText = stateText & ";"
    End If

    AddWorkspaceStatePair = stateText & keyName & "=" & EscapeWorkspaceStateValue(valueText)
End Function

Private Function EscapeWorkspaceStateValue(ByVal valueText As String) As String
    valueText = Replace(modDaoHelper.NzString(valueText), "%", "%25")
    valueText = Replace(valueText, ";", "%3B")
    valueText = Replace(valueText, "=", "%3D")
    EscapeWorkspaceStateValue = valueText
End Function

Private Function GetWorkspaceStateValue(ByVal stateText As String, ByVal keyName As String) As String
    Dim parts() As String
    Dim pairText As Variant
    Dim separatorPosition As Long
    Dim currentKey As String

    If LenB(stateText) = 0 Then
        Exit Function
    End If

    parts = Split(stateText, ";")
    For Each pairText In parts
        separatorPosition = InStr(1, CStr(pairText), "=", vbBinaryCompare)
        If separatorPosition > 0 Then
            currentKey = Left$(CStr(pairText), separatorPosition - 1)
            If StrComp(currentKey, keyName, vbTextCompare) = 0 Then
                GetWorkspaceStateValue = UnescapeWorkspaceStateValue(Mid$(CStr(pairText), separatorPosition + 1))
                Exit Function
            End If
        End If
    Next pairText
End Function

Private Function UnescapeWorkspaceStateValue(ByVal valueText As String) As String
    valueText = Replace(valueText, "%3D", "=")
    valueText = Replace(valueText, "%3B", ";")
    valueText = Replace(valueText, "%25", "%")
    UnescapeWorkspaceStateValue = valueText
End Function

Private Function NullIfEmpty(ByVal valueText As String) As Variant
    valueText = modDaoHelper.NzString(valueText)
    If LenB(valueText) = 0 Then
        NullIfEmpty = Null
    Else
        NullIfEmpty = valueText
    End If
End Function

Private Sub ClearAuditDetail()
    On Error GoTo SafeExit

    If HasControl("cmdOpenTranslation") Then
        Me!cmdOpenTranslation.Enabled = False
    End If

    If HasControl("cmdCreateMissingRows") Then
        Me!cmdCreateMissingRows.Enabled = False
    End If

SafeExit:
End Sub

Private Function HasCurrentAuditRow() As Boolean
    On Error GoTo SafeExit

    If Me.NewRecord Then
        Exit Function
    End If

    If Me.Recordset Is Nothing Then
        Exit Function
    End If

    If Me.Recordset.BOF And Me.Recordset.EOF Then
        Exit Function
    End If

    If LenB(GetCurrentBoundFieldValue("audit_status")) = 0 And LenB(GetCurrentBoundFieldValue("translation_key")) = 0 Then
        Exit Function
    End If

    HasCurrentAuditRow = True

SafeExit:
End Function

Private Function HasVisibleAuditRows() As Boolean
    On Error GoTo SafeExit

    Dim recordsetClone As DAO.Recordset

    Set recordsetClone = Me.RecordsetClone
    HasVisibleAuditRows = RecordsetHasRows(recordsetClone)

SafeExit:
    On Error Resume Next
    If Not recordsetClone Is Nothing Then recordsetClone.Close
    Set recordsetClone = Nothing
End Function

Private Function RecordsetHasRows(ByVal sourceRecordset As DAO.Recordset) As Boolean
    On Error GoTo SafeExit

    If sourceRecordset Is Nothing Then
        Exit Function
    End If

    RecordsetHasRows = Not (sourceRecordset.BOF And sourceRecordset.EOF)

SafeExit:
End Function

Private Function GetVisibleAuditRowCount() As Long
    On Error GoTo SafeExit

    Dim recordsetClone As DAO.Recordset

    Set recordsetClone = Me.RecordsetClone
    If RecordsetHasRows(recordsetClone) Then
        recordsetClone.MoveLast
        GetVisibleAuditRowCount = recordsetClone.RecordCount
    End If

SafeExit:
    On Error Resume Next
    If Not recordsetClone Is Nothing Then recordsetClone.Close
    Set recordsetClone = Nothing
End Function

Private Function GetCurrentBoundFieldValue(ByVal controlName As String) As String
    On Error GoTo SafeExit

    If HasControl(controlName) Then
        GetCurrentBoundFieldValue = Trim$(modDaoHelper.NzString(Me.Controls(controlName).Value))
    End If

SafeExit:
End Function

Private Function NormalizeAuditStatusValue(ByVal statusValue As String) As String
    statusValue = UCase$(Trim$(modDaoHelper.NzString(statusValue)))

    Select Case statusValue
        Case STATUS_MISSING_ROWS_LEGACY
            NormalizeAuditStatusValue = STATUS_MISSING_ROW
        Case vbNullString
            NormalizeAuditStatusValue = STATUS_ALL
        Case Else
            NormalizeAuditStatusValue = statusValue
    End Select
End Function
