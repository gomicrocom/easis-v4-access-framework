Attribute VB_Name = "modAppWorkspaceService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modAppWorkspaceService
' Purpose   : Workspace host helpers and history support for the application shell.
' Author    : Codex
' Version   : 0.4.0
'===============================================================================

Private Const MODULE_NAME As String = "modAppWorkspaceService"
Private Const WORKSPACE_SUBFORM_CONTROL As String = "subWorkspaceHost"
Private Const DASHBOARD_FORM_NAME As String = "frmAppDashboard"

Private Const HISTORY_KEY_FORM_NAME As String = "form_name"
Private Const HISTORY_KEY_WHERE_CONDITION As String = "where_condition"
Private Const HISTORY_KEY_FILTER As String = "filter"
Private Const HISTORY_KEY_ORDER_BY As String = "order_by"
Private Const HISTORY_KEY_CURRENT_RECORD_ID As String = "current_record_id"
Private Const HISTORY_KEY_WORKSPACE_STATE As String = "workspace_state"
Private Const HISTORY_KEY_OPEN_ARGS As String = "open_args"
Private Const HISTORY_KEY_DATA_MODE As String = "data_mode"
Private Const HISTORY_KEY_TIMESTAMP As String = "timestamp"

Private m_workspaceHistory As Collection
Private m_isRestoringHistory As Boolean

Public Function OpenWorkspaceForm( _
    ByVal shellForm As Access.Form, _
    ByVal form_name As String, _
    Optional ByVal where_condition As String = "", _
    Optional ByVal data_mode As AcFormOpenDataMode = acFormPropertySettings, _
    Optional ByVal open_args As String = "", _
    Optional ByVal track_history As Boolean = True) As Boolean
    On Error GoTo ErrorHandler

    Dim hostForm As Access.Form
    Dim workspaceHost As Control
    Dim loadedWorkspaceForm As Access.Form
    Dim historyItemText As String
    Dim previousSourceObject As String
    Dim previousWorkspaceState As String
    Dim targetFormName As String
    Dim targetSourceObject As String
    Dim currentSourceObject As String
    Dim loadSucceeded As Boolean

    targetFormName = Trim$(form_name)
    If LenB(targetFormName) = 0 Then
        Exit Function
    End If

    targetSourceObject = "Form." & targetFormName
    modFwDiagnostics.LogSystemSnapshot "BeforeOpenWorkspaceForm:" & targetFormName

    If Not FormExists(targetFormName) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".OpenWorkspaceForm", _
            "Workspace form '" & targetFormName & "' was not found."
        Exit Function
    End If

    Set hostForm = ResolveWorkspaceHostForm(shellForm)

    If hostForm Is Nothing Then
        DoCmd.OpenForm targetFormName, acNormal, , where_condition, data_mode, , open_args
        OpenWorkspaceForm = True
        modLoggingHandler.LogInfo MODULE_NAME & ".OpenWorkspaceForm", _
            "Opened form '" & targetFormName & "' without shell host."
        Exit Function
    End If

    Set workspaceHost = GetWorkspaceHostControl(hostForm)
    If workspaceHost Is Nothing Then
        DoCmd.OpenForm targetFormName, acNormal, , where_condition, data_mode, , open_args
        OpenWorkspaceForm = True
        modLoggingHandler.LogWarning MODULE_NAME & ".OpenWorkspaceForm", _
            "Shell workspace host is missing. Opened form '" & targetFormName & "' standalone."
        Exit Function
    End If

    If Not CanReplaceWorkspaceContent(workspaceHost) Then
        modLoggingHandler.LogInfo MODULE_NAME & ".OpenWorkspaceForm", _
            "Workspace navigation to '" & targetFormName & "' was cancelled by the active workspace form."
        Exit Function
    End If

    If track_history And Not m_isRestoringHistory Then
        historyItemText = CaptureCurrentWorkspaceHistory(workspaceHost)
    End If

    previousSourceObject = Trim$(Nz(workspaceHost.SourceObject, vbNullString))
    previousWorkspaceState = historyItemText

    workspaceHost.SourceObject = vbNullString
    workspaceHost.SourceObject = targetSourceObject

    Set loadedWorkspaceForm = TryGetHostedWorkspaceForm(workspaceHost)
    If loadedWorkspaceForm Is Nothing Then
        Err.Raise 2467, MODULE_NAME & ".OpenWorkspaceForm", _
            "Workspace target form '" & targetFormName & "' is not available after SourceObject switch."
    End If

    If StrComp(loadedWorkspaceForm.Name, targetFormName, vbTextCompare) <> 0 Then
        Err.Raise 2467, MODULE_NAME & ".OpenWorkspaceForm", _
            "Loaded workspace form mismatch. Expected '" & targetFormName & "', got '" & loadedWorkspaceForm.Name & "'."
    End If

    ApplyWorkspaceFormState loadedWorkspaceForm, where_condition, data_mode

    If LenB(Trim$(open_args)) > 0 Then
        ApplyWorkspaceOpenArgs loadedWorkspaceForm, open_args, targetFormName
    End If

    SetWorkspaceFocus workspaceHost

    If LenB(historyItemText) > 0 Then
        AppendHistoryItem historyItemText
    End If

    modAppShell.RefreshShellStatus hostForm
    modFwDiagnostics.LogWorkspaceState "AfterOpenWorkspaceForm:" & targetFormName, hostForm

    loadSucceeded = True
    OpenWorkspaceForm = True
    modLoggingHandler.LogInfo MODULE_NAME & ".OpenWorkspaceForm", _
        "Loaded form '" & targetFormName & "' into the shell workspace."
    Exit Function

ErrorHandler:
    OpenWorkspaceForm = False
    currentSourceObject = ResolveCurrentWorkspaceSourceObject(workspaceHost)
    modLoggingHandler.LogWarning MODULE_NAME & ".OpenWorkspaceForm", _
        "Workspace load failed. target_form_name='" & targetFormName & _
        "'; previous_source_object='" & previousSourceObject & _
        "'; current_source_object='" & currentSourceObject & _
        "'; err_number=" & CStr(Err.Number) & _
        "; err_description='" & Replace(Err.Description, "'", "''") & "'."
    If Not loadSucceeded Then
        RecoverWorkspaceAfterLoadFailure hostForm, workspaceHost, previousSourceObject, previousWorkspaceState, targetFormName
    End If
    modErrorHandler.HandleError MODULE_NAME, "OpenWorkspaceForm", Err
End Function

Public Function PushWorkspaceState(ByVal workspaceForm As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    Dim historyItemText As String

    historyItemText = SerializeWorkspaceHistoryItem(workspaceForm)
    If LenB(historyItemText) = 0 Then
        PushWorkspaceState = True
        Exit Function
    End If

    AppendHistoryItem historyItemText
    PushWorkspaceState = True
    Exit Function

ErrorHandler:
    PushWorkspaceState = False
    modErrorHandler.HandleError MODULE_NAME, "PushWorkspaceState", Err
End Function

Public Function GoBack(ByVal shellForm As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    Dim historyItemText As String
    Dim FormName As String
    Dim whereCondition As String
    Dim openArgs As String
    Dim hostForm As Access.Form
    Dim workspaceHost As Control
    Dim dataMode As AcFormOpenDataMode

    If Not CanGoBack() Then
        Exit Function
    End If

    historyItemText = PopHistoryItem()
    If LenB(historyItemText) = 0 Then
        Exit Function
    End If

    FormName = GetHistoryValue(historyItemText, HISTORY_KEY_FORM_NAME)
    whereCondition = GetHistoryValue(historyItemText, HISTORY_KEY_WHERE_CONDITION)
    openArgs = GetHistoryValue(historyItemText, HISTORY_KEY_OPEN_ARGS)
    dataMode = ResolveHistoryDataMode(historyItemText)

    If LenB(Trim$(FormName)) = 0 Then
        Exit Function
    End If

    m_isRestoringHistory = True

    If Not OpenWorkspaceForm(shellForm, FormName, whereCondition, dataMode, openArgs, False) Then
        AppendHistoryItem historyItemText, True
        GoBack = False
        GoTo CleanExit
    End If

    Set hostForm = ResolveWorkspaceHostForm(shellForm)
    Set workspaceHost = GetWorkspaceHostControl(hostForm)

    If Not workspaceHost Is Nothing Then
        RestoreWorkspaceHistoryState TryGetHostedWorkspaceForm(workspaceHost), historyItemText
    End If

    modAppShell.RefreshShellStatus hostForm

    GoBack = True

    modLoggingHandler.LogInfo MODULE_NAME & ".GoBack", _
        "Workspace history restored form '" & FormName & "'."

CleanExit:
    m_isRestoringHistory = False
    Exit Function

ErrorHandler:
    m_isRestoringHistory = False
    GoBack = False
    modErrorHandler.HandleError MODULE_NAME, "GoBack", Err
End Function

Public Function CanGoBack() As Boolean
    EnsureWorkspaceHistory
    CanGoBack = (m_workspaceHistory.count > 0)
End Function

Public Sub ClearWorkspaceHistory()
    Set m_workspaceHistory = New Collection
End Sub

Public Function PeekWorkspaceHistory() As String
    EnsureWorkspaceHistory

    If m_workspaceHistory.count > 0 Then
        PeekWorkspaceHistory = m_workspaceHistory(m_workspaceHistory.count)
    End If
End Function

Private Sub ApplyWorkspaceOpenArgs(ByVal workspaceForm As Access.Form, ByVal openArgs As String, ByVal formName As String)
    On Error GoTo ErrorHandler

    Dim formObject As Object

    If workspaceForm Is Nothing Then
        Exit Sub
    End If

    Set formObject = workspaceForm
    CallByName formObject, "ApplyWorkspaceOpenArgs", VbMethod, openArgs

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyWorkspaceOpenArgs", _
        "Applied workspace OpenArgs to '" & formName & "'."
    Exit Sub

ErrorHandler:
    If Err.Number = 438 Then
        modLoggingHandler.LogInfo MODULE_NAME & ".ApplyWorkspaceOpenArgs", _
            "Workspace form '" & formName & "' does not implement ApplyWorkspaceOpenArgs."
    Else
        modErrorHandler.HandleError MODULE_NAME, "ApplyWorkspaceOpenArgs", Err
    End If
End Sub

Private Function CanReplaceWorkspaceContent(ByVal workspaceHost As Control) As Boolean
    On Error GoTo SafeExit

    Dim workspaceForm As Access.Form
    Dim formObject As Object

    CanReplaceWorkspaceContent = True

    If workspaceHost Is Nothing Then
        Exit Function
    End If

    If LenB(Trim$(Nz(workspaceHost.SourceObject, vbNullString))) = 0 Then
        Exit Function
    End If

    Set workspaceForm = TryGetHostedWorkspaceForm(workspaceHost)
    If workspaceForm Is Nothing Then
        Exit Function
    End If

    Set formObject = workspaceForm
    CanReplaceWorkspaceContent = CBool(CallByName(formObject, "CanLeaveWorkspace", VbMethod))

SafeExit:
End Function

Public Function PreviewWorkspaceReport( _
    ByVal shellForm As Access.Form, _
    ByVal report_name As String, _
    Optional ByVal where_condition As String = "") As Boolean
    On Error GoTo ErrorHandler

    If LenB(Trim$(report_name)) = 0 Then
        Exit Function
    End If

    If Not ReportExists(report_name) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".PreviewWorkspaceReport", _
            "Workspace report '" & report_name & "' was not found."
        Exit Function
    End If

    DoCmd.OpenReport report_name, acViewPreview, , where_condition

    PreviewWorkspaceReport = True
    modLoggingHandler.LogInfo MODULE_NAME & ".PreviewWorkspaceReport", _
        "Opened report '" & report_name & "' in preview mode."
    Exit Function

ErrorHandler:
    PreviewWorkspaceReport = False
    modErrorHandler.HandleError MODULE_NAME, "PreviewWorkspaceReport", Err
End Function

Public Function LoadDashboard(ByVal shellForm As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    LoadDashboard = OpenWorkspaceForm(shellForm, DASHBOARD_FORM_NAME)
    Exit Function

ErrorHandler:
    LoadDashboard = False
    modErrorHandler.HandleError MODULE_NAME, "LoadDashboard", Err
End Function

Public Function ClearWorkspace(ByVal shellForm As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    Dim hostForm As Access.Form
    Dim workspaceHost As Control

    Set hostForm = ResolveWorkspaceHostForm(shellForm)

    If hostForm Is Nothing Then
        ClearWorkspace = True
        Exit Function
    End If

    Set workspaceHost = GetWorkspaceHostControl(hostForm)
    If workspaceHost Is Nothing Then
        ClearWorkspace = True
        Exit Function
    End If

    workspaceHost.SourceObject = vbNullString
    modAppShell.RefreshShellStatus hostForm
    ClearWorkspace = True

    modLoggingHandler.LogInfo MODULE_NAME & ".ClearWorkspace", _
        "Workspace content cleared."
    Exit Function

ErrorHandler:
    ClearWorkspace = False
    modErrorHandler.HandleError MODULE_NAME, "ClearWorkspace", Err
End Function

Private Sub ApplyWorkspaceFormState( _
    ByVal targetForm As Access.Form, _
    ByVal where_condition As String, _
    ByVal data_mode As AcFormOpenDataMode)
    On Error GoTo ErrorHandler

    If targetForm Is Nothing Then
        Exit Sub
    End If

    ' Reset state first.
    targetForm.FilterOn = False
    targetForm.Filter = vbNullString
    targetForm.DataEntry = False

    Select Case data_mode
        Case acFormAdd
            targetForm.AllowAdditions = True
            targetForm.DataEntry = True
            targetForm.SetFocus
            DoCmd.GoToRecord , , acNewRec

        Case Else
            If LenB(Trim$(where_condition)) > 0 Then
                targetForm.Filter = where_condition
                targetForm.FilterOn = True
                targetForm.Requery
            End If
    End Select

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyWorkspaceFormState", Err
End Sub

Private Sub AppendHistoryItem( _
    ByVal historyItemText As String, _
    Optional ByVal allowDuplicateTopItem As Boolean = False)
    EnsureWorkspaceHistory

    If LenB(historyItemText) = 0 Then
        Exit Sub
    End If

    If Not allowDuplicateTopItem Then
        If m_workspaceHistory.count > 0 Then
            If StrComp(m_workspaceHistory(m_workspaceHistory.count), historyItemText, vbBinaryCompare) = 0 Then
                Exit Sub
            End If
        End If
    End If

    m_workspaceHistory.Add historyItemText
End Sub

Private Function CaptureCurrentWorkspaceHistory(ByVal workspaceHost As Control) As String
    On Error GoTo SafeExit

    Dim workspaceForm As Access.Form

    If workspaceHost Is Nothing Then
        Exit Function
    End If

    If LenB(Trim$(Nz(workspaceHost.SourceObject, vbNullString))) = 0 Then
        Exit Function
    End If

    Set workspaceForm = TryGetHostedWorkspaceForm(workspaceHost)
    CaptureCurrentWorkspaceHistory = SerializeWorkspaceHistoryItem(workspaceForm)

SafeExit:
End Function

Private Sub EnsureWorkspaceHistory()
    If m_workspaceHistory Is Nothing Then
        Set m_workspaceHistory = New Collection
    End If
End Sub

Private Function EscapeHistoryValue(ByVal Value As String) As String
    Value = Replace(Value, "%", "%25")
    Value = Replace(Value, ";", "%3B")
    Value = Replace(Value, "=", "%3D")
    Value = Replace(Value, vbCrLf, "%0D%0A")
    Value = Replace(Value, vbCr, "%0D")
    Value = Replace(Value, vbLf, "%0A")
    EscapeHistoryValue = Value
End Function

Private Function FormExists(ByVal form_name As String) As Boolean
    On Error GoTo SafeExit

    Dim accessObject As Access.accessObject

    For Each accessObject In CurrentProject.AllForms
        If StrComp(accessObject.Name, form_name, vbTextCompare) = 0 Then
            FormExists = True
            Exit Function
        End If
    Next accessObject

SafeExit:
End Function

Private Function GetCurrentRecordIdText(ByVal workspaceForm As Access.Form) As String
    On Error GoTo SafeExit

    Dim formObject As Object

    Set formObject = workspaceForm
    GetCurrentRecordIdText = CStr(CallByName(formObject, "GetCurrentWorkspaceRecordId", VbMethod))

SafeExit:
End Function

Private Function GetHistoryValue(ByVal historyItemText As String, ByVal keyName As String) As String
    Dim parts() As String
    Dim pairText As Variant
    Dim separatorPosition As Long
    Dim currentKey As String

    If LenB(historyItemText) = 0 Then
        Exit Function
    End If

    parts = Split(historyItemText, ";")
    For Each pairText In parts
        separatorPosition = InStr(1, CStr(pairText), "=", vbBinaryCompare)
        If separatorPosition > 0 Then
            currentKey = Left$(CStr(pairText), separatorPosition - 1)
            If StrComp(currentKey, keyName, vbTextCompare) = 0 Then
                GetHistoryValue = UnescapeHistoryValue(Mid$(CStr(pairText), separatorPosition + 1))
                Exit Function
            End If
        End If
    Next pairText
End Function

Private Function GetWorkspaceHostControl(ByVal shellForm As Access.Form) As Control
    On Error GoTo SafeExit

    If shellForm Is Nothing Then
        Exit Function
    End If

    If HasControl(shellForm, WORKSPACE_SUBFORM_CONTROL) Then
        Set GetWorkspaceHostControl = shellForm.Controls(WORKSPACE_SUBFORM_CONTROL)
    End If

SafeExit:
End Function

Private Function HasControl(ByVal formInstance As Access.Form, ByVal ControlName As String) As Boolean
    On Error GoTo SafeExit

    Dim ctl As Control

    If formInstance Is Nothing Then
        Exit Function
    End If

    For Each ctl In formInstance.Controls
        If StrComp(ctl.Name, ControlName, vbTextCompare) = 0 Then
            HasControl = True
            Exit Function
        End If
    Next ctl

SafeExit:
End Function

Private Function PopHistoryItem() As String
    EnsureWorkspaceHistory

    If m_workspaceHistory.count = 0 Then
        Exit Function
    End If

    PopHistoryItem = m_workspaceHistory(m_workspaceHistory.count)
    m_workspaceHistory.Remove m_workspaceHistory.count
End Function

Private Function ReportExists(ByVal report_name As String) As Boolean
    On Error GoTo SafeExit

    Dim accessObject As Access.accessObject

    For Each accessObject In CurrentProject.AllReports
        If StrComp(accessObject.Name, report_name, vbTextCompare) = 0 Then
            ReportExists = True
            Exit Function
        End If
    Next accessObject

SafeExit:
End Function

Private Function ResolveHistoryDataMode(ByVal historyItemText As String) As AcFormOpenDataMode
    If StrComp(GetHistoryValue(historyItemText, HISTORY_KEY_DATA_MODE), "ADD", vbTextCompare) = 0 Then
        ResolveHistoryDataMode = acFormAdd
    Else
        ResolveHistoryDataMode = acFormPropertySettings
    End If
End Function

Private Function ResolveWorkspaceHostForm(ByVal shellForm As Access.Form) As Access.Form
    On Error GoTo SafeExit

    If shellForm Is Nothing Then
        Exit Function
    End If

    If HasControl(shellForm, WORKSPACE_SUBFORM_CONTROL) Then
        Set ResolveWorkspaceHostForm = shellForm
        Exit Function
    End If

    Set ResolveWorkspaceHostForm = shellForm.Parent

    If Not ResolveWorkspaceHostForm Is Nothing Then
        If Not HasControl(ResolveWorkspaceHostForm, WORKSPACE_SUBFORM_CONTROL) Then
            Set ResolveWorkspaceHostForm = Nothing
        End If
    End If

SafeExit:
End Function

Private Sub RestoreWorkspaceHistoryState( _
    ByVal workspaceForm As Access.Form, _
    ByVal historyItemText As String)
    On Error GoTo ErrorHandler

    Dim filterText As String
    Dim orderByText As String
    Dim stateText As String

    If workspaceForm Is Nothing Then
        Exit Sub
    End If

    filterText = GetHistoryValue(historyItemText, HISTORY_KEY_FILTER)
    orderByText = GetHistoryValue(historyItemText, HISTORY_KEY_ORDER_BY)
    stateText = GetHistoryValue(historyItemText, HISTORY_KEY_WORKSPACE_STATE)

    If LenB(filterText) > 0 Then
        workspaceForm.Filter = filterText
        workspaceForm.FilterOn = True
    Else
        workspaceForm.FilterOn = False
        workspaceForm.Filter = vbNullString
    End If

    workspaceForm.OrderBy = orderByText
    workspaceForm.OrderByOn = (LenB(orderByText) > 0)

    TryRestoreWorkspaceState workspaceForm, stateText
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreWorkspaceHistoryState", Err
End Sub

Private Function SerializeWorkspaceHistoryItem(ByVal workspaceForm As Access.Form) As String
    On Error GoTo ErrorHandler

    Dim historyItemText As String
    Dim filterText As String
    Dim orderByText As String
    Dim dataModeText As String
    Dim whereConditionText As String

    If workspaceForm Is Nothing Then
        Exit Function
    End If

    If workspaceForm.FilterOn Then
        filterText = Nz(workspaceForm.Filter, vbNullString)
        whereConditionText = filterText
    End If

    If workspaceForm.OrderByOn Then
        orderByText = Nz(workspaceForm.OrderBy, vbNullString)
    End If

    If workspaceForm.DataEntry Then
        dataModeText = "ADD"
    Else
        dataModeText = "NORMAL"
    End If

    historyItemText = AddHistoryPair(historyItemText, HISTORY_KEY_FORM_NAME, workspaceForm.Name)
    historyItemText = AddHistoryPair(historyItemText, HISTORY_KEY_WHERE_CONDITION, whereConditionText)
    historyItemText = AddHistoryPair(historyItemText, HISTORY_KEY_FILTER, filterText)
    historyItemText = AddHistoryPair(historyItemText, HISTORY_KEY_ORDER_BY, orderByText)
    historyItemText = AddHistoryPair(historyItemText, HISTORY_KEY_CURRENT_RECORD_ID, GetCurrentRecordIdText(workspaceForm))
    historyItemText = AddHistoryPair(historyItemText, HISTORY_KEY_WORKSPACE_STATE, TryGetWorkspaceState(workspaceForm))
    historyItemText = AddHistoryPair(historyItemText, HISTORY_KEY_OPEN_ARGS, Nz(workspaceForm.openArgs, vbNullString))
    historyItemText = AddHistoryPair(historyItemText, HISTORY_KEY_DATA_MODE, dataModeText)
    historyItemText = AddHistoryPair(historyItemText, HISTORY_KEY_TIMESTAMP, Format$(Now(), "yyyy-mm-dd hh:nn:ss"))

    SerializeWorkspaceHistoryItem = historyItemText
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SerializeWorkspaceHistoryItem", Err
End Function

Private Sub SetWorkspaceFocus(ByVal workspaceHost As Control)
    On Error GoTo SafeExit

    If workspaceHost Is Nothing Then
        Exit Sub
    End If

    workspaceHost.SetFocus

SafeExit:
End Sub

Private Function TryGetHostedWorkspaceForm(ByVal workspaceHost As Control) As Access.Form
    On Error GoTo SafeExit

    If workspaceHost Is Nothing Then
        Exit Function
    End If

    If LenB(Trim$(Nz(workspaceHost.SourceObject, vbNullString))) = 0 Then
        Exit Function
    End If

    Set TryGetHostedWorkspaceForm = workspaceHost.Form

SafeExit:
End Function

Private Function ResolveCurrentWorkspaceSourceObject(ByVal workspaceHost As Control) As String
    On Error GoTo SafeExit

    If workspaceHost Is Nothing Then
        Exit Function
    End If

    ResolveCurrentWorkspaceSourceObject = Trim$(Nz(workspaceHost.SourceObject, vbNullString))

SafeExit:
End Function

Private Sub RecoverWorkspaceAfterLoadFailure( _
    ByVal hostForm As Access.Form, _
    ByVal workspaceHost As Control, _
    ByVal previousSourceObject As String, _
    ByVal previousWorkspaceState As String, _
    ByVal requestedFormName As String)
    On Error GoTo SafeExit

    Dim restoredWorkspaceForm As Access.Form

    If workspaceHost Is Nothing Then
        Exit Sub
    End If

    If LenB(previousSourceObject) > 0 Then
        workspaceHost.SourceObject = vbNullString
        workspaceHost.SourceObject = previousSourceObject
        Set restoredWorkspaceForm = TryGetHostedWorkspaceForm(workspaceHost)
        If Not restoredWorkspaceForm Is Nothing Then
            If LenB(previousWorkspaceState) > 0 Then
                RestoreWorkspaceHistoryState restoredWorkspaceForm, previousWorkspaceState
            End If
        End If

        modLoggingHandler.LogWarning MODULE_NAME & ".RecoverWorkspaceAfterLoadFailure", _
            "Workspace load failed for '" & requestedFormName & "'. Restored previous SourceObject '" & previousSourceObject & "'."
    ElseIf StrComp(requestedFormName, DASHBOARD_FORM_NAME, vbTextCompare) <> 0 Then
        workspaceHost.SourceObject = vbNullString
        workspaceHost.SourceObject = "Form." & DASHBOARD_FORM_NAME

        modLoggingHandler.LogWarning MODULE_NAME & ".RecoverWorkspaceAfterLoadFailure", _
            "Workspace load failed for '" & requestedFormName & "'. Loaded fallback dashboard."
    Else
        modLoggingHandler.LogWarning MODULE_NAME & ".RecoverWorkspaceAfterLoadFailure", _
            "Workspace load failed for dashboard and no previous SourceObject was available."
    End If

    If Not hostForm Is Nothing Then
        modAppShell.RefreshShellStatus hostForm
        modFwDiagnostics.LogWorkspaceState "OpenWorkspaceFormFailureRecovery:" & requestedFormName, hostForm
    End If

SafeExit:
End Sub

Private Function TryGetWorkspaceState(ByVal workspaceForm As Access.Form) As String
    On Error GoTo SafeExit

    Dim formObject As Object

    Set formObject = workspaceForm
    TryGetWorkspaceState = CStr(CallByName(formObject, "GetWorkspaceState", VbMethod))

SafeExit:
End Function

Private Sub TryRestoreWorkspaceState(ByVal workspaceForm As Access.Form, ByVal stateText As String)
    On Error GoTo RestoreFailure

    Dim formObject As Object

    If workspaceForm Is Nothing Then
        Exit Sub
    End If

    Set formObject = workspaceForm
    CallByName formObject, "RestoreWorkspaceState", VbMethod, stateText
    Exit Sub

RestoreFailure:
    modLoggingHandler.LogWarning MODULE_NAME & ".TryRestoreWorkspaceState", _
        "Custom workspace state restore was skipped for form '" & workspaceForm.Name & "'."
End Sub

Private Function UnescapeHistoryValue(ByVal Value As String) As String
    Value = Replace(Value, "%0D%0A", vbCrLf)
    Value = Replace(Value, "%0D", vbCr)
    Value = Replace(Value, "%0A", vbLf)
    Value = Replace(Value, "%3D", "=")
    Value = Replace(Value, "%3B", ";")
    Value = Replace(Value, "%25", "%")
    UnescapeHistoryValue = Value
End Function

Private Function AddHistoryPair( _
    ByVal historyItemText As String, _
    ByVal keyName As String, _
    ByVal valueText As String) As String

    If LenB(historyItemText) > 0 Then
        historyItemText = historyItemText & ";"
    End If

    AddHistoryPair = historyItemText & keyName & "=" & EscapeHistoryValue(Nz(valueText, vbNullString))
End Function
