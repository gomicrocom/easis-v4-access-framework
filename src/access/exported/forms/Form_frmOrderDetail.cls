VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmOrderDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ExportKind: CODE_BEHIND_ONLY
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmOrderDetail
' Purpose   : Workspace-aware Sales Order detail form for Phase 1 header editing.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmOrderDetail"
Private Const TABLE_ORD_ORDER As String = "ord_order"
Private Const FIELD_ORDER_ID As String = "order_id"
Private Const DASHBOARD_FORM_NAME As String = "frmAppDashboard"
Private Const ORDER_LINES_SUBFORM_CONTROL As String = "subOrderLines"
Private Const ORDER_LINES_SUBFORM_NAME As String = "sfrmOrderLines"

Private m_workspaceOpenArgsApplied As Boolean

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "OpenArgs raw='" & Replace(modDaoHelper.NzString(Me.OpenArgs), "'", "''") & "'."

    InitializeRecordSource
    ConfigureOrderLinesSubform
    modFwTranslationRuntime.ApplyTranslations Me
    UpdateLocalCommandButtons
    UpdateHeaderCaptionUi

    If LenB(Trim$(modDaoHelper.NzString(Me.OpenArgs))) > 0 Then
        modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
            "Applying standalone OpenArgs during Form_Load."
        ApplyWorkspaceOpenArgs CStr(Me.OpenArgs)
    Else
        modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
            "No OpenArgs available during Form_Load. Waiting for workspace ApplyWorkspaceOpenArgs if hosted."
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    UpdateHeaderCaptionUi
    UpdateLocalCommandButtons
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler

    DetailSave
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdSave_Click", Err
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo ErrorHandler

    DetailCancel
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdCancel_Click", Err
End Sub

Public Function CanLeaveWorkspace() As Boolean
    On Error GoTo ErrorHandler

    If Not HasUnsavedChanges() Then
        CanLeaveWorkspace = True
        Exit Function
    End If

    If MsgBox( _
        modFwTranslationRuntime.ResolveText("MSG.ORDER_DETAIL_CANCEL_CONFIRM", "Aenderungen verwerfen?"), _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        MODULE_NAME) = vbYes Then
        CancelCurrentEdit False, False
        CanLeaveWorkspace = True
    End If
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CanLeaveWorkspace", Err
End Function

Public Function SupportsWorkspaceCommandBar() As Boolean
    SupportsWorkspaceCommandBar = True
End Function

Public Function GetWorkspaceCommandBarConfig() As Object
    Dim cfg As Object
    Dim shellForm As Access.Form
    Dim hasShellHost As Boolean

    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = vbTextCompare
    Set shellForm = ResolveHostingShellForm()
    hasShellHost = Not shellForm Is Nothing

    cfg("search.visible") = False
    cfg("search.enabled") = False

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

    cfg(modAppShell.WSLOT_R2 & ".visible") = True
    cfg(modAppShell.WSLOT_R2 & ".enabled") = True
    cfg(modAppShell.WSLOT_R2 & ".caption") = modFwTranslationRuntime.ResolveText("COMMON.CANCEL", "Abbrechen")
    cfg(modAppShell.WSLOT_R2 & ".command") = modAppShell.WCMD_DETAIL_CANCEL

    cfg(modAppShell.WSLOT_R1 & ".visible") = True
    cfg(modAppShell.WSLOT_R1 & ".enabled") = True
    cfg(modAppShell.WSLOT_R1 & ".caption") = modFwTranslationRuntime.ResolveText("COMMON.SAVE", "Speichern")
    cfg(modAppShell.WSLOT_R1 & ".command") = modAppShell.WCMD_DETAIL_SAVE

    Set GetWorkspaceCommandBarConfig = cfg
End Function

Public Function CanExecuteWorkspaceCommand(ByVal commandKey As String) As Boolean
    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()

    Select Case UCase$(Trim$(commandKey))
        Case modAppShell.WCMD_NAV_HOME
            CanExecuteWorkspaceCommand = Not shellForm Is Nothing
        Case modAppShell.WCMD_NAV_BACK
            CanExecuteWorkspaceCommand = (Not shellForm Is Nothing) And modAppWorkspaceService.CanGoBack()
        Case modAppShell.WCMD_DETAIL_SAVE, modAppShell.WCMD_DETAIL_CANCEL
            CanExecuteWorkspaceCommand = True
    End Select
End Function

Public Sub ExecuteWorkspaceCommand(ByVal commandKey As String, Optional ByVal commandValue As String = "")
    Select Case UCase$(Trim$(commandKey))
        Case modAppShell.WCMD_NAV_HOME
            NavigateHome
        Case modAppShell.WCMD_NAV_BACK
            NavigateBack
        Case modAppShell.WCMD_DETAIL_SAVE
            DetailSave
        Case modAppShell.WCMD_DETAIL_CANCEL
            DetailCancel
    End Select
End Sub

Public Function GetCurrentWorkspaceRecordId() As Long
    GetCurrentWorkspaceRecordId = GetCurrentOrderId()
End Function

Public Function GetCurrentOrderDetailId() As Long
    GetCurrentOrderDetailId = GetCurrentOrderId()
End Function

Public Function GetWorkspaceState() As String
    Dim stateText As String

    stateText = AddWorkspaceStatePair(stateText, "order_id", CStr(GetCurrentOrderId()))
    GetWorkspaceState = stateText
End Function

Public Sub RestoreWorkspaceState(ByVal stateText As String)
    On Error GoTo ErrorHandler

    Dim OrderId As Long

    OrderId = NzLong(GetWorkspaceStateValue(stateText, "order_id"), 0)
    If OrderId > 0 Then
        RestoreCurrentOrderSelection OrderId
    End If

    UpdateHeaderCaptionUi
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreWorkspaceState", Err
End Sub

Public Sub ApplyWorkspaceOpenArgs(ByVal openArgs As String)
    On Error GoTo ErrorHandler

    Dim OrderId As Long

    OrderId = NzLong(openArgs, 0)
    m_workspaceOpenArgsApplied = False

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyWorkspaceOpenArgs", _
        "Parsed order_id=" & CStr(OrderId) & "; raw_open_args='" & Replace(openArgs, "'", "''") & "'."

    If OrderId <= 0 Then
        HandleInvalidOrderContext "MSG.ORDER_DETAIL_INVALID_ID", "Die uebergebene Bestell-ID ist ungueltig."
        Exit Sub
    End If

    If Not modOrderRepository.OrderExists(OrderId) Then
        HandleInvalidOrderContext "MSG.ORDER_DETAIL_NOT_FOUND", "Die Bestellung konnte nicht gefunden werden."
        Exit Sub
    End If

    RestoreCurrentOrderSelection OrderId
    RequeryOrderLinesSubform
    UpdateHeaderCaptionUi
    m_workspaceOpenArgsApplied = True

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyWorkspaceOpenArgs", _
        "Order context applied successfully for order_id=" & CStr(OrderId) & "."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyWorkspaceOpenArgs", Err
End Sub

Public Function HasUnsavedChanges() As Boolean
    On Error GoTo SafeExit

    HasUnsavedChanges = Me.Dirty

SafeExit:
End Function

Public Sub DetailSave()
    SaveCurrentRecord
End Sub

Public Sub DetailCancel()
    CancelCurrentEdit
End Sub

Public Sub RefreshCalculatedTotals(Optional ByVal requeryLines As Boolean = False)
    On Error GoTo ErrorHandler

    Dim OrderId As Long

    OrderId = GetCurrentOrderId()

    Me.Requery
    If OrderId > 0 Then
        RestoreCurrentOrderSelection OrderId
    End If

    If requeryLines Then
        RequeryOrderLinesSubform
    End If

    UpdateHeaderCaptionUi
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RefreshCalculatedTotals", Err
End Sub

Private Sub InitializeRecordSource()
    On Error GoTo ErrorHandler

    If LenB(Trim$(modDaoHelper.NzString(Me.RecordSource))) = 0 Then
        Me.RecordSource = TABLE_ORD_ORDER
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "InitializeRecordSource", Err
End Sub

Private Sub ConfigureOrderLinesSubform()
    On Error GoTo SafeExit

    If Not HasControl(ORDER_LINES_SUBFORM_CONTROL) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ConfigureOrderLinesSubform", _
            "Subform control '" & ORDER_LINES_SUBFORM_CONTROL & "' was not found."
        Exit Sub
    End If

    Me.Controls(ORDER_LINES_SUBFORM_CONTROL).SourceObject = "Form." & ORDER_LINES_SUBFORM_NAME
    Me.Controls(ORDER_LINES_SUBFORM_CONTROL).LinkMasterFields = FIELD_ORDER_ID
    Me.Controls(ORDER_LINES_SUBFORM_CONTROL).LinkChildFields = FIELD_ORDER_ID

    modLoggingHandler.LogInfo MODULE_NAME & ".ConfigureOrderLinesSubform", _
        "Configured subform control='" & ORDER_LINES_SUBFORM_CONTROL & _
        "'; source_object='" & Me.Controls(ORDER_LINES_SUBFORM_CONTROL).SourceObject & _
        "'; link_master_fields='" & Me.Controls(ORDER_LINES_SUBFORM_CONTROL).LinkMasterFields & _
        "'; link_child_fields='" & Me.Controls(ORDER_LINES_SUBFORM_CONTROL).LinkChildFields & "'."

SafeExit:
End Sub

Private Sub RequeryOrderLinesSubform()
    On Error GoTo SafeExit

    If Not HasControl(ORDER_LINES_SUBFORM_CONTROL) Then
        Exit Sub
    End If

    Me.Controls(ORDER_LINES_SUBFORM_CONTROL).Form.Requery

SafeExit:
End Sub

Private Sub SaveCurrentRecord()
    On Error GoTo ErrorHandler

    Dim OrderId As Long

    OrderId = GetCurrentOrderId()

    If Me.Dirty Then
        modAuditHelper.ApplyAuditFields Me
        Me.Dirty = False
    End If

    If OrderId > 0 Then
        Call modOrderCalculationService.RecalculateOrder(OrderId)
    End If

    UpdateHeaderCaptionUi

    modLoggingHandler.LogInfo MODULE_NAME & ".SaveCurrentRecord", _
        "Sales order saved for order_id=" & CStr(OrderId) & "."

    ReturnToPreviousWorkspace "SaveCurrentRecord"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SaveCurrentRecord", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.ORDER_DETAIL_SAVE_ERROR", "Fehler beim Speichern der Bestellung."), vbExclamation, MODULE_NAME
End Sub

Private Sub CancelCurrentEdit( _
    Optional ByVal showConfirmation As Boolean = True, _
    Optional ByVal navigateBack As Boolean = True)
    On Error GoTo ErrorHandler

    If Not HasUnsavedChanges() Then
        If navigateBack Then
            ReturnToPreviousWorkspace "CancelCurrentEdit"
        End If
        Exit Sub
    End If

    If showConfirmation Then
        If MsgBox(modFwTranslationRuntime.ResolveText("MSG.ORDER_DETAIL_CANCEL_CONFIRM", "Aenderungen verwerfen?"), vbQuestion + vbYesNo + vbDefaultButton2, MODULE_NAME) <> vbYes Then
            Exit Sub
        End If
    End If

    Me.Undo
    UpdateHeaderCaptionUi

    modLoggingHandler.LogInfo MODULE_NAME & ".CancelCurrentEdit", _
        "Unsaved Sales Order changes were discarded."

    If navigateBack Then
        ReturnToPreviousWorkspace "CancelCurrentEdit"
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CancelCurrentEdit", Err
End Sub

Private Sub UpdateHeaderCaptionUi()
    On Error GoTo SafeExit

    Dim headerCaption As String
    Dim displayName As String

    headerCaption = modFwTranslationRuntime.ResolveText("FORM.FRMORDERDETAIL.FORM_TITLE", "Bestellung")

    If GetCurrentOrderId() > 0 Then
        displayName = Trim$(NzString(Me.Controls("order_no").Value))
        If LenB(displayName) = 0 Then
            displayName = CStr(GetCurrentOrderId())
        End If

        headerCaption = modFwTranslationRuntime.ResolveText("FORM.FRMORDERDETAIL.FORM_TITLE", "Bestellung") & ": " & displayName
    End If

    If HasControl("lblFormTitle") Then
        Me!lblFormTitle.Caption = headerCaption
    ElseIf HasControl("txtFormTitle") Then
        Me!txtFormTitle.Value = headerCaption
    End If

SafeExit:
End Sub

Private Sub UpdateLocalCommandButtons()
    On Error GoTo SafeExit

    If HasControl("cmdSave") Then
        Me!cmdSave.Visible = False
    End If

    If HasControl("cmdCancel") Then
        Me!cmdCancel.Visible = False
    End If

SafeExit:
End Sub

Private Function GetCurrentOrderId() As Long
    On Error GoTo SafeExit

    If Me.NewRecord Then
        Exit Function
    End If

    If HasControl("txtOrderId") Then
        GetCurrentOrderId = NzLong(Me!txtOrderId.Value, 0)
    ElseIf HasControl(FIELD_ORDER_ID) Then
        GetCurrentOrderId = NzLong(Me.Controls(FIELD_ORDER_ID).Value, 0)
    End If

SafeExit:
End Function

Private Sub RestoreCurrentOrderSelection(ByVal OrderId As Long)
    On Error GoTo ErrorHandler

    Dim recordsetClone As DAO.Recordset

    Set recordsetClone = Me.RecordsetClone
    If recordsetClone.BOF And recordsetClone.EOF Then
        modLoggingHandler.LogWarning MODULE_NAME & ".RestoreCurrentOrderSelection", _
            "RecordsetClone is empty while trying to restore order_id=" & CStr(OrderId) & "."
        GoTo CleanExit
    End If

    recordsetClone.FindFirst "[" & FIELD_ORDER_ID & "]=" & CStr(OrderId)
    If Not recordsetClone.NoMatch Then
        modLoggingHandler.LogInfo MODULE_NAME & ".RestoreCurrentOrderSelection", _
            "Order found for order_id=" & CStr(OrderId) & "."
        Me.Bookmark = recordsetClone.Bookmark
    Else
        modLoggingHandler.LogWarning MODULE_NAME & ".RestoreCurrentOrderSelection", _
            "Order not found for order_id=" & CStr(OrderId) & "."
        HandleInvalidOrderContext "MSG.ORDER_DETAIL_NOT_FOUND", "Die Bestellung konnte nicht gefunden werden."
    End If

CleanExit:
    On Error Resume Next
    If Not recordsetClone Is Nothing Then recordsetClone.Close
    Set recordsetClone = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreCurrentOrderSelection", Err
    Resume CleanExit
End Sub

Private Sub HandleInvalidOrderContext(ByVal messageKey As String, ByVal fallbackText As String)
    On Error Resume Next

    MsgBox modFwTranslationRuntime.ResolveText(messageKey, fallbackText), vbExclamation, MODULE_NAME
    ReturnToPreviousWorkspace "HandleInvalidOrderContext"
End Sub

Private Function ResolveHostingShellForm() As Access.Form
    On Error GoTo SafeExit

    Dim parentForm As Access.Form

    Set parentForm = Me.Parent
    If parentForm Is Nothing Then
        Exit Function
    End If

    If StrComp(parentForm.Name, "frmAppShell", vbTextCompare) = 0 Then
        Set ResolveHostingShellForm = parentForm
    End If

SafeExit:
End Function

Private Sub NavigateHome()
    On Error GoTo ErrorHandler

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()
    If shellForm Is Nothing Then
        Exit Sub
    End If

    If Not modAppWorkspaceService.OpenWorkspaceForm(shellForm, DASHBOARD_FORM_NAME) Then
        Err.Raise vbObjectError + 5410, MODULE_NAME & ".NavigateHome", _
            "Workspace navigation to dashboard failed."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NavigateHome", Err
End Sub

Private Sub NavigateBack()
    On Error GoTo ErrorHandler

    ReturnToPreviousWorkspace "NavigateBack"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NavigateBack", Err
End Sub

Private Sub ReturnToPreviousWorkspace(ByVal callerName As String)
    On Error GoTo ErrorHandler

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()
    If shellForm Is Nothing Then
        Exit Sub
    End If

    If modAppWorkspaceService.GoBack(shellForm) Then
        Exit Sub
    End If

    If Not modAppWorkspaceService.OpenWorkspaceForm(shellForm, DASHBOARD_FORM_NAME) Then
        modLoggingHandler.LogWarning MODULE_NAME & "." & callerName, _
            "Workspace history was not available and dashboard fallback could not be opened."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, callerName & ".ReturnToPreviousWorkspace", Err
End Sub

Private Function HasControl(ByVal ControlName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentControl As Control

    For Each currentControl In Me.Controls
        If StrComp(currentControl.Name, ControlName, vbTextCompare) = 0 Then
            HasControl = True
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

Private Function NzString(ByVal Value As Variant, Optional ByVal defaultValue As String = "") As String
    If IsNull(Value) Or IsEmpty(Value) Then
        NzString = defaultValue
    Else
        NzString = CStr(Value)
    End If
End Function

Private Function NzLong(ByVal Value As Variant, Optional ByVal defaultValue As Long = 0) As Long
    If IsNull(Value) Or IsEmpty(Value) Or LenB(Trim$(CStr(Value))) = 0 Then
        NzLong = defaultValue
    Else
        NzLong = CLng(Value)
    End If
End Function
