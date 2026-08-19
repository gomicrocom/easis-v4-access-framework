VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAddressCockpit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ExportKind: CODE_BEHIND_ONLY
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmAddressCockpit
' Purpose   : Address-centered cockpit and command bridge for customer work.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmAddressCockpit"
Private Const FORM_ORDER_DETAIL As String = "frmOrderDetail"
Private Const DASHBOARD_FORM_NAME As String = "frmAppDashboard"

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    InitializeRecordSource
    modFwTranslationRuntime.ApplyTranslations Me
    UpdateLocalActionState
    UpdateLockIndicator

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", "Address cockpit loaded."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    UpdateLocalActionState
    UpdateLockIndicator
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub cmdNewOrder_Click()
    On Error GoTo ErrorHandler

    CreateNewSalesOrder
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdNewOrder_Click", Err
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo ErrorHandler

    RefreshCockpitData
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdRefresh_Click", Err
End Sub

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

    cfg(modAppShell.WSLOT_R1 & ".visible") = True
    cfg(modAppShell.WSLOT_R1 & ".enabled") = True
    cfg(modAppShell.WSLOT_R1 & ".caption") = modFwTranslationRuntime.ResolveText("COMMON.REFRESH", "Aktualisieren")
    cfg(modAppShell.WSLOT_R1 & ".command") = modAppShell.WCMD_LIST_REFRESH

    cfg(modAppShell.WSLOT_R2 & ".visible") = False
    cfg(modAppShell.WSLOT_R2 & ".enabled") = False
    cfg(modAppShell.WSLOT_R2 & ".caption") = vbNullString
    cfg(modAppShell.WSLOT_R2 & ".command") = vbNullString

    cfg(modAppShell.WSLOT_R3 & ".visible") = False
    cfg(modAppShell.WSLOT_R3 & ".enabled") = False
    cfg(modAppShell.WSLOT_R3 & ".caption") = vbNullString
    cfg(modAppShell.WSLOT_R3 & ".command") = vbNullString

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
        Case modAppShell.WCMD_LIST_REFRESH
            RefreshCockpitData
    End Select
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ExecuteWorkspaceCommand", Err
End Sub

Public Function GetCurrentWorkspaceRecordId() As Long
    GetCurrentWorkspaceRecordId = GetCurrentAddressId()
End Function

Public Function GetWorkspaceState() As String
    Dim stateText As String

    stateText = AddWorkspaceStatePair(stateText, "address_id", CStr(GetCurrentAddressId()))
    GetWorkspaceState = stateText
End Function

Public Sub RestoreWorkspaceState(ByVal stateText As String)
    On Error GoTo ErrorHandler

    Dim addressId As Long

    addressId = modDaoHelper.NzLong(GetWorkspaceStateValue(stateText, "address_id"), 0)
    If addressId > 0 Then
        RestoreCurrentAddressSelection addressId
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreWorkspaceState", Err
End Sub

Public Sub ApplyWorkspaceOpenArgs(ByVal openArgs As String)
    On Error GoTo ErrorHandler

    Dim addressId As Long

    addressId = modDaoHelper.NzLong(openArgs, 0)
    If addressId > 0 Then
        RestoreCurrentAddressSelection addressId
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyWorkspaceOpenArgs", Err
End Sub

Private Sub InitializeRecordSource()
    On Error GoTo ErrorHandler

    If LenB(Trim$(modDaoHelper.NzString(Me.RecordSource))) = 0 Then
        Me.RecordSource = modAddressCockpitService.GetAddressCockpitRecordSource()
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "InitializeRecordSource", Err
End Sub

Private Sub RefreshCockpitData()
    On Error GoTo ErrorHandler

    Me.Requery
    UpdateLocalActionState
    UpdateLockIndicator
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RefreshCockpitData", Err
End Sub

Private Sub CreateNewSalesOrder()
    On Error GoTo ErrorHandler

    Dim addressId As Long
    Dim orderId As Long

    addressId = GetCurrentAddressId()
    If addressId <= 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ADDRESS_COCKPIT_SELECT_ADDRESS_FIRST", "Bitte zuerst eine Adresse auswaehlen."), vbExclamation, MODULE_NAME
        Exit Sub
    End If

    orderId = modAddressCockpitService.CreateDraftSalesOrderForAddress(addressId)
    If orderId <= 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ADDRESS_COCKPIT_ORDER_CREATE_FAILED", "Die Bestellung konnte nicht erstellt werden."), vbExclamation, MODULE_NAME
        Exit Sub
    End If

    If FormExists(FORM_ORDER_DETAIL) Then
        If Not modAppWorkspaceService.OpenWorkspaceForm(Me, FORM_ORDER_DETAIL, "[order_id]=" & CStr(orderId)) Then
            Err.Raise vbObjectError + 5320, MODULE_NAME & ".CreateNewSalesOrder", "Workspace navigation to frmOrderDetail failed."
        End If
    Else
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ADDRESS_COCKPIT_ORDER_DETAIL_MISSING", "Die Bestellung wurde erstellt, aber frmOrderDetail ist nicht verfuegbar."), vbInformation, MODULE_NAME
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CreateNewSalesOrder", Err
End Sub

Private Sub UpdateLocalActionState()
    On Error GoTo ErrorHandler

    SetControlEnabledIfPresent "cmdNewOrder", (GetCurrentAddressId() > 0)

    SetPlaceholderControlState "cmdManageOrders"
    SetPlaceholderControlState "cmdNewSubscription"
    SetPlaceholderControlState "cmdManageSubscriptions"
    SetPlaceholderControlState "cmdAccountStatement"
    SetPlaceholderControlState "cmdCapturePayment"
    SetPlaceholderControlState "cmdDunning"
    SetPlaceholderControlState "cmdNotes"
    SetPlaceholderControlState "cmdEmails"
    SetPlaceholderControlState "cmdReports"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "UpdateLocalActionState", Err
End Sub

Private Sub UpdateLockIndicator()
    On Error GoTo ErrorHandler

    If HasControl("lblLockHint") Then
        Me!lblLockHint.Visible = IsCurrentAddressLocked()
    End If

    If HasControl("txtLockReason") Then
        Me!txtLockReason.Visible = IsCurrentAddressLocked()
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "UpdateLockIndicator", Err
End Sub

Private Function GetCurrentAddressId() As Long
    On Error GoTo SafeExit

    If Me.NewRecord Then
        Exit Function
    End If

    If HasControl("address_id") Then
        GetCurrentAddressId = modDaoHelper.NzLong(Me!address_id.Value, 0)
    ElseIf HasControl("txtAddressId") Then
        GetCurrentAddressId = modDaoHelper.NzLong(Me!txtAddressId.Value, 0)
    End If

SafeExit:
End Function

Private Function IsCurrentAddressLocked() As Boolean
    If HasControl("is_locked") Then
        IsCurrentAddressLocked = modDaoHelper.NzBoolean(Me!is_locked.Value, False)
    End If
End Function

Private Sub RestoreCurrentAddressSelection(ByVal addressId As Long)
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset

    Set rs = Me.RecordsetClone
    If rs.BOF And rs.EOF Then
        GoTo CleanExit
    End If

    rs.FindFirst "[address_id]=" & CStr(addressId)
    If Not rs.NoMatch Then
        Me.Bookmark = rs.Bookmark
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreCurrentAddressSelection", Err
    Resume CleanExit
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
        Err.Raise vbObjectError + 5321, MODULE_NAME & ".NavigateHome", "Workspace navigation to dashboard failed."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NavigateHome", Err
End Sub

Private Sub NavigateBack()
    On Error GoTo ErrorHandler

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()
    If shellForm Is Nothing Then
        Exit Sub
    End If

    If Not modAppWorkspaceService.GoBack(shellForm) Then
        Err.Raise vbObjectError + 5322, MODULE_NAME & ".NavigateBack", "Workspace back navigation failed."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NavigateBack", Err
End Sub

Private Sub SetPlaceholderControlState(ByVal controlName As String)
    On Error Resume Next

    If HasControl(controlName) Then
        Me.Controls(controlName).Visible = True
        Me.Controls(controlName).Enabled = False
    End If
End Sub

Private Sub SetControlEnabledIfPresent(ByVal controlName As String, ByVal enabledState As Boolean)
    On Error Resume Next

    If HasControl(controlName) Then
        Me.Controls(controlName).Enabled = enabledState
    End If
End Sub

Private Function HasControl(ByVal controlName As String) As Boolean
    On Error GoTo SafeExit

    Dim ctl As Control
    Set ctl = Me.Controls(controlName)
    HasControl = Not ctl Is Nothing

SafeExit:
    Set ctl = Nothing
End Function

Private Function FormExists(ByVal FormName As String) As Boolean
    On Error GoTo SafeExit

    Dim accessObject As AccessObject

    For Each accessObject In CurrentProject.AllForms
        If StrComp(accessObject.Name, FormName, vbTextCompare) = 0 Then
            FormExists = True
            Exit Function
        End If
    Next accessObject

SafeExit:
End Function

Private Function EscapeWorkspaceStateValue(ByVal Value As String) As String
    Value = Replace(Nz(Value, vbNullString), "%", "%25")
    Value = Replace(Value, ";", "%3B")
    Value = Replace(Value, "=", "%3D")
    EscapeWorkspaceStateValue = Value
End Function

Private Function UnescapeWorkspaceStateValue(ByVal Value As String) As String
    Value = Replace(Nz(Value, vbNullString), "%3D", "=")
    Value = Replace(Value, "%3B", ";")
    Value = Replace(Value, "%25", "%")
    UnescapeWorkspaceStateValue = Value
End Function

Private Function AddWorkspaceStatePair(ByVal stateText As String, ByVal keyName As String, ByVal keyValue As String) As String
    If LenB(Trim$(keyName)) = 0 Then
        AddWorkspaceStatePair = stateText
        Exit Function
    End If

    If LenB(stateText) > 0 Then
        stateText = stateText & ";"
    End If

    AddWorkspaceStatePair = stateText & keyName & "=" & EscapeWorkspaceStateValue(keyValue)
End Function

Private Function GetWorkspaceStateValue(ByVal stateText As String, ByVal keyName As String) As String
    Dim parts() As String
    Dim part As Variant
    Dim position As Long

    If LenB(stateText) = 0 Or LenB(keyName) = 0 Then
        Exit Function
    End If

    parts = Split(stateText, ";")
    For Each part In parts
        position = InStr(1, CStr(part), "=", vbBinaryCompare)
        If position > 0 Then
            If StrComp(Left$(CStr(part), position - 1), keyName, vbTextCompare) = 0 Then
                GetWorkspaceStateValue = UnescapeWorkspaceStateValue(Mid$(CStr(part), position + 1))
                Exit Function
            End If
        End If
    Next part
End Function
