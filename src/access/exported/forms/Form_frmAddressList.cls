VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAddressList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmAddressList
' Purpose   : Searchable address overview and workflow entry point.
' Author    : Codex
' Version   : 0.1.1
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmAddressList"
Private Const LIST_CODE As String = "ADDRESS"
Private Const ACTION_NEW As String = "NEW"
Private Const ACTION_EDIT As String = "EDIT"
Private Const ACTION_ORDERS As String = "ORDERS"
Private Const ACTION_NEW_ORDER As String = "NEW_ORDER"
Private Const ACTION_SUBSCRIPTIONS As String = "SUBSCRIPTIONS"
Private Const ACTION_NEW_SUBSCRIPTION As String = "NEW_SUBSCRIPTION"
Private Const ACTION_EMAILS As String = "EMAILS"
Private Const ACTION_ACCOUNT As String = "ACCOUNT"
Private Const ACTION_REPORTS As String = "REPORTS"
Private Const FORM_ADDRESS_DETAIL As String = "frmAddressDetail"
Private Const FORM_ORDER_LIST As String = "frmOrderList"
Private Const FORM_ORDER_DETAIL As String = "frmOrderDetail"

Private Sub cmdClearSearch_Click()
    ClearSearchFilter
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    If HasControl("txtSearch") Then
        Me!txtSearch.Value = vbNullString
        Me.Filter = vbNullString
        Me.FilterOn = False
    End If

    LoadActionMenu
    modFwTranslationRuntime.ApplyTranslations Me
    UpdateActionState

    If HasControl("txtSearch") Then
        Me!txtSearch.SetFocus
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "Address list loaded."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox "Fehler beim Laden der Adressliste: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub txtSearch_Change()
    On Error GoTo ErrorHandler

    ApplySearchFilter
    Me!txtSearch.SetFocus
    Me!txtSearch.SelStart = Len(Nz(Me!txtSearch.text, vbNullString))

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtSearch_Change", Err
End Sub

Private Sub cmdExecuteAction_Click()
    On Error GoTo ErrorHandler

    ExecuteSelectedAction
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdExecuteAction_Click", Err
    MsgBox "Fehler beim Ausfuehren der Aktion: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cboListAction_AfterUpdate()
    On Error GoTo ErrorHandler

    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboListAction_AfterUpdate", Err
End Sub

Private Sub LoadActionMenu()
    On Error GoTo ErrorHandler

    If Not HasControl("cboListAction") Then
        Exit Sub
    End If

    Me!cboListAction.RowSourceType = "Table/Query"
    Me!cboListAction.RowSource = BuildActionRowSource()
    Me!cboListAction.BoundColumn = 1
    Me!cboListAction.ColumnCount = 4
    Me!cboListAction.ColumnWidths = "0cm;5cm;0cm;0cm"
    Me!cboListAction.LimitToList = True
    Me!cboListAction = Null
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadActionMenu", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Function BuildActionRowSource() As String
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "SELECT "
    sqlStatement = sqlStatement & "action_code, "
    sqlStatement = sqlStatement & "action_label, "
    sqlStatement = sqlStatement & "Nz(target_form, '') AS target_form_value, "
    sqlStatement = sqlStatement & "requires_selection "
    sqlStatement = sqlStatement & "FROM fw_list_action "
    sqlStatement = sqlStatement & "WHERE list_code = " & SqlText(LIST_CODE) & " "
    sqlStatement = sqlStatement & "AND is_active = True "
    sqlStatement = sqlStatement & "ORDER BY sort_order, action_label;"

    BuildActionRowSource = sqlStatement
End Function
Private Sub ApplySearchFilter()
    On Error GoTo ErrorHandler

    Dim filterExpression As String
    Dim SearchText As String

    If Not HasControl("txtSearch") Then
        Exit Sub
    End If

    SearchText = Trim$(Nz(Me!txtSearch.text, vbNullString))

    If LenB(SearchText) = 0 Then
        Me.FilterOn = False
        Me.Filter = vbNullString
    Else
        filterExpression = BuildSearchFilter(SearchText)
        Me.Filter = filterExpression
        Me.FilterOn = True
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplySearchFilter", _
        "Search filter updated."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplySearchFilter", Err
End Sub

Private Function BuildSearchFilter(ByVal SearchText As String) As String
    Dim searchExpression As String

    SearchText = EscapeLikeValue(UCase$(Trim$(SearchText)))

    searchExpression = ""
    searchExpression = searchExpression & "UCase(Nz([AddressSearchText],'')) "
    searchExpression = searchExpression & "Like '*" & SearchText & "*'"

    BuildSearchFilter = searchExpression
End Function

Public Sub ExecuteSelectedAction()
    On Error GoTo ErrorHandler

    Dim ActionCode As String
    Dim TargetForm As String
    Dim requiresSelection As Boolean
    Dim currentAddressId As Long

    ActionCode = GetSelectedActionCode()
    If LenB(ActionCode) = 0 Then
        MsgBox "Bitte zuerst eine Aktion auswaehlen.", vbExclamation, MODULE_NAME
        Exit Sub
    End If

    TargetForm = ResolveTargetForm(ActionCode)
    requiresSelection = SelectedActionRequiresSelection()
    currentAddressId = GetCurrentAddressId()

    If requiresSelection And currentAddressId <= 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExecuteSelectedAction", _
            "Action '" & ActionCode & "' requires a selected address."
        MsgBox "Bitte zuerst eine Adresse auswaehlen.", vbExclamation, MODULE_NAME
        Exit Sub
    End If

    Select Case ActionCode
        Case ACTION_NEW
            If Not FormExists(FORM_ADDRESS_DETAIL) Then
                ShowPlaceholderActionMessage ActionCode, FORM_ADDRESS_DETAIL, currentAddressId
                Exit Sub
            End If
            DoCmd.OpenForm FORM_ADDRESS_DETAIL, acNormal, , , acFormAdd

        Case ACTION_EDIT
            If Not FormExists(FORM_ADDRESS_DETAIL) Then
                ShowPlaceholderActionMessage ActionCode, FORM_ADDRESS_DETAIL, currentAddressId
                Exit Sub
            End If
            DoCmd.OpenForm FORM_ADDRESS_DETAIL, acNormal, , "[address_id]=" & CStr(currentAddressId)

        Case ACTION_ORDERS
            If LenB(TargetForm) = 0 Then TargetForm = FORM_ORDER_LIST
            If Not FormExists(TargetForm) Then
                ShowPlaceholderActionMessage ActionCode, TargetForm, currentAddressId
                Exit Sub
            End If
            DoCmd.OpenForm TargetForm, acNormal, , , , , CStr(currentAddressId)

        Case ACTION_NEW_ORDER
            If LenB(TargetForm) = 0 Then TargetForm = FORM_ORDER_DETAIL
            If Not FormExists(TargetForm) Then
                ShowPlaceholderActionMessage ActionCode, TargetForm, currentAddressId
                Exit Sub
            End If
            DoCmd.OpenForm TargetForm, acNormal, , , acFormAdd, , CStr(currentAddressId)

        Case ACTION_SUBSCRIPTIONS, ACTION_NEW_SUBSCRIPTION, ACTION_EMAILS, ACTION_ACCOUNT, ACTION_REPORTS
            If LenB(TargetForm) = 0 Or Not FormExists(TargetForm) Then
                ShowPlaceholderActionMessage ActionCode, TargetForm, currentAddressId
                Exit Sub
            End If

            If currentAddressId > 0 Then
                DoCmd.OpenForm TargetForm, acNormal, , , , , CStr(currentAddressId)
            Else
                DoCmd.OpenForm TargetForm
            End If

        Case Else
            If LenB(TargetForm) = 0 Or Not FormExists(TargetForm) Then
                ShowPlaceholderActionMessage ActionCode, TargetForm, currentAddressId
                Exit Sub
            End If

            If currentAddressId > 0 Then
                DoCmd.OpenForm TargetForm, acNormal, , , , , CStr(currentAddressId)
            Else
                DoCmd.OpenForm TargetForm
            End If
    End Select

    modLoggingHandler.LogInfo MODULE_NAME & ".ExecuteSelectedAction", _
        "Executed action '" & ActionCode & "' for address_id=" & CStr(currentAddressId) & "."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ExecuteSelectedAction", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Function GetSelectedActionCode() As String
    If HasControl("cboListAction") Then
        GetSelectedActionCode = Trim$(Nz(Me!cboListAction.Value, vbNullString))
    End If
End Function

Private Function ResolveTargetForm(ByVal ActionCode As String) As String
    If HasControl("cboListAction") Then
        If Me!cboListAction.ListIndex >= 0 Then
            ResolveTargetForm = Trim$(Nz(Me!cboListAction.Column(2), vbNullString))
        End If
    End If

    If LenB(ResolveTargetForm) > 0 Then
        Exit Function
    End If

    Select Case UCase$(Trim$(ActionCode))
        Case ACTION_NEW, ACTION_EDIT
            ResolveTargetForm = FORM_ADDRESS_DETAIL
        Case ACTION_ORDERS
            ResolveTargetForm = FORM_ORDER_LIST
        Case ACTION_NEW_ORDER
            ResolveTargetForm = FORM_ORDER_DETAIL
    End Select
End Function

Private Function SelectedActionRequiresSelection() As Boolean
    If HasControl("cboListAction") Then
        If Me!cboListAction.ListIndex >= 0 Then
            SelectedActionRequiresSelection = NzBoolean(Me!cboListAction.Column(3), False)
        End If
    End If
End Function

Private Function GetCurrentAddressId() As Long
    On Error GoTo SafeExit

    If Me.NewRecord Then
        Exit Function
    End If

    If HasControl("txtAddressId") Then
        GetCurrentAddressId = NzLong(Me!txtAddressId.Value, 0)
    ElseIf HasControl("AddressId") Then
        GetCurrentAddressId = NzLong(Me!AddressId.Value, 0)
    End If

SafeExit:
End Function

Private Sub UpdateActionState()
    On Error GoTo ErrorHandler

    If HasControl("cmdExecuteAction") Then
        Me!cmdExecuteAction.Enabled = (LenB(GetSelectedActionCode()) > 0)
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "UpdateActionState", Err
End Sub

Private Sub ShowPlaceholderActionMessage( _
    ByVal ActionCode As String, _
    ByVal TargetForm As String, _
    ByVal AddressId As Long)

    Dim messageText As String

    messageText = "Aktion '" & ActionCode & "' ist noch nicht verfuegbar."
    If LenB(Trim$(TargetForm)) > 0 Then
        messageText = messageText & vbCrLf & "Zielformular: " & TargetForm
    End If
    If AddressId > 0 Then
        messageText = messageText & vbCrLf & "AddressId: " & CStr(AddressId)
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".ShowPlaceholderActionMessage", _
        "Placeholder shown for action '" & ActionCode & "'."
    MsgBox messageText, vbInformation, MODULE_NAME
End Sub

Private Function FormExists(ByVal FormName As String) As Boolean
    On Error GoTo SafeExit

    Dim accessObject As Access.accessObject

    For Each accessObject In CurrentProject.AllForms
        If StrComp(accessObject.Name, FormName, vbTextCompare) = 0 Then
            FormExists = True
            Exit Function
        End If
    Next accessObject

SafeExit:
End Function

Private Function EscapeLikeValue(ByVal Value As String) As String
    Value = Replace(Value, "'", "''")
    Value = Replace(Value, "[", "[[]")
    Value = Replace(Value, "*", "[*]")
    Value = Replace(Value, "?", "[?]")
    Value = Replace(Value, "#", "[#]")
    EscapeLikeValue = Value
End Function

Private Function HasControl(ByVal ControlName As String) As Boolean
    On Error GoTo SafeExit

    Dim ctl As Control

    For Each ctl In Me.Controls
        If StrComp(ctl.Name, ControlName, vbTextCompare) = 0 Then
            HasControl = True
            Exit Function
        End If
    Next ctl

SafeExit:
End Function

Private Function SqlText(ByVal Value As String) As String
    SqlText = "'" & Replace(Nz(Value, vbNullString), "'", "''") & "'"
End Function

Private Function NzLong(ByVal Value As Variant, Optional ByVal DefaultValue As Long = 0) As Long
    If IsNull(Value) Or IsEmpty(Value) Or LenB(Trim$(CStr(Value))) = 0 Then
        NzLong = DefaultValue
    Else
        NzLong = CLng(Value)
    End If
End Function

Private Function NzBoolean(ByVal Value As Variant, Optional ByVal DefaultValue As Boolean = False) As Boolean
    If IsNull(Value) Or IsEmpty(Value) Then
        NzBoolean = DefaultValue
    Else
        NzBoolean = CBool(Value)
    End If
End Function


Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler

    If KeyCode = vbKeyEscape Then
        ClearSearchFilter
        KeyCode = 0
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtSearch_KeyDown", Err
End Sub
Private Sub ClearSearchFilter()
    On Error GoTo ErrorHandler

    Me!txtSearch = vbNullString
    Me.Filter = vbNullString
    Me.FilterOn = False

    Me!txtSearch.SetFocus

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ClearSearchFilter", Err
End Sub





