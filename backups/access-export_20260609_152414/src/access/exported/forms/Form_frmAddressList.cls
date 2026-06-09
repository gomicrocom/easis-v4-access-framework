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
' Purpose   : Searchable address overview and workspace-aware workflow entry point.
' Author    : Codex
' Version   : 0.3.1
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

Private m_currentSearchText As String

Private Sub cmdEdit_Click()
    On Error GoTo ErrorHandler

    ExecuteActionByCode ACTION_EDIT
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdEdit_Click", Err
End Sub

Private Sub cmdExecuteAction_Click()
    On Error GoTo ErrorHandler

    ExecuteSelectedAction
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdExecuteAction_Click", Err
    MsgBox "Fehler beim Ausfuehren der Aktion: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdNew_Click()
    On Error GoTo ErrorHandler

    ExecuteActionByCode ACTION_NEW
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdNew_Click", Err
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo ErrorHandler

    RefreshListData
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdRefresh_Click", Err
End Sub

Private Sub cboListAction_AfterUpdate()
    On Error GoTo ErrorHandler

    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboListAction_AfterUpdate", Err
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    m_currentSearchText = vbNullString
    Me.Filter = vbNullString
    Me.FilterOn = False

    LoadActionMenu
    modFwTranslationRuntime.ApplyTranslations Me
    UpdateActionState

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "Address list loaded."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox "Fehler beim Laden der Adressliste: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub ApplySearchFilter(Optional ByVal searchText As String = vbNullString)
    On Error GoTo ErrorHandler

    Dim filterExpression As String

    searchText = Trim$(Nz(searchText, vbNullString))
    m_currentSearchText = searchText

    Me.FilterOn = False

    If LenB(searchText) = 0 Then
        Me.Filter = vbNullString
    Else
        filterExpression = BuildSearchFilter(searchText)
        Me.Filter = filterExpression
        Me.FilterOn = True
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplySearchFilter", _
        "Search filter updated."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplySearchFilter", Err
End Sub

Private Sub ClearSearchFilter()
    On Error GoTo ErrorHandler

    m_currentSearchText = vbNullString
    Me.Filter = vbNullString
    Me.FilterOn = False
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ClearSearchFilter", Err
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

Private Function BuildSearchFilter(ByVal searchText As String) As String
    Dim searchExpression As String

    searchText = EscapeLikeValue(UCase$(Trim$(searchText)))

    searchExpression = ""
    searchExpression = searchExpression & "UCase(Nz([AddressSearchText],'')) "
    searchExpression = searchExpression & "Like '*" & searchText & "*'"

    BuildSearchFilter = searchExpression
End Function

Public Function GetCurrentWorkspaceRecordId() As Long
    GetCurrentWorkspaceRecordId = GetCurrentAddressId()
End Function

Public Function SupportsListCommandBar() As Boolean
    SupportsListCommandBar = True
End Function

Public Function SupportsListNew() As Boolean
    SupportsListNew = True
End Function

Public Function SupportsListEdit() As Boolean
    SupportsListEdit = True
End Function

Public Sub ListSearch(ByVal searchText As String)
    On Error GoTo ErrorHandler

    m_currentSearchText = Trim$(Nz(searchText, vbNullString))
    ApplySearchFilter m_currentSearchText
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ListSearch", Err
End Sub

Public Sub ListClearSearch()
    On Error GoTo ErrorHandler

    ClearSearchFilter
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ListClearSearch", Err
End Sub

Public Sub ListNew()
    On Error GoTo ErrorHandler

    ExecuteActionByCode ACTION_NEW
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ListNew", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Public Sub ListEdit()
    On Error GoTo ErrorHandler

    ExecuteActionByCode ACTION_EDIT
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ListEdit", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Public Sub ListRefresh()
    On Error GoTo ErrorHandler

    RefreshListData
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ListRefresh", Err
End Sub

Public Function GetWorkspaceState() As String
    Dim searchText As String
    Dim orderByText As String
    Dim stateText As String

    searchText = GetSearchTextValue()

    If Me.OrderByOn Then
        orderByText = Nz(Me.OrderBy, vbNullString)
    End If

    stateText = AddWorkspaceStatePair(stateText, "search", searchText)
    stateText = AddWorkspaceStatePair(stateText, "address_id", CStr(GetCurrentAddressId()))
    stateText = AddWorkspaceStatePair(stateText, "order_by", orderByText)

    GetWorkspaceState = stateText
End Function

Public Sub RestoreWorkspaceState(ByVal stateText As String)
    On Error GoTo ErrorHandler

    Dim searchText As String
    Dim addressId As Long
    Dim orderByText As String

    searchText = GetWorkspaceStateValue(stateText, "search")
    orderByText = GetWorkspaceStateValue(stateText, "order_by")
    addressId = NzLong(GetWorkspaceStateValue(stateText, "address_id"), 0)

    m_currentSearchText = searchText
    ApplySearchFilter m_currentSearchText

    Me.OrderBy = orderByText
    Me.OrderByOn = (LenB(orderByText) > 0)

    If addressId > 0 Then
        RestoreCurrentAddressSelection addressId
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreWorkspaceState", Err
End Sub

Public Sub ExecuteSelectedAction()
    On Error GoTo ErrorHandler

    Dim actionCode As String

    actionCode = GetSelectedActionCode()
    If LenB(actionCode) = 0 Then
        MsgBox "Bitte zuerst eine Aktion auswaehlen.", vbExclamation, MODULE_NAME
        Exit Sub
    End If

    ExecuteActionByCode actionCode
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ExecuteSelectedAction", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub ExecuteActionByCode(ByVal actionCode As String)
    On Error GoTo ErrorHandler

    Dim targetForm As String
    Dim requiresSelection As Boolean
    Dim currentAddressId As Long

    targetForm = ResolveTargetForm(actionCode)
    requiresSelection = ResolveActionRequiresSelection(actionCode)
    currentAddressId = GetCurrentAddressId()

    If requiresSelection And currentAddressId <= 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExecuteActionByCode", _
            "Action '" & actionCode & "' requires a selected address."
        MsgBox "Bitte zuerst eine Adresse auswaehlen.", vbExclamation, MODULE_NAME
        Exit Sub
    End If

    Select Case actionCode
        Case ACTION_NEW
            If Not FormExists(FORM_ADDRESS_DETAIL) Then
                ShowPlaceholderActionMessage actionCode, FORM_ADDRESS_DETAIL, currentAddressId
                Exit Sub
            End If
            OpenTargetForm FORM_ADDRESS_DETAIL, vbNullString, acFormAdd

        Case ACTION_EDIT
            If Not FormExists(FORM_ADDRESS_DETAIL) Then
                ShowPlaceholderActionMessage actionCode, FORM_ADDRESS_DETAIL, currentAddressId
                Exit Sub
            End If
            OpenTargetForm FORM_ADDRESS_DETAIL, "[address_id]=" & CStr(currentAddressId)

        Case ACTION_ORDERS
            If LenB(targetForm) = 0 Then targetForm = FORM_ORDER_LIST
            If Not FormExists(targetForm) Then
                ShowPlaceholderActionMessage actionCode, targetForm, currentAddressId
                Exit Sub
            End If
            OpenTargetForm targetForm, vbNullString, acFormPropertySettings, CStr(currentAddressId)

        Case ACTION_NEW_ORDER
            If LenB(targetForm) = 0 Then targetForm = FORM_ORDER_DETAIL
            If Not FormExists(targetForm) Then
                ShowPlaceholderActionMessage actionCode, targetForm, currentAddressId
                Exit Sub
            End If
            OpenTargetForm targetForm, vbNullString, acFormAdd, CStr(currentAddressId)

        Case ACTION_SUBSCRIPTIONS, ACTION_NEW_SUBSCRIPTION, ACTION_EMAILS, ACTION_ACCOUNT, ACTION_REPORTS
            If LenB(targetForm) = 0 Or Not FormExists(targetForm) Then
                ShowPlaceholderActionMessage actionCode, targetForm, currentAddressId
                Exit Sub
            End If

            If currentAddressId > 0 Then
                OpenTargetForm targetForm, vbNullString, acFormPropertySettings, CStr(currentAddressId)
            Else
                OpenTargetForm targetForm
            End If

        Case Else
            If LenB(targetForm) = 0 Or Not FormExists(targetForm) Then
                ShowPlaceholderActionMessage actionCode, targetForm, currentAddressId
                Exit Sub
            End If

            If currentAddressId > 0 Then
                OpenTargetForm targetForm, vbNullString, acFormPropertySettings, CStr(currentAddressId)
            Else
                OpenTargetForm targetForm
            End If
    End Select

    modLoggingHandler.LogInfo MODULE_NAME & ".ExecuteActionByCode", _
        "Executed action '" & actionCode & "' for address_id=" & CStr(currentAddressId) & "."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ExecuteActionByCode", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Function GetCurrentAddressId() As Long
    On Error GoTo SafeExit

    If Me.NewRecord Then
        Exit Function
    End If

    If HasControl("txtAddressId") Then
        GetCurrentAddressId = NzLong(Me!txtAddressId.Value, 0)
    ElseIf HasControl("AddressId") Then
        GetCurrentAddressId = NzLong(Me!addressId.Value, 0)
    End If

SafeExit:
End Function

Private Function GetSelectedActionCode() As String
    If HasControl("cboListAction") Then
        GetSelectedActionCode = Trim$(Nz(Me!cboListAction.Value, vbNullString))
    End If
End Function

Private Function ResolveActionRequiresSelection(ByVal actionCode As String) As Boolean
    If HasControl("cboListAction") Then
        If Me!cboListAction.ListIndex >= 0 Then
            ResolveActionRequiresSelection = NzBoolean(Me!cboListAction.Column(3), False)
            Exit Function
        End If
    End If

    Select Case UCase$(Trim$(actionCode))
        Case ACTION_EDIT, ACTION_ORDERS, ACTION_NEW_ORDER, ACTION_SUBSCRIPTIONS, _
             ACTION_NEW_SUBSCRIPTION, ACTION_EMAILS, ACTION_ACCOUNT, ACTION_REPORTS
            ResolveActionRequiresSelection = True
    End Select
End Function

Private Function ResolveTargetForm(ByVal actionCode As String) As String
    If HasControl("cboListAction") Then
        If Me!cboListAction.ListIndex >= 0 Then
            ResolveTargetForm = Trim$(Nz(Me!cboListAction.Column(2), vbNullString))
        End If
    End If

    If LenB(ResolveTargetForm) > 0 Then
        Exit Function
    End If

    Select Case UCase$(Trim$(actionCode))
        Case ACTION_NEW, ACTION_EDIT
            ResolveTargetForm = FORM_ADDRESS_DETAIL
        Case ACTION_ORDERS
            ResolveTargetForm = FORM_ORDER_LIST
        Case ACTION_NEW_ORDER
            ResolveTargetForm = FORM_ORDER_DETAIL
    End Select
End Function

Private Sub LoadActionMenu()
    On Error GoTo ErrorHandler

    If Not HasControl("cboListAction") Then
        Exit Sub
    End If

    Me!cboListAction.RowSourceType = "Table/Query"
    Me!cboListAction.rowSource = BuildActionRowSource()
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

Private Sub OpenTargetForm( _
    ByVal targetForm As String, _
    Optional ByVal whereCondition As String = "", _
    Optional ByVal dataMode As AcFormOpenDataMode = acFormPropertySettings, _
    Optional ByVal openArgs As String = "")
    On Error GoTo ErrorHandler

    If Not modAppWorkspaceService.OpenWorkspaceForm(Me, targetForm, whereCondition, dataMode, openArgs) Then
        Err.Raise vbObjectError + 2000, MODULE_NAME & ".OpenTargetForm", _
            "Workspace navigation failed for form '" & targetForm & "'."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "OpenTargetForm", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub RefreshListData()
    On Error GoTo ErrorHandler

    Me.Requery
    UpdateActionState

    modLoggingHandler.LogInfo MODULE_NAME & ".RefreshListData", _
        "Address list requeried."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RefreshListData", Err
End Sub

Private Sub ShowPlaceholderActionMessage( _
    ByVal actionCode As String, _
    ByVal targetForm As String, _
    ByVal addressId As Long)

    Dim messageText As String

    messageText = "Aktion '" & actionCode & "' ist noch nicht verfuegbar."
    If LenB(Trim$(targetForm)) > 0 Then
        messageText = messageText & vbCrLf & "Zielformular: " & targetForm
    End If
    If addressId > 0 Then
        messageText = messageText & vbCrLf & "AddressId: " & CStr(addressId)
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".ShowPlaceholderActionMessage", _
        "Placeholder shown for action '" & actionCode & "'."
    MsgBox messageText, vbInformation, MODULE_NAME
End Sub

Private Sub UpdateActionState()
    On Error GoTo ErrorHandler

    If HasControl("cmdExecuteAction") Then
        Me!cmdExecuteAction.Enabled = (LenB(GetSelectedActionCode()) > 0)
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "UpdateActionState", Err
End Sub

Private Function EscapeLikeValue(ByVal Value As String) As String
    Value = Replace(Value, "'", "''")
    Value = Replace(Value, "[", "[[]")
    Value = Replace(Value, "*", "[*]")
    Value = Replace(Value, "?", "[?]")
    Value = Replace(Value, "#", "[#]")
    EscapeLikeValue = Value
End Function

Private Function EscapeWorkspaceStateValue(ByVal Value As String) As String
    Value = Replace(Nz(Value, vbNullString), "%", "%25")
    Value = Replace(Value, ";", "%3B")
    Value = Replace(Value, "=", "%3D")
    EscapeWorkspaceStateValue = Value
End Function

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

Private Function GetSearchTextValue() As String
    GetSearchTextValue = m_currentSearchText
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

Private Function NzBoolean(ByVal Value As Variant, Optional ByVal DefaultValue As Boolean = False) As Boolean
    If IsNull(Value) Or IsEmpty(Value) Then
        NzBoolean = DefaultValue
    Else
        NzBoolean = CBool(Value)
    End If
End Function

Private Function NzLong(ByVal Value As Variant, Optional ByVal DefaultValue As Long = 0) As Long
    If IsNull(Value) Or IsEmpty(Value) Or LenB(Trim$(CStr(Value))) = 0 Then
        NzLong = DefaultValue
    Else
        NzLong = CLng(Value)
    End If
End Function

Private Function AddWorkspaceStatePair( _
    ByVal stateText As String, _
    ByVal keyName As String, _
    ByVal valueText As String) As String

    If LenB(stateText) > 0 Then
        stateText = stateText & ";"
    End If

    AddWorkspaceStatePair = stateText & keyName & "=" & EscapeWorkspaceStateValue(valueText)
End Function

Private Sub RestoreCurrentAddressSelection(ByVal addressId As Long)
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset

    Set rs = Me.recordsetClone
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

Private Function UnescapeWorkspaceStateValue(ByVal Value As String) As String
    Value = Replace(Value, "%3D", "=")
    Value = Replace(Value, "%3B", ";")
    Value = Replace(Value, "%25", "%")
    UnescapeWorkspaceStateValue = Value
End Function

