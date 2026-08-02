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
Private Const TABLE_ORD_ORDER_LINE As String = "ord_order_line"
Private Const OPEN_ARGS_NEW_ORDER_PREFIX As String = "TMP_ORDER"
Private Const FIELD_ORDER_ID As String = "order_id"
Private Const FIELD_ORDER_LINE_ID As String = "order_line_id"
Private Const FIELD_TMP_ORDER_ID As String = "tmp_order_id"
Private Const FIELD_TMP_ORDER_LINE_ID As String = "tmp_order_line_id"
Private Const FIELD_ORDER_NO As String = "order_no"
Private Const FIELD_ORDER_TYPE_CODE As String = "order_type_code"
Private Const FIELD_ORDER_STATUS_CODE As String = "order_status_code"
Private Const FIELD_ORDER_DATE As String = "order_date"
Private Const FIELD_CUSTOMER_ADDRESS_ID As String = "customer_address_id"
Private Const FIELD_INVOICE_ADDRESS_ID As String = "invoice_address_id"
Private Const FIELD_DELIVERY_ADDRESS_ID As String = "delivery_address_id"
Private Const FIELD_CUSTOMER_NAME As String = "customer_name"
Private Const FIELD_LANGUAGE_CODE As String = "language_code"
Private Const FIELD_CURRENCY_CODE As String = "currency_code"
Private Const FIELD_PAYMENT_TERM_CODE As String = "payment_term_code"
Private Const FIELD_VAT_MODE As String = "vat_mode"
Private Const FIELD_VAT_CODE As String = "vat_code"
Private Const FIELD_VAT_RATE As String = "vat_rate"
Private Const TABLE_TMP_ORDER As String = "tmp_order"
Private Const TABLE_TMP_ORDER_LINE As String = "tmp_order_line"
Private Const DASHBOARD_FORM_NAME As String = "frmAppDashboard"
Private Const ORDER_LINES_SUBFORM_CONTROL As String = "subOrderLines"
Private Const ORDER_LINES_SUBFORM_NAME As String = "sfrmOrderLines"

Private m_workspaceOpenArgsApplied As Boolean
Private m_isTemporaryOrder As Boolean
Private m_temporaryOrderId As Long
Private m_temporaryOrderDisposed As Boolean
Private m_lastAppliedWorkspaceOpenArgs As String

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo ErrorHandler

    PrimePaymentTermComboEarly
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Open", Err
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    Dim deferWorkspaceContext As Boolean
    Dim effectiveOpenArgs As String
    Dim hasEffectiveOpenArgs As Boolean

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "OpenArgs raw='" & Replace(modDaoHelper.NzString(Me.OpenArgs), "'", "''") & "'."

    effectiveOpenArgs = modDaoHelper.NzString(Me.OpenArgs)
    If LenB(Trim$(effectiveOpenArgs)) = 0 Then
        effectiveOpenArgs = modAppWorkspaceService.ConsumePendingWorkspaceOpenArgs(Me.Name)
        If LenB(Trim$(effectiveOpenArgs)) > 0 Then
            modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
                "Consumed pending workspace OpenArgs='" & Replace(effectiveOpenArgs, "'", "''") & "' during Form_Load."
        End If
    End If
    hasEffectiveOpenArgs = (LenB(Trim$(effectiveOpenArgs)) > 0)

    InitializeRecordSource
    If hasEffectiveOpenArgs Then
        modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
            "Applying effective OpenArgs during Form_Load."
        ApplyWorkspaceOpenArgs effectiveOpenArgs
    Else
        deferWorkspaceContext = ShouldDeferWorkspaceContextBinding(effectiveOpenArgs)
        If deferWorkspaceContext Then
            PrepareDeferredWorkspaceContext
        End If
        ConfigureOrderLinesSubform
        ConfigureHeaderControls
        UpdateHeaderCaptionUi

        modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
            "No OpenArgs available during Form_Load. Waiting for workspace ApplyWorkspaceOpenArgs if hosted."
    End If

    modFwTranslationRuntime.ApplyTranslations Me
    UpdateLocalCommandButtons
    UpdateHeaderCaptionUi

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
End Sub

Private Function ShouldDeferWorkspaceContextBinding(ByVal effectiveOpenArgs As String) As Boolean
    On Error GoTo SafeExit

    If ResolveHostingShellForm() Is Nothing Then
        Exit Function
    End If

    If LenB(Trim$(effectiveOpenArgs)) > 0 Then
        Exit Function
    End If

    ShouldDeferWorkspaceContextBinding = True

SafeExit:
End Function

Private Sub PrepareDeferredWorkspaceContext()
    On Error GoTo ErrorHandler

    m_workspaceOpenArgsApplied = False
    m_isTemporaryOrder = False
    m_temporaryOrderId = 0
    m_temporaryOrderDisposed = False
    Me.Filter = "[" & FIELD_ORDER_ID & "]=0"
    Me.FilterOn = True

    modLoggingHandler.LogInfo MODULE_NAME & ".PrepareDeferredWorkspaceContext", _
        "Hosted workspace form was initialized without direct OpenArgs. Applied neutral deferred binding until workspace OpenArgs arrive."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "PrepareDeferredWorkspaceContext", Err
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    RefreshPaymentTermDisplay
    EnsureDefaultPaymentTerm
    RefreshVatModeDisplay
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

Private Sub cmdApply_Click()
    On Error GoTo ErrorHandler

    DetailApply
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdApply_Click", Err
End Sub

Private Sub cboDeliveryAddressId_AfterUpdate()
    On Error GoTo ErrorHandler

    ApplyDeliveryAddressVatDefaults
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboDeliveryAddressId_AfterUpdate", Err
End Sub

Private Sub cboLanguageCode_AfterUpdate()
    On Error GoTo ErrorHandler

    RefreshPaymentTermDisplay
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboLanguageCode_AfterUpdate", Err
End Sub

Private Sub cboVatMode_AfterUpdate()
    On Error GoTo ErrorHandler

    RefreshVatModeDisplay

    If m_isTemporaryOrder And m_temporaryOrderId > 0 Then
        modOrderCalculationService.RecalculateTemporaryOrder m_temporaryOrderId
        RefreshCalculatedTotals True
    ElseIf GetCurrentOrderId() > 0 Then
        modOrderCalculationService.RecalculateOrder GetCurrentOrderId()
        RefreshCalculatedTotals True
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboVatMode_AfterUpdate", Err
End Sub

Private Sub txtVatRate_AfterUpdate()
    On Error GoTo ErrorHandler

    RefreshVatModeDisplay
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtVatRate_AfterUpdate", Err
End Sub

Public Function CanLeaveWorkspace() As Boolean
    On Error GoTo ErrorHandler

    If m_isTemporaryOrder And m_temporaryOrderId > 0 And Not m_temporaryOrderDisposed Then
        If MsgBox( _
            modFwTranslationRuntime.ResolveText("MSG.ORDER_DETAIL_CANCEL_CONFIRM", "Aenderungen verwerfen?"), _
            vbQuestion + vbYesNo + vbDefaultButton2, _
            MODULE_NAME) = vbYes Then
            Call modOrderRepository.DeleteTemporaryOrder(m_temporaryOrderId)
            m_temporaryOrderDisposed = True
            CanLeaveWorkspace = True
        End If
        Exit Function
    End If

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

    cfg(modAppShell.WSLOT_R3 & ".visible") = True
    cfg(modAppShell.WSLOT_R3 & ".enabled") = True
    If m_isTemporaryOrder Then
        cfg(modAppShell.WSLOT_R3 & ".caption") = modFwTranslationRuntime.ResolveText("COMMON.REFRESH", "Aktualisieren")
    Else
        cfg(modAppShell.WSLOT_R3 & ".caption") = "Uebernehmen"
    End If
    cfg(modAppShell.WSLOT_R3 & ".command") = modAppShell.WCMD_DETAIL_APPLY

    cfg(modAppShell.WSLOT_R2 & ".visible") = True
    cfg(modAppShell.WSLOT_R2 & ".enabled") = True
    cfg(modAppShell.WSLOT_R2 & ".caption") = modFwTranslationRuntime.ResolveText("COMMON.CANCEL", "Abbrechen")
    cfg(modAppShell.WSLOT_R2 & ".command") = modAppShell.WCMD_DETAIL_CANCEL

    cfg(modAppShell.WSLOT_R1 & ".visible") = True
    cfg(modAppShell.WSLOT_R1 & ".enabled") = True
    If m_isTemporaryOrder Then
        cfg(modAppShell.WSLOT_R1 & ".caption") = "Bestellung speichern"
    Else
        cfg(modAppShell.WSLOT_R1 & ".caption") = modFwTranslationRuntime.ResolveText("COMMON.SAVE", "Speichern")
    End If
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
        Case modAppShell.WCMD_DETAIL_APPLY, modAppShell.WCMD_DETAIL_SAVE, modAppShell.WCMD_DETAIL_CANCEL
            CanExecuteWorkspaceCommand = True
    End Select
End Function

Public Sub ExecuteWorkspaceCommand(ByVal commandKey As String, Optional ByVal commandValue As String = "")
    Select Case UCase$(Trim$(commandKey))
        Case modAppShell.WCMD_NAV_HOME
            NavigateHome
        Case modAppShell.WCMD_NAV_BACK
            NavigateBack
        Case modAppShell.WCMD_DETAIL_APPLY
            DetailApply
        Case modAppShell.WCMD_DETAIL_SAVE
            DetailSave
        Case modAppShell.WCMD_DETAIL_CANCEL
            DetailCancel
    End Select
End Sub

Public Function GetCurrentWorkspaceRecordId() As Long
    If m_isTemporaryOrder Then
        GetCurrentWorkspaceRecordId = m_temporaryOrderId
    Else
        GetCurrentWorkspaceRecordId = GetCurrentOrderId()
    End If
End Function

Public Function GetCurrentOrderDetailId() As Long
    GetCurrentOrderDetailId = GetCurrentOrderHeaderId()
End Function

Public Function GetWorkspaceState() As String
    Dim stateText As String

    If m_isTemporaryOrder Then
        stateText = AddWorkspaceStatePair(stateText, "tmp_order_id", CStr(m_temporaryOrderId))
    Else
        stateText = AddWorkspaceStatePair(stateText, "order_id", CStr(GetCurrentOrderId()))
    End If
    GetWorkspaceState = stateText
End Function

Public Sub RestoreWorkspaceState(ByVal stateText As String)
    On Error GoTo ErrorHandler

    Dim OrderId As Long
    Dim tmpOrderId As Long

    tmpOrderId = NzLong(GetWorkspaceStateValue(stateText, "tmp_order_id"), 0)
    If tmpOrderId > 0 Then
        ActivateTemporaryOrderContext tmpOrderId
        Exit Sub
    End If

    OrderId = NzLong(GetWorkspaceStateValue(stateText, "order_id"), 0)
    If OrderId > 0 Then
        RestoreCurrentOrderSelection OrderId
    End If

    ConfigureOrderLinesSubform
    ConfigureHeaderControls
    UpdateHeaderCaptionUi
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreWorkspaceState", Err
End Sub

Public Sub ApplyWorkspaceOpenArgs(ByVal openArgs As String)
    On Error GoTo ErrorHandler

    Dim OrderId As Long
    Dim normalizedOpenArgs As String

    normalizedOpenArgs = Trim$(modDaoHelper.NzString(openArgs))

    If m_workspaceOpenArgsApplied Then
        If StrComp(normalizedOpenArgs, Trim$(m_lastAppliedWorkspaceOpenArgs), vbTextCompare) = 0 Then
            modLoggingHandler.LogInfo MODULE_NAME & ".ApplyWorkspaceOpenArgs", _
                "Skipped duplicate workspace OpenArgs='" & Replace(normalizedOpenArgs, "'", "''") & "'."
            Exit Sub
        End If
    End If

    If IsNewOrderOpenArgs(normalizedOpenArgs) Then
        ActivateTemporaryOrderContext ParseTemporaryOrderId(normalizedOpenArgs)
        m_workspaceOpenArgsApplied = True
        m_lastAppliedWorkspaceOpenArgs = normalizedOpenArgs
        Exit Sub
    End If

    OrderId = NzLong(normalizedOpenArgs, 0)
    m_workspaceOpenArgsApplied = False
    m_isTemporaryOrder = False
    m_temporaryOrderId = 0
    m_temporaryOrderDisposed = False
    Me.FilterOn = False
    Me.Filter = vbNullString
    Me.RecordSource = TABLE_ORD_ORDER

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyWorkspaceOpenArgs", _
        "Parsed order_id=" & CStr(OrderId) & "; raw_open_args='" & Replace(normalizedOpenArgs, "'", "''") & "'."

    If OrderId <= 0 Then
        HandleInvalidOrderContext "MSG.ORDER_DETAIL_INVALID_ID", "Die uebergebene Bestell-ID ist ungueltig."
        Exit Sub
    End If

    If Not modOrderRepository.OrderExists(OrderId) Then
        HandleInvalidOrderContext "MSG.ORDER_DETAIL_NOT_FOUND", "Die Bestellung konnte nicht gefunden werden."
        Exit Sub
    End If

    RestoreCurrentOrderSelection OrderId
    ConfigureOrderLinesSubform
    ConfigureHeaderControls
    RequeryOrderLinesSubform
    UpdateOrderLinesAvailability
    UpdateHeaderCaptionUi
    m_workspaceOpenArgsApplied = True
    m_lastAppliedWorkspaceOpenArgs = normalizedOpenArgs

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
    If m_isTemporaryOrder Then
        PersistTemporaryOrderAndReload
    ElseIf ApplyCurrentRecord() Then
        ReturnToPreviousWorkspace "DetailSave"
    End If
End Sub

Public Sub DetailApply()
    Call ApplyCurrentRecord
End Sub

Public Sub DetailCancel()
    CancelCurrentEdit
End Sub

Public Sub RefreshCalculatedTotals(Optional ByVal requeryLines As Boolean = False)
    On Error GoTo ErrorHandler

    Dim OrderId As Long

    OrderId = GetCurrentOrderId()

    Me.Requery
    If m_isTemporaryOrder And m_temporaryOrderId > 0 Then
        RestoreCurrentTemporaryOrderSelection m_temporaryOrderId
    ElseIf OrderId > 0 Then
        RestoreCurrentOrderSelection OrderId
    End If

    If requeryLines Then
        RequeryOrderLinesSubform
    End If

    RefreshVatModeDisplay
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

Private Sub ConfigureHeaderControls()
    On Error GoTo SafeExit

    ConfigureAddressCombo "cboInvoiceAddressId", FIELD_INVOICE_ADDRESS_ID
    ConfigureAddressCombo "cboDeliveryAddressId", FIELD_DELIVERY_ADDRESS_ID
    ConfigurePaymentTermCombo
    ConfigureCurrencyCombo
    ConfigureLanguageCombo
    ConfigureVatModeCombo
    EnsureControlTranslationTag "cboInvoiceAddressId", "FORM.FRMORDERDETAIL.CBOINVOICEADDRESSID"
    EnsureControlTranslationTag "cboDeliveryAddressId", "FORM.FRMORDERDETAIL.CBODELIVERYADDRESSID"
    EnsureControlTranslationTag "cboTenPaymentTerm", "FORM.FRMORDERDETAIL.CBOTENPAYMENTTERM"
    EnsureControlTranslationTag "cboCurrencyCode", "FORM.FRMORDERDETAIL.CBOCURRENCYCODE"
    EnsureControlTranslationTag "cboLanguageCode", "FORM.FRMORDERDETAIL.CBOLANGUAGECODE"
    EnsureControlTranslationTag "cboVatMode", "FORM.FRMORDERDETAIL.CBOVATMODE"
    EnsureControlTranslationTag "lblInvoiceAddressId", "FORM.FRMORDERDETAIL.LBLINVOICEADDRESSID"
    EnsureControlTranslationTag "lblDeliveryAddressId", "FORM.FRMORDERDETAIL.LBLDELIVERYADDRESSID"
    EnsureControlTranslationTag "lblPaymentTermCode", "FORM.FRMORDERDETAIL.LBLPAYMENTTERMCODE"
    EnsureControlTranslationTag "lblCurrencyCode", "FORM.FRMORDERDETAIL.LBLCURRENCYCODE"
    EnsureControlTranslationTag "lblLanguageCode", "FORM.FRMORDERDETAIL.LBLLANGUAGECODE"
    EnsureControlTranslationTag "lblVatMode", "FORM.FRMORDERDETAIL.LBLVATMODE"

    SetControlVisibleIfPresent "txtCustomerAddressId", False
    SetControlVisibleIfPresent FIELD_CUSTOMER_ADDRESS_ID, False
    SetControlVisibleIfPresent "txtCustomerName", False
    SetControlVisibleIfPresent FIELD_CUSTOMER_NAME, False
    SetControlVisibleIfPresent "txtPaymentTermCode", False
    SetControlVisibleIfPresent "txtCurrencyCode", False
    SetControlVisibleIfPresent "txtLanguageCode", False

SafeExit:
End Sub

Private Sub ConfigureOrderLinesSubform()
    On Error GoTo SafeExit

    If Not HasControl(ORDER_LINES_SUBFORM_CONTROL) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ConfigureOrderLinesSubform", _
            "Subform control '" & ORDER_LINES_SUBFORM_CONTROL & "' was not found."
        Exit Sub
    End If

    Me.Controls(ORDER_LINES_SUBFORM_CONTROL).SourceObject = "Form." & ORDER_LINES_SUBFORM_NAME
    If m_isTemporaryOrder Then
        Me.Controls(ORDER_LINES_SUBFORM_CONTROL).LinkMasterFields = FIELD_TMP_ORDER_ID
        Me.Controls(ORDER_LINES_SUBFORM_CONTROL).LinkChildFields = FIELD_TMP_ORDER_ID
    Else
        Me.Controls(ORDER_LINES_SUBFORM_CONTROL).LinkMasterFields = FIELD_ORDER_ID
        Me.Controls(ORDER_LINES_SUBFORM_CONTROL).LinkChildFields = FIELD_ORDER_ID
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".ConfigureOrderLinesSubform", _
        "Configured subform control='" & ORDER_LINES_SUBFORM_CONTROL & _
        "'; source_object='" & Me.Controls(ORDER_LINES_SUBFORM_CONTROL).SourceObject & _
        "'; link_master_fields='" & Me.Controls(ORDER_LINES_SUBFORM_CONTROL).LinkMasterFields & _
        "'; link_child_fields='" & Me.Controls(ORDER_LINES_SUBFORM_CONTROL).LinkChildFields & "'."

    ConfigureOrderLinesSubformInstance
    UpdateOrderLinesAvailability

SafeExit:
End Sub

Private Sub RequeryOrderLinesSubform()
    On Error GoTo SafeExit

    If Not HasControl(ORDER_LINES_SUBFORM_CONTROL) Then
        Exit Sub
    End If

    ConfigureOrderLinesSubformInstance
    Me.Controls(ORDER_LINES_SUBFORM_CONTROL).Form.Requery

SafeExit:
End Sub

Private Sub ConfigureAddressCombo(ByVal controlName As String, ByVal controlSource As String)
    On Error GoTo SafeExit

    If Not HasControl(controlName) Then
        Exit Sub
    End If

    With Me.Controls(controlName)
        .ControlSource = controlSource
        .RowSourceType = "Table/Query"
        .RowSource = modOrderRepository.GetAddressRowSourceSql()
        .BoundColumn = 1
        .ColumnCount = 4
        .ColumnWidths = "0cm;5cm;3cm;1.5cm"
        .LimitToList = True
    End With

SafeExit:
End Sub

Private Sub ConfigureVatModeCombo()
    On Error GoTo SafeExit

    If Not HasControl("cboVatMode") Then
        Exit Sub
    End If

    With Me.Controls("cboVatMode")
        .ControlSource = FIELD_VAT_MODE
        .RowSourceType = "Value List"
        .RowSource = BuildVatModeRowSource()
        .BoundColumn = 1
        .ColumnCount = 2
        .ColumnWidths = "0cm;5cm"
        .LimitToList = True
    End With

SafeExit:
End Sub

Private Sub ConfigurePaymentTermCombo()
    On Error GoTo SafeExit

    If Not HasControl("cboTenPaymentTerm") Then
        Exit Sub
    End If

    With Me.Controls("cboTenPaymentTerm")
        .ControlSource = FIELD_PAYMENT_TERM_CODE
        .RowSourceType = "Table/Query"
        .RowSource = BuildPaymentTermRowSourceSql()
        .BoundColumn = 1
        .ColumnCount = 2
        .ColumnWidths = "0cm;6cm"
        .LimitToList = True
    End With

SafeExit:
End Sub

Private Sub PrimePaymentTermComboEarly()
    On Error GoTo SafeExit

    If Not HasControl("cboTenPaymentTerm") Then
        Exit Sub
    End If

    With Me.Controls("cboTenPaymentTerm")
        .RowSourceType = "Table/Query"
        .RowSource = "SELECT payment_term_code, payment_term_code " & _
                     "FROM ten_payment_term " & _
                     "WHERE Len(Trim(Nz([payment_term_code], ''))) > 0 " & _
                     "AND Nz([is_active], True)=True " & _
                     "ORDER BY Nz([sort_order], 0), Nz([payment_term_code], '');"
        .BoundColumn = 1
        .ColumnCount = 2
        .ColumnWidths = "0cm;6cm"
        .LimitToList = True
    End With

SafeExit:
End Sub

Private Sub RefreshPaymentTermDisplay()
    On Error GoTo SafeExit

    Dim currentPaymentTermCode As String

    If Not HasControl("cboTenPaymentTerm") Then
        Exit Sub
    End If

    currentPaymentTermCode = Trim$(GetControlTextValue("cboTenPaymentTerm", FIELD_PAYMENT_TERM_CODE))
    ConfigurePaymentTermCombo
    If LenB(currentPaymentTermCode) > 0 Then
        Me.Controls("cboTenPaymentTerm").Value = currentPaymentTermCode
    End If

SafeExit:
End Sub

Private Sub ConfigureCurrencyCombo()
    On Error GoTo SafeExit

    If Not HasControl("cboCurrencyCode") Then
        Exit Sub
    End If

    With Me.Controls("cboCurrencyCode")
        .ControlSource = FIELD_CURRENCY_CODE
        .RowSourceType = "Table/Query"
        .RowSource = BuildCurrencyRowSourceSql()
        .BoundColumn = 1
        .ColumnCount = 1
        .ColumnWidths = "2.5cm"
        .LimitToList = True
    End With

SafeExit:
End Sub

Private Sub ConfigureLanguageCombo()
    On Error GoTo SafeExit

    If Not HasControl("cboLanguageCode") Then
        Exit Sub
    End If

    With Me.Controls("cboLanguageCode")
        .ControlSource = FIELD_LANGUAGE_CODE
        .RowSourceType = "Table/Query"
        .RowSource = BuildLanguageRowSourceSql()
        .BoundColumn = 1
        .ColumnCount = 2
        .ColumnWidths = "0cm;5cm"
        .LimitToList = True
    End With

SafeExit:
End Sub

Private Sub ApplyDeliveryAddressVatDefaults()
    On Error GoTo ErrorHandler

    Dim vatContext As Object
    Dim OrderId As Long

    Set vatContext = modOrderRepository.GetDefaultVatContextForOrder(GetCurrentCustomerAddressId(), GetCurrentDeliveryAddressId())
    If vatContext Is Nothing Then
        Exit Sub
    End If

    SetControlValueIfPresent FIELD_VAT_MODE, GetDictionaryString(vatContext, FIELD_VAT_MODE)
    SetControlValueIfPresent "cboVatMode", GetDictionaryString(vatContext, FIELD_VAT_MODE)
    SetControlValueIfPresent FIELD_VAT_CODE, GetDictionaryString(vatContext, FIELD_VAT_CODE)
    SetControlValueIfPresent "txtVatCode", GetDictionaryString(vatContext, FIELD_VAT_CODE)
    SetControlValueIfPresent FIELD_VAT_RATE, GetDictionaryDouble(vatContext, FIELD_VAT_RATE, 0)
    SetControlValueIfPresent "txtVatRate", GetDictionaryDouble(vatContext, FIELD_VAT_RATE, 0)
    RefreshVatModeDisplay

    OrderId = GetCurrentOrderId()
    If m_isTemporaryOrder And m_temporaryOrderId > 0 Then
        modOrderCalculationService.RecalculateTemporaryOrder m_temporaryOrderId
        RefreshCalculatedTotals True
    ElseIf OrderId > 0 Then
        modOrderCalculationService.RecalculateOrder OrderId
        RefreshCalculatedTotals True
    Else
        RequeryOrderLinesSubform
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyDeliveryAddressVatDefaults", Err
End Sub

Private Function ApplyCurrentRecord() As Boolean
    On Error GoTo ErrorHandler

    Dim OrderId As Long

    ApplyCurrentRecord = False

    OrderId = GetCurrentOrderId()

    If Me.Dirty Then
        modAuditHelper.ApplyAuditFields Me
        Me.Dirty = False
        OrderId = GetCurrentOrderId()
    End If

    If m_isTemporaryOrder And m_temporaryOrderId > 0 Then
        Call modOrderCalculationService.RecalculateTemporaryOrder(m_temporaryOrderId)
    ElseIf OrderId > 0 Then
        Call modOrderCalculationService.RecalculateOrder(OrderId)
    End If

    RefreshCalculatedTotals True
    UpdateOrderLinesAvailability

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyCurrentRecord", _
        "Sales order applied for order_id=" & CStr(OrderId) & "."

    ApplyCurrentRecord = True
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyCurrentRecord", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.ORDER_DETAIL_SAVE_ERROR", "Fehler beim Speichern der Bestellung."), vbExclamation, MODULE_NAME
End Function

Private Sub CancelCurrentEdit( _
    Optional ByVal showConfirmation As Boolean = True, _
    Optional ByVal navigateBack As Boolean = True)
    On Error GoTo ErrorHandler

    If m_isTemporaryOrder Then
        If showConfirmation Then
            If MsgBox(modFwTranslationRuntime.ResolveText("MSG.ORDER_DETAIL_CANCEL_CONFIRM", "Aenderungen verwerfen?"), vbQuestion + vbYesNo + vbDefaultButton2, MODULE_NAME) <> vbYes Then
                Exit Sub
            End If
        End If

        If m_temporaryOrderId > 0 Then
            Call modOrderRepository.DeleteTemporaryOrder(m_temporaryOrderId)
        End If
        m_temporaryOrderDisposed = True
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
    UpdateOrderLinesAvailability
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

    displayName = Trim$(GetControlTextValue(FIELD_CUSTOMER_NAME, "txtCustomerName"))

    If GetCurrentOrderId() > 0 Or m_isTemporaryOrder Then
        If LenB(displayName) = 0 Then
            displayName = Trim$(GetControlTextValue("order_no", "txtOrderNo"))
        End If
        If LenB(displayName) = 0 And GetCurrentOrderId() > 0 Then
            displayName = CStr(GetCurrentOrderId())
        End If

        If LenB(displayName) > 0 Then
            headerCaption = "Bestellung fuer: " & displayName
        End If
    End If

    If HasControl("lblFormTitle") Then
        Me!lblFormTitle.Caption = headerCaption
    ElseIf HasControl("txtFormTitle") Then
        Me!txtFormTitle.Value = headerCaption
    End If

SafeExit:
End Sub

Private Sub UpdateOrderLinesAvailability()
    On Error GoTo SafeExit

    Dim canEditLines As Boolean

    canEditLines = (GetCurrentOrderHeaderId() > 0)

    If HasControl(ORDER_LINES_SUBFORM_CONTROL) Then
        Me.Controls(ORDER_LINES_SUBFORM_CONTROL).Enabled = canEditLines
        If LenB(Me.Controls(ORDER_LINES_SUBFORM_CONTROL).SourceObject) > 0 Then
            Me.Controls(ORDER_LINES_SUBFORM_CONTROL).Form.SetOrderEntryEnabled canEditLines
        End If
    End If

    SetControlVisibleIfPresent "lblOrderLinesHint", (Not canEditLines)

SafeExit:
End Sub

Private Sub UpdateLocalCommandButtons()
    On Error GoTo SafeExit

    If HasControl("cmdSave") Then
        Me!cmdSave.Visible = False
    End If

    If HasControl("cmdApply") Then
        Me!cmdApply.Visible = False
    End If

    If HasControl("cmdCancel") Then
        Me!cmdCancel.Visible = False
    End If

SafeExit:
End Sub

Private Function GetCurrentOrderId() As Long
    On Error GoTo SafeExit

    If Me.NewRecord Then
        If RecordsetHasField(FIELD_ORDER_ID) Then
            GetCurrentOrderId = NzLong(Me.Recordset.Fields(FIELD_ORDER_ID).Value, 0)
        End If
        Exit Function
    End If

    If HasControl("txtOrderId") Then
        GetCurrentOrderId = NzLong(Me!txtOrderId.Value, 0)
    ElseIf HasControl(FIELD_ORDER_ID) Then
        GetCurrentOrderId = NzLong(Me.Controls(FIELD_ORDER_ID).Value, 0)
    ElseIf RecordsetHasField(FIELD_ORDER_ID) Then
        GetCurrentOrderId = NzLong(Me.Recordset.Fields(FIELD_ORDER_ID).Value, 0)
    End If

SafeExit:
End Function

Private Function IsNewOrderOpenArgs(ByVal openArgs As String) As Boolean
    IsNewOrderOpenArgs = (StrComp(Left$(Trim$(openArgs), Len(OPEN_ARGS_NEW_ORDER_PREFIX)), OPEN_ARGS_NEW_ORDER_PREFIX, vbTextCompare) = 0)
End Function

Private Function ParseTemporaryOrderId(ByVal openArgs As String) As Long
    Dim parts() As String

    parts = Split(Trim$(openArgs), ";")
    If UBound(parts) >= 1 Then
        ParseTemporaryOrderId = NzLong(parts(1), 0)
    End If
End Function

Private Function GetCurrentTemporaryOrderId() As Long
    On Error GoTo SafeExit

    If HasControl(FIELD_TMP_ORDER_ID) Then
        GetCurrentTemporaryOrderId = NzLong(Me.Controls(FIELD_TMP_ORDER_ID).Value, 0)
    ElseIf RecordsetHasField(FIELD_TMP_ORDER_ID) Then
        GetCurrentTemporaryOrderId = NzLong(Me.Recordset.Fields(FIELD_TMP_ORDER_ID).Value, 0)
    End If

SafeExit:
End Function

Public Function GetCurrentOrderHeaderId() As Long
    If m_isTemporaryOrder Then
        GetCurrentOrderHeaderId = GetCurrentTemporaryOrderId()
    Else
        GetCurrentOrderHeaderId = GetCurrentOrderId()
    End If
End Function

Public Function RecalculateCurrentOrderFromChild() As Boolean
    If m_isTemporaryOrder Then
        RecalculateCurrentOrderFromChild = modOrderCalculationService.RecalculateTemporaryOrder(GetCurrentTemporaryOrderId())
    Else
        RecalculateCurrentOrderFromChild = modOrderCalculationService.RecalculateOrder(GetCurrentOrderId())
    End If
End Function

Public Function RecalculateCurrentOrderTotalsOnly() As Boolean
    If m_isTemporaryOrder Then
        RecalculateCurrentOrderTotalsOnly = modOrderCalculationService.CalculateTemporaryOrderTotals(GetCurrentTemporaryOrderId())
    Else
        RecalculateCurrentOrderTotalsOnly = modOrderCalculationService.CalculateOrderTotals(GetCurrentOrderId())
    End If
End Function

Private Sub ActivateTemporaryOrderContext(ByVal tmpOrderId As Long)
    On Error GoTo ErrorHandler

    If tmpOrderId <= 0 Then
        HandleInvalidOrderContext "MSG.ORDER_DETAIL_INVALID_ID", "Die uebergebene Bestell-ID ist ungueltig."
        Exit Sub
    End If

    If Not modOrderRepository.TemporaryOrderExists(tmpOrderId) Then
        HandleInvalidOrderContext "MSG.ORDER_DETAIL_NOT_FOUND", "Die Bestellung konnte nicht gefunden werden."
        Exit Sub
    End If

    m_isTemporaryOrder = True
    m_temporaryOrderId = tmpOrderId
    m_temporaryOrderDisposed = False
    Me.FilterOn = False
    Me.Filter = vbNullString
    Me.RecordSource = TABLE_TMP_ORDER
    Me.Requery
    RestoreCurrentTemporaryOrderSelection tmpOrderId
    ConfigureOrderLinesSubform
    ConfigureHeaderControls
    UpdateOrderLinesAvailability
    RequeryOrderLinesSubform
    UpdateHeaderCaptionUi

    modLoggingHandler.LogInfo MODULE_NAME & ".ActivateTemporaryOrderContext", _
        "Temporary order context activated for tmp_order_id=" & CStr(tmpOrderId) & "."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ActivateTemporaryOrderContext", Err
End Sub

Private Function ResolveCurrentOrderDate() As Date
    On Error GoTo SafeExit

    If HasControl(FIELD_ORDER_DATE) Then
        If IsDate(Me.Controls(FIELD_ORDER_DATE).Value) Then
            ResolveCurrentOrderDate = CDate(Me.Controls(FIELD_ORDER_DATE).Value)
            Exit Function
        End If
    ElseIf RecordsetHasField(FIELD_ORDER_DATE) Then
        If IsDate(Me.Recordset.Fields(FIELD_ORDER_DATE).Value) Then
            ResolveCurrentOrderDate = CDate(Me.Recordset.Fields(FIELD_ORDER_DATE).Value)
            Exit Function
        End If
    End If

SafeExit:
    If ResolveCurrentOrderDate = 0 Then
        ResolveCurrentOrderDate = Date
    End If
End Function

Private Sub ConfigureOrderLinesSubformInstance()
    On Error GoTo SafeExit

    If Not HasControl(ORDER_LINES_SUBFORM_CONTROL) Then
        Exit Sub
    End If

    If LenB(Me.Controls(ORDER_LINES_SUBFORM_CONTROL).SourceObject) = 0 Then
        Exit Sub
    End If

    If m_isTemporaryOrder Then
        Me.Controls(ORDER_LINES_SUBFORM_CONTROL).Form.ConfigureForOrderContext TABLE_TMP_ORDER_LINE, FIELD_TMP_ORDER_ID, FIELD_TMP_ORDER_LINE_ID
    Else
        Me.Controls(ORDER_LINES_SUBFORM_CONTROL).Form.ConfigureForOrderContext TABLE_ORD_ORDER_LINE, FIELD_ORDER_ID, FIELD_ORDER_LINE_ID
    End If

SafeExit:
End Sub

Private Sub PersistTemporaryOrderAndReload()
    On Error GoTo ErrorHandler

    Dim OrderId As Long

    If Me.Dirty Then
        modAuditHelper.ApplyAuditFields Me
        Me.Dirty = False
    End If

    OrderId = modOrderRepository.PersistTemporaryOrder(m_temporaryOrderId)
    If OrderId <= 0 Then
        Err.Raise vbObjectError + 5480, MODULE_NAME & ".PersistTemporaryOrderAndReload", "Temporary order could not be persisted."
    End If

    m_temporaryOrderDisposed = True
    m_isTemporaryOrder = False
    m_temporaryOrderId = 0

    If Not modAppWorkspaceService.OpenWorkspaceForm(Me, "frmOrderDetail", vbNullString, acFormPropertySettings, CStr(OrderId)) Then
        Err.Raise vbObjectError + 5481, MODULE_NAME & ".PersistTemporaryOrderAndReload", "Workspace navigation to persisted frmOrderDetail failed."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "PersistTemporaryOrderAndReload", Err
End Sub

Public Function GetCurrentOrderVatCode() As String
    GetCurrentOrderVatCode = Trim$(GetControlTextValue(FIELD_VAT_CODE, "txtVatCode"))
End Function

Public Function GetCurrentOrderVatRate() As Double
    GetCurrentOrderVatRate = GetControlDoubleValue(FIELD_VAT_RATE, "txtVatRate")
End Function

Private Function GetCurrentCustomerAddressId() As Long
    GetCurrentCustomerAddressId = GetControlLongValue(FIELD_CUSTOMER_ADDRESS_ID, "txtCustomerAddressId")
End Function

Private Function GetCurrentDeliveryAddressId() As Long
    GetCurrentDeliveryAddressId = GetControlLongValue(FIELD_DELIVERY_ADDRESS_ID, "cboDeliveryAddressId")
    If GetCurrentDeliveryAddressId <= 0 Then
        GetCurrentDeliveryAddressId = GetCurrentCustomerAddressId()
    End If
End Function

Private Sub RestoreCurrentTemporaryOrderSelection(ByVal tmpOrderId As Long)
    On Error GoTo ErrorHandler

    Dim recordsetClone As DAO.Recordset

    Set recordsetClone = Me.RecordsetClone
    If recordsetClone.BOF And recordsetClone.EOF Then
        GoTo CleanExit
    End If

    recordsetClone.FindFirst "[" & FIELD_TMP_ORDER_ID & "]=" & CStr(tmpOrderId)
    If Not recordsetClone.NoMatch Then
        Me.Bookmark = recordsetClone.Bookmark
    Else
        HandleInvalidOrderContext "MSG.ORDER_DETAIL_NOT_FOUND", "Die Bestellung konnte nicht gefunden werden."
    End If

CleanExit:
    On Error Resume Next
    If Not recordsetClone Is Nothing Then recordsetClone.Close
    Set recordsetClone = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreCurrentTemporaryOrderSelection", Err
    Resume CleanExit
End Sub

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

Private Function BuildPaymentTermRowSourceSql() As String
    BuildPaymentTermRowSourceSql = _
        "SELECT payment_term_code, " & _
        "ResolveText('PAYMENT_TERM.' & Nz([payment_term_code],'') & '.TITLE', Nz([payment_term_code],''), " & _
        SqlText(ResolveCurrentDocumentLanguageCode()) & ") AS payment_term_title " & _
        "FROM ten_payment_term " & _
        "WHERE Len(Trim(Nz([payment_term_code], ''))) > 0 " & _
        "AND Nz([is_active], True)=True " & _
        "ORDER BY Nz([sort_order], 0), Nz([payment_term_code], '');"
End Function

Private Sub EnsureDefaultPaymentTerm()
    On Error GoTo SafeExit

    Dim currentPaymentTermCode As String
    Dim defaultPaymentTermCode As String

    currentPaymentTermCode = Trim$(GetControlTextValue("cboTenPaymentTerm", FIELD_PAYMENT_TERM_CODE))
    If LenB(currentPaymentTermCode) > 0 Then
        Exit Sub
    End If

    defaultPaymentTermCode = ResolveDefaultPaymentTermCode()
    If LenB(defaultPaymentTermCode) = 0 Then
        Exit Sub
    End If

    SetControlValueIfPresent FIELD_PAYMENT_TERM_CODE, defaultPaymentTermCode
    SetControlValueIfPresent "cboTenPaymentTerm", defaultPaymentTermCode

SafeExit:
End Sub

Private Function ResolveDefaultPaymentTermCode() As String
    On Error GoTo SafeExit

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = modDb.GetCurrentTenantDatabase()
    If db Is Nothing Then
        GoTo SafeExit
    End If
    Set rs = db.OpenRecordset( _
        "SELECT TOP 1 payment_term_code " & _
        "FROM ten_payment_term " & _
        "WHERE Nz(is_active, True)=True " & _
        "AND Nz(is_default, False)=True " & _
        "ORDER BY Nz(sort_order, 0), Nz(payment_term_code, '');", _
        dbOpenSnapshot)

    If Not (rs.BOF And rs.EOF) Then
        ResolveDefaultPaymentTermCode = Trim$(modDaoHelper.NzString(rs.Fields("payment_term_code").Value, vbNullString))
    End If

SafeExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

Private Function ResolveCurrentDocumentLanguageCode() As String
    Dim languageCode As String

    languageCode = Trim$(GetControlTextValue(FIELD_LANGUAGE_CODE, "cboLanguageCode"))
    If LenB(languageCode) = 0 Then
        languageCode = Trim$(modTenantRepository.GetTenantParameter("DEFAULT_LANGUAGE", "en-US"))
    End If

    languageCode = Trim$(modFwTranslationRuntime.NormalizeProjectLanguageCode(languageCode))
    If LenB(languageCode) = 0 Then
        languageCode = "en-US"
    End If

    ResolveCurrentDocumentLanguageCode = languageCode
End Function

Private Function BuildCurrencyRowSourceSql() As String
    BuildCurrencyRowSourceSql = _
        "SELECT currency_code " & _
        "FROM ref_currency " & _
        "WHERE Nz(is_active, True)=True " & _
        "ORDER BY currency_code;"
End Function

Private Function BuildVatModeRowSource() As String
    BuildVatModeRowSource = _
        "EXCLUSIVE;" & ResolveVatModeDisplayText("EXCLUSIVE") & ";" & _
        "INCLUSIVE;" & ResolveVatModeDisplayText("INCLUSIVE") & ";" & _
        "NONE;" & ResolveVatModeDisplayText("NONE")
End Function

Private Function BuildLanguageRowSourceSql() As String
    BuildLanguageRowSourceSql = _
        "SELECT language_code, language_name " & _
        "FROM ref_language " & _
        "WHERE Nz(is_active, True)=True " & _
        "ORDER BY Nz(sort_order, 0), Nz(language_name, '');"
End Function

Private Sub RefreshVatModeDisplay()
    On Error GoTo SafeExit

    If Not HasControl("cboVatMode") Then
        Exit Sub
    End If

    Me.Controls("cboVatMode").RowSource = BuildVatModeRowSource()

SafeExit:
End Sub

Private Function ResolveVatModeDisplayText(ByVal vatMode As String) As String
    Select Case UCase$(Trim$(vatMode))
        Case "INCLUSIVE"
            ResolveVatModeDisplayText = modFwTranslationRuntime.ResolveText( _
                "FORM.FRMORDERDETAIL.VAT_MODE_INCLUSIVE", _
                "Inklusive MwSt.")
        Case "NONE"
            ResolveVatModeDisplayText = modFwTranslationRuntime.ResolveText( _
                "FORM.FRMORDERDETAIL.VAT_MODE_NONE", _
                "Keine MwSt.")
        Case Else
            ResolveVatModeDisplayText = modFwTranslationRuntime.ResolveText( _
                "FORM.FRMORDERDETAIL.VAT_MODE_EXCLUSIVE", _
                "Exklusive MwSt.")
    End Select
End Function

Private Function ResolveExistingFieldName(ByVal tableName As String, ByVal candidateFieldNames As Variant) As String
    On Error GoTo SafeExit

    Dim db As DAO.Database
    Dim tableDefinition As DAO.TableDef
    Dim fieldDefinition As DAO.Field
    Dim candidateFieldName As Variant

    Set db = CurrentDb
    Set tableDefinition = db.TableDefs(tableName)

    For Each candidateFieldName In candidateFieldNames
        For Each fieldDefinition In tableDefinition.Fields
            If StrComp(Trim$(fieldDefinition.Name), Trim$(CStr(candidateFieldName)), vbTextCompare) = 0 Then
                ResolveExistingFieldName = fieldDefinition.Name
                GoTo SafeExit
            End If
        Next fieldDefinition
    Next candidateFieldName

SafeExit:
    Set fieldDefinition = Nothing
    Set tableDefinition = Nothing
    Set db = Nothing
End Function

Private Sub SetControlVisibleIfPresent(ByVal ControlName As String, ByVal visibleState As Boolean)
    On Error GoTo SafeExit

    If HasControl(ControlName) Then
        Me.Controls(ControlName).Visible = visibleState
    End If

SafeExit:
End Sub

Private Sub SetControlValueIfPresent(ByVal ControlName As String, ByVal Value As Variant)
    On Error GoTo SafeExit

    If HasControl(ControlName) Then
        Me.Controls(ControlName).Value = Value
    ElseIf RecordsetHasField(ControlName) Then
        Me.Recordset.Fields(ControlName).Value = Value
    End If

SafeExit:
End Sub

Private Sub EnsureControlTranslationTag(ByVal ControlName As String, ByVal translationKey As String)
    On Error GoTo SafeExit

    If HasControl(ControlName) Then
        Me.Controls(ControlName).Tag = modFwTranslationRuntime.SetTranslationKeyInTag( _
            modDaoHelper.NzString(Me.Controls(ControlName).Tag), _
            translationKey)
    End If

SafeExit:
End Sub

Private Function GetControlTextValue(ByVal primaryControlName As String, Optional ByVal fallbackControlName As String = "") As String
    On Error GoTo SafeExit

    If HasControl(primaryControlName) Then
        GetControlTextValue = NzString(Me.Controls(primaryControlName).Value)
        Exit Function
    ElseIf RecordsetHasField(primaryControlName) Then
        GetControlTextValue = NzString(Me.Recordset.Fields(primaryControlName).Value)
        Exit Function
    End If

    If LenB(fallbackControlName) > 0 Then
        If HasControl(fallbackControlName) Then
            GetControlTextValue = NzString(Me.Controls(fallbackControlName).Value)
        End If
    End If

SafeExit:
End Function

Private Function GetControlLongValue(ByVal primaryControlName As String, Optional ByVal fallbackControlName As String = "") As Long
    On Error GoTo SafeExit

    If HasControl(primaryControlName) Then
        GetControlLongValue = NzLong(Me.Controls(primaryControlName).Value, 0)
        Exit Function
    ElseIf RecordsetHasField(primaryControlName) Then
        GetControlLongValue = NzLong(Me.Recordset.Fields(primaryControlName).Value, 0)
        Exit Function
    End If

    If LenB(fallbackControlName) > 0 Then
        If HasControl(fallbackControlName) Then
            GetControlLongValue = NzLong(Me.Controls(fallbackControlName).Value, 0)
        End If
    End If

SafeExit:
End Function

Private Function GetControlDoubleValue(ByVal primaryControlName As String, Optional ByVal fallbackControlName As String = "") As Double
    On Error GoTo SafeExit

    If HasControl(primaryControlName) Then
        GetControlDoubleValue = NzDouble(Me.Controls(primaryControlName).Value, 0)
        Exit Function
    ElseIf RecordsetHasField(primaryControlName) Then
        GetControlDoubleValue = NzDouble(Me.Recordset.Fields(primaryControlName).Value, 0)
        Exit Function
    End If

    If LenB(fallbackControlName) > 0 Then
        If HasControl(fallbackControlName) Then
            GetControlDoubleValue = NzDouble(Me.Controls(fallbackControlName).Value, 0)
        End If
    End If

SafeExit:
End Function

Private Function RecordsetHasField(ByVal fieldName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentField As DAO.Field

    For Each currentField In Me.Recordset.Fields
        If StrComp(currentField.Name, fieldName, vbTextCompare) = 0 Then
            RecordsetHasField = True
            Exit Function
        End If
    Next currentField

SafeExit:
End Function

Private Function GetDictionaryString(ByVal dictionaryObject As Object, ByVal keyName As String) As String
    On Error GoTo SafeExit

    If dictionaryObject Is Nothing Then
        Exit Function
    End If

    If dictionaryObject.Exists(keyName) Then
        GetDictionaryString = NzString(dictionaryObject(keyName))
    End If

SafeExit:
End Function

Private Function GetDictionaryDouble(ByVal dictionaryObject As Object, ByVal keyName As String, ByVal defaultValue As Double) As Double
    On Error GoTo SafeExit

    Dim rawValue As String

    GetDictionaryDouble = defaultValue
    If dictionaryObject Is Nothing Then
        Exit Function
    End If

    If dictionaryObject.Exists(keyName) Then
        rawValue = NzString(dictionaryObject(keyName), CStr(defaultValue))
        If IsNumeric(rawValue) Then
            GetDictionaryDouble = CDbl(rawValue)
        End If
    End If

SafeExit:
End Function

Private Function GetDictionaryLong(ByVal dictionaryObject As Object, ByVal keyName As String, ByVal defaultValue As Long) As Long
    On Error GoTo SafeExit

    If dictionaryObject Is Nothing Then
        GetDictionaryLong = defaultValue
        Exit Function
    End If

    If dictionaryObject.Exists(keyName) Then
        GetDictionaryLong = NzLong(dictionaryObject(keyName), defaultValue)
    Else
        GetDictionaryLong = defaultValue
    End If

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

Private Function NzDouble(ByVal Value As Variant, Optional ByVal defaultValue As Double = 0) As Double
    If IsNull(Value) Or IsEmpty(Value) Or LenB(Trim$(CStr(Value))) = 0 Then
        NzDouble = defaultValue
    ElseIf IsNumeric(Value) Then
        NzDouble = CDbl(Value)
    Else
        NzDouble = defaultValue
    End If
End Function

Private Function SqlText(ByVal valueText As String) As String
    SqlText = "'" & Replace(Trim$(valueText), "'", "''") & "'"
End Function
