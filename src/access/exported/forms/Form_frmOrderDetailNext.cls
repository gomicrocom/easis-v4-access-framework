VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmOrderDetailNext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ExportKind: CODE_BEHIND_ONLY
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmOrderDetailNext
' Purpose   : Controlled Phase-1 order header editor bound only to tmp_order.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmOrderDetailNext"
Private Const TABLE_TMP_ORDER As String = "tmp_order"
Private Const DASHBOARD_FORM_NAME As String = "frmAppDashboard"
Private Const SUBFORM_ORDER_LINES As String = "subOrderLines"
Private Const SUBFORM_ORDER_LINES_SOURCE_OBJECT As String = "Form.sfrmOrderLinesNext"
Private Const RECORDSOURCE_EMPTY As String = "SELECT * FROM [tmp_order] WHERE 1 = 0"

Private Const FIELD_TMP_ORDER_ID As String = "tmp_order_id"
Private Const FIELD_ORDER_ID As String = "order_id"
Private Const FIELD_SOURCE_ORDER_ID As String = "source_order_id"
Private Const FIELD_ORDER_NO As String = "order_no"
Private Const FIELD_ORDER_TYPE_CODE As String = "order_type_code"
Private Const FIELD_ORDER_STATUS_CODE As String = "order_status_code"
Private Const FIELD_CUSTOMER_ADDRESS_ID As String = "customer_address_id"
Private Const FIELD_INVOICE_ADDRESS_ID As String = "invoice_address_id"
Private Const FIELD_DELIVERY_ADDRESS_ID As String = "delivery_address_id"
Private Const FIELD_CUSTOMER_NAME As String = "customer_name"
Private Const FIELD_ORDER_DATE As String = "order_date"
Private Const FIELD_DELIVERY_DATE As String = "delivery_date"
Private Const FIELD_VALID_UNTIL As String = "valid_until"
Private Const FIELD_REFERENCE_TEXT As String = "reference_text"
Private Const FIELD_EXTERNAL_REFERENCE As String = "external_reference"
Private Const FIELD_LANGUAGE_CODE As String = "language_code"
Private Const FIELD_CURRENCY_CODE As String = "currency_code"
Private Const FIELD_PAYMENT_TERM_CODE As String = "payment_term_code"
Private Const FIELD_VAT_MODE As String = "vat_mode"
Private Const FIELD_VAT_CODE As String = "vat_code"
Private Const FIELD_VAT_RATE As String = "vat_rate"

Private Const STATE_LOADING As String = "LOADING"
Private Const STATE_READY As String = "READY"
Private Const STATE_SAVING As String = "SAVING"
Private Const STATE_CANCELLING As String = "CANCELLING"
Private Const STATE_CLOSED As String = "CLOSED"

Private m_lifecycleState As String
Private m_initializedTmpOrderId As Long
Private m_lastAppliedPayload As String
Private m_sessionInitialized As Boolean
Private m_deleteSessionOnClose As Boolean
Private m_cleanupCompleted As Boolean
Private m_committedOrderId As Long

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    m_lifecycleState = STATE_LOADING
    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "OpenArgs raw='" & Replace(modDaoHelper.NzString(Me.OpenArgs), "'", "''") & "'."
    EnsureEmptyRecordSource
    ConfigureHeaderControls
    HidePhase2Controls
    ApplyInitialPayloadIfAvailable
    modFwTranslationRuntime.ApplyTranslations Me
    UpdateFormUi
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    UpdateFormUi
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandler

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Unload", _
        "tmp_order_id=" & CStr(m_initializedTmpOrderId) & "; delete_session_on_close=" & CStr(m_deleteSessionOnClose) & "; cleanup_completed=" & CStr(m_cleanupCompleted) & "."
    CleanupTemporarySessionIfNeeded
    m_lifecycleState = STATE_CLOSED
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Unload", Err
End Sub

Private Sub cmdApply_Click()
    On Error GoTo ErrorHandler

    DetailApply
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdApply_Click", Err
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo ErrorHandler

    DetailCancel
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdCancel_Click", Err
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler

    DetailSave
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdSave_Click", Err
End Sub

Private Sub cboLanguageCode_AfterUpdate()
    On Error GoTo ErrorHandler

    ConfigurePaymentTermCombo
    UpdateFormUi
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboLanguageCode_AfterUpdate", Err
End Sub

Private Sub cboVatMode_AfterUpdate()
    On Error GoTo ErrorHandler

    ConfigureVatModeCombo
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboVatMode_AfterUpdate", Err
End Sub

Public Sub ApplyWorkspaceOpenArgs(ByVal openArgs As String)
    On Error GoTo ErrorHandler

    InitializeFromPayload openArgs
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyWorkspaceOpenArgs", Err
End Sub

Public Sub InitializeOrderEditSession(ByVal tmpOrderId As Long)
    On Error GoTo ErrorHandler

    Dim recordsetClone As DAO.Recordset

    If tmpOrderId <= 0 Then
        HandleInvalidOrderContext "MSG.ORDER_DETAIL_INVALID_ID", "Die uebergebene Bestell-ID ist ungueltig."
        Exit Sub
    End If

    If m_sessionInitialized Then
        If m_initializedTmpOrderId = tmpOrderId Then
            Exit Sub
        End If
    End If

    If Not modOrderRepository.TemporaryOrderExists(tmpOrderId) Then
        HandleInvalidOrderContext "MSG.ORDER_DETAIL_NOT_FOUND", "Die Bestellung konnte nicht gefunden werden."
        Exit Sub
    End If

    m_lifecycleState = STATE_LOADING
    SetSessionRecordSource tmpOrderId

    Set recordsetClone = Me.RecordsetClone
    recordsetClone.FindFirst "[" & FIELD_TMP_ORDER_ID & "]=" & CStr(tmpOrderId)
    If recordsetClone.NoMatch Then
        GoTo NotFound
    End If

    Me.Bookmark = recordsetClone.Bookmark
    Me.Filter = vbNullString
    Me.FilterOn = False

    m_initializedTmpOrderId = tmpOrderId
    m_sessionInitialized = True
    m_deleteSessionOnClose = False
    m_cleanupCompleted = False
    m_committedOrderId = 0
    m_lifecycleState = STATE_READY
    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeOrderEditSession", _
        "tmp_order_id=" & CStr(tmpOrderId) & "; source_order_id=" & CStr(GetLongFieldValue(FIELD_SOURCE_ORDER_ID)) & "; order_id=" & CStr(GetLongFieldValue(FIELD_ORDER_ID)) & "."

    ConfigureHeaderControls
    UpdateFormUi
    RefreshCommandBar

CleanExit:
    On Error Resume Next
    If Not recordsetClone Is Nothing Then recordsetClone.Close
    Set recordsetClone = Nothing
    Exit Sub

NotFound:
    HandleInvalidOrderContext "MSG.ORDER_DETAIL_NOT_FOUND", "Die Bestellung konnte nicht gefunden werden."
    GoTo CleanExit

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "InitializeOrderEditSession", Err
    Resume CleanExit
End Sub

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
    cfg(modAppShell.WSLOT_R3 & ".caption") = modFwTranslationRuntime.ResolveText("COMMON.APPLY", "Uebernehmen")
    cfg(modAppShell.WSLOT_R3 & ".command") = modAppShell.WCMD_DETAIL_APPLY

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
        Case modAppShell.WCMD_DETAIL_APPLY, modAppShell.WCMD_DETAIL_CANCEL, modAppShell.WCMD_DETAIL_SAVE
            CanExecuteWorkspaceCommand = m_sessionInitialized
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
        Case modAppShell.WCMD_DETAIL_CANCEL
            DetailCancel
        Case modAppShell.WCMD_DETAIL_SAVE
            DetailSave
    End Select
End Sub

Public Function CanLeaveWorkspace() As Boolean
    On Error GoTo ErrorHandler

    If m_lifecycleState = STATE_SAVING Or m_lifecycleState = STATE_CANCELLING Then
        CanLeaveWorkspace = True
        Exit Function
    End If

    If Not m_sessionInitialized Then
        CanLeaveWorkspace = True
        Exit Function
    End If

    If HasUnsavedChanges() Then
        If MsgBox( _
            modFwTranslationRuntime.ResolveText("MSG.ORDER_DETAIL_CANCEL_CONFIRM", "Aenderungen verwerfen?"), _
            vbQuestion + vbYesNo + vbDefaultButton2, _
            MODULE_NAME) <> vbYes Then
            Exit Function
        End If
    End If

    m_deleteSessionOnClose = True
    CanLeaveWorkspace = True
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CanLeaveWorkspace", Err
End Function

Public Sub DetailApply()
    On Error GoTo ErrorHandler

    If Not SaveTemporaryHeader() Then
        Exit Sub
    End If

    m_lifecycleState = STATE_READY
    RefreshCommandBar
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "DetailApply", Err
End Sub

Public Sub DetailCancel()
    On Error GoTo ErrorHandler

    If HasUnsavedChanges() Then
        If MsgBox( _
            modFwTranslationRuntime.ResolveText("MSG.ORDER_DETAIL_CANCEL_CONFIRM", "Aenderungen verwerfen?"), _
            vbQuestion + vbYesNo + vbDefaultButton2, _
            MODULE_NAME) <> vbYes Then
            Exit Sub
        End If
    End If

    m_lifecycleState = STATE_CANCELLING
    m_deleteSessionOnClose = True

    If Not LeaveCurrentView() Then
        m_deleteSessionOnClose = False
        m_lifecycleState = STATE_READY
    End If
    Exit Sub

ErrorHandler:
    m_deleteSessionOnClose = False
    m_lifecycleState = STATE_READY
    modErrorHandler.HandleError MODULE_NAME, "DetailCancel", Err
End Sub

Public Sub DetailSave()
    On Error GoTo ErrorHandler

    Dim committedOrderId As Long

    If Not ValidateHeaderForCommit() Then
        Exit Sub
    End If

    If Not SaveTemporaryHeader() Then
        Exit Sub
    End If

    committedOrderId = modOrderRepository.CommitTemporaryOrderHeader(m_initializedTmpOrderId)
    If committedOrderId <= 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ORDER_DETAIL_SAVE_ERROR", "Fehler beim Speichern der Bestellung."), vbExclamation, MODULE_NAME
        Exit Sub
    End If

    m_committedOrderId = committedOrderId
    m_lifecycleState = STATE_SAVING
    m_deleteSessionOnClose = True

    If Not LeaveCurrentView() Then
        m_deleteSessionOnClose = False
        m_lifecycleState = STATE_READY
    End If
    Exit Sub

ErrorHandler:
    m_deleteSessionOnClose = False
    m_lifecycleState = STATE_READY
    modErrorHandler.HandleError MODULE_NAME, "DetailSave", Err
End Sub

Private Sub ApplyInitialPayloadIfAvailable()
    Dim effectivePayload As String

    effectivePayload = Trim$(modDaoHelper.NzString(Me.OpenArgs))
    If LenB(effectivePayload) = 0 Then
        effectivePayload = modAppWorkspaceService.ConsumePendingWorkspaceOpenArgs(Me.Name)
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyInitialPayloadIfAvailable", _
        "effective_payload='" & Replace(effectivePayload, "'", "''") & "'."

    If LenB(effectivePayload) = 0 Then
        EnsureEmptyRecordSource
        If ResolveHostingShellForm() Is Nothing Then
            HandleInvalidOrderContext "MSG.ORDER_DETAIL_MISSING_ID", "Keine Bestellung uebergeben."
        End If
        Exit Sub
    End If

    InitializeFromPayload effectivePayload
End Sub

Private Sub InitializeFromPayload(ByVal payloadText As String)
    Dim normalizedPayload As String
    Dim tmpOrderId As Long

    normalizedPayload = Trim$(modDaoHelper.NzString(payloadText))
    If LenB(normalizedPayload) = 0 Then
        EnsureEmptyRecordSource
        Exit Sub
    End If

    If StrComp(normalizedPayload, m_lastAppliedPayload, vbTextCompare) = 0 Then
        Exit Sub
    End If

    tmpOrderId = modOrderEditSessionService.ParseOrderEditWorkspaceId(normalizedPayload)
    If tmpOrderId <= 0 Then
        EnsureEmptyRecordSource
        HandleInvalidOrderContext "MSG.ORDER_DETAIL_INVALID_ID", "Die uebergebene Bestell-ID ist ungueltig."
        Exit Sub
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeFromPayload", _
        "payload='" & Replace(normalizedPayload, "'", "''") & "'; tmp_order_id=" & CStr(tmpOrderId) & "."
    m_lastAppliedPayload = normalizedPayload
    InitializeOrderEditSession tmpOrderId
End Sub

Private Sub EnsureEmptyRecordSource()
    If StrComp(modDaoHelper.NzString(Me.RecordSource), RECORDSOURCE_EMPTY, vbTextCompare) <> 0 Then
        Me.RecordSource = RECORDSOURCE_EMPTY
    End If
End Sub

Private Sub SetSessionRecordSource(ByVal tmpOrderId As Long)
    Dim sqlStatement As String

    sqlStatement = "SELECT * FROM [" & TABLE_TMP_ORDER & "] WHERE [" & FIELD_TMP_ORDER_ID & "]=" & CStr(tmpOrderId)
    If StrComp(modDaoHelper.NzString(Me.RecordSource), sqlStatement, vbTextCompare) <> 0 Then
        Me.RecordSource = sqlStatement
    End If
End Sub

Private Sub ConfigureHeaderControls()
    ConfigureAddressCombo "cboInvoiceAddressId"
    ConfigureAddressCombo "cboDeliveryAddressId"
    ConfigureCurrencyCombo
    ConfigureLanguageCombo
    ConfigurePaymentTermCombo
    ConfigureVatModeCombo
    HideTechnicalControls
End Sub

Private Sub ConfigureAddressCombo(ByVal controlName As String)
    WithControl controlName, True, "Table/Query", modOrderRepository.GetAddressRowSourceSql(), 4, 1, "0cm;6cm;0cm;0cm", True
End Sub

Private Sub ConfigureCurrencyCombo()
    WithControl "cboCurrencyCode", True, "Table/Query", modOrderRepository.GetCurrencyRowSourceSql(), 1, 1, "3cm", True
End Sub

Private Sub ConfigureLanguageCombo()
    WithControl "cboLanguageCode", True, "Table/Query", modOrderRepository.GetLanguageRowSourceSql(), 2, 1, "0cm;6cm", True
End Sub

Private Sub ConfigurePaymentTermCombo()
    Dim currentLanguageCode As String

    currentLanguageCode = ResolveCurrentDocumentLanguageCode()
    WithControl "cboTenPaymentTerm", True, "Table/Query", modOrderRepository.GetPaymentTermRowSourceSql(currentLanguageCode), 2, 1, "0cm;6cm", True
End Sub

Private Sub ConfigureVatModeCombo()
    WithControl "cboVatMode", True, "Value List", BuildVatModeRowSource(), 2, 1, "0cm;5cm", True
End Sub

Private Sub WithControl( _
    ByVal controlName As String, _
    ByVal visibleState As Boolean, _
    ByVal rowSourceType As String, _
    ByVal rowSource As String, _
    ByVal columnCount As Long, _
    ByVal boundColumn As Long, _
    ByVal columnWidths As String, _
    ByVal limitToList As Boolean)

    On Error GoTo SafeExit

    Dim comboControl As Access.ComboBox

    If Not HasControl(controlName) Then
        Exit Sub
    End If

    Set comboControl = Me.Controls(controlName)
    With comboControl
        .Visible = visibleState
        .RowSourceType = rowSourceType
        .RowSource = rowSource
        .ColumnCount = columnCount
        .BoundColumn = boundColumn
        .ColumnWidths = columnWidths
        .LimitToList = limitToList
    End With

SafeExit:
End Sub

Private Sub HideTechnicalControls()
    SetControlVisibleIfPresent FIELD_TMP_ORDER_ID, False
    SetControlVisibleIfPresent FIELD_ORDER_ID, False
    SetControlVisibleIfPresent FIELD_SOURCE_ORDER_ID, False
    SetControlVisibleIfPresent FIELD_CUSTOMER_ADDRESS_ID, False
    SetControlVisibleIfPresent FIELD_CUSTOMER_NAME, False
End Sub

Private Sub HidePhase2Controls()
    SetControlVisibleIfPresent SUBFORM_ORDER_LINES, False
    SetControlVisibleIfPresent "cmdApply", False
    SetControlVisibleIfPresent "cmdSave", False
    SetControlVisibleIfPresent "cmdCancel", False
End Sub

Private Sub UpdateFormUi()
    UpdateFormTitle
    UpdateHeaderCaption
    ConfigurePaymentTermCombo
    ConfigureVatModeCombo
    ConfigureOrderLinesSubform
    UpdateLocalCommandButtons
End Sub

Private Sub ConfigureOrderLinesSubform()
    On Error GoTo SafeExit

    Dim subformControl As Access.SubForm

    If Not HasControl(SUBFORM_ORDER_LINES) Then
        Exit Sub
    End If

    Set subformControl = Me.Controls(SUBFORM_ORDER_LINES)
    subformControl.Visible = m_sessionInitialized

    If Not m_sessionInitialized Then
        Exit Sub
    End If

    If StrComp(modDaoHelper.NzString(subformControl.SourceObject), SUBFORM_ORDER_LINES_SOURCE_OBJECT, vbTextCompare) <> 0 Then
        subformControl.SourceObject = SUBFORM_ORDER_LINES_SOURCE_OBJECT
    End If

    If StrComp(modDaoHelper.NzString(subformControl.LinkMasterFields), FIELD_TMP_ORDER_ID, vbTextCompare) <> 0 Then
        subformControl.LinkMasterFields = FIELD_TMP_ORDER_ID
    End If

    If StrComp(modDaoHelper.NzString(subformControl.LinkChildFields), FIELD_TMP_ORDER_ID, vbTextCompare) <> 0 Then
        subformControl.LinkChildFields = FIELD_TMP_ORDER_ID
    End If

SafeExit:
End Sub

Private Sub UpdateFormTitle()
    Dim captionText As String

    If IsExistingOrderSession() Then
        captionText = modFwTranslationRuntime.ResolveText("FORM.FRMORDERDETAILNEXT.FORM_TITLE_EDIT", "Bestellung bearbeiten")
    Else
        captionText = modFwTranslationRuntime.ResolveText("FORM.FRMORDERDETAILNEXT.FORM_TITLE_NEW", "Neue Bestellung")
    End If

    If HasControl("lblFormTitle") Then
        Me.Controls("lblFormTitle").Caption = captionText
    ElseIf HasControl("txtFormTitle") Then
        Me.Controls("txtFormTitle").Value = captionText
    End If
End Sub

Private Sub UpdateHeaderCaption()
    Dim customerDisplayName As String

    customerDisplayName = Trim$(GetTextFieldValue(FIELD_CUSTOMER_NAME))
    If LenB(customerDisplayName) = 0 Then
        customerDisplayName = CStr(GetLongFieldValue(FIELD_CUSTOMER_ADDRESS_ID))
    End If

    If HasControl("lblHeaderCaption") Then
        If LenB(customerDisplayName) > 0 Then
            Me.Controls("lblHeaderCaption").Caption = customerDisplayName
        End If
    End If
End Sub

Private Sub UpdateLocalCommandButtons()
    On Error GoTo SafeExit

    If HasControl("cmdApply") Then Me.Controls("cmdApply").Visible = False
    If HasControl("cmdSave") Then Me.Controls("cmdSave").Visible = False
    If HasControl("cmdCancel") Then Me.Controls("cmdCancel").Visible = False

SafeExit:
End Sub

Private Function BuildVatModeRowSource() As String
    BuildVatModeRowSource = _
        "EXCLUSIVE;" & ResolveVatModeDisplayText("EXCLUSIVE") & ";" & _
        "INCLUSIVE;" & ResolveVatModeDisplayText("INCLUSIVE") & ";" & _
        "NONE;" & ResolveVatModeDisplayText("NONE")
End Function

Private Function ResolveVatModeDisplayText(ByVal vatMode As String) As String
    Select Case UCase$(Trim$(vatMode))
        Case "INCLUSIVE"
            ResolveVatModeDisplayText = modFwTranslationRuntime.ResolveText("FORM.FRMORDERDETAIL.VAT_MODE_INCLUSIVE", "Inklusive MwSt.")
        Case "NONE"
            ResolveVatModeDisplayText = modFwTranslationRuntime.ResolveText("FORM.FRMORDERDETAIL.VAT_MODE_NONE", "Keine MwSt.")
        Case Else
            ResolveVatModeDisplayText = modFwTranslationRuntime.ResolveText("FORM.FRMORDERDETAIL.VAT_MODE_EXCLUSIVE", "Exklusive MwSt.")
    End Select
End Function

Private Function SaveTemporaryHeader() As Boolean
    On Error GoTo ErrorHandler

    If Not m_sessionInitialized Then
        Exit Function
    End If

    If Me.Dirty Then
        Me.Dirty = False
    End If

    SaveTemporaryHeader = True
    Exit Function

ErrorHandler:
    SaveTemporaryHeader = False
    modErrorHandler.HandleError MODULE_NAME, "SaveTemporaryHeader", Err
End Function

Public Function GetCurrentTemporaryOrderId() As Long
    GetCurrentTemporaryOrderId = m_initializedTmpOrderId
    If GetCurrentTemporaryOrderId <= 0 Then
        GetCurrentTemporaryOrderId = GetLongFieldValue(FIELD_TMP_ORDER_ID)
    End If
End Function

Public Function GetCurrentPersistedOrderId() As Long
    GetCurrentPersistedOrderId = GetLongFieldValue(FIELD_ORDER_ID)
End Function

Public Function GetCurrentOrderVatMode() As String
    GetCurrentOrderVatMode = Trim$(GetTextFieldValue(FIELD_VAT_MODE))
End Function

Public Function GetCurrentOrderVatCode() As String
    GetCurrentOrderVatCode = Trim$(GetTextFieldValue(FIELD_VAT_CODE))
End Function

Public Function GetCurrentOrderVatRate() As Double
    GetCurrentOrderVatRate = GetDoubleFieldValue(FIELD_VAT_RATE)
End Function

Public Sub RefreshCalculatedTotals()
    On Error GoTo ErrorHandler

    Me.Refresh
    UpdateHeaderCaption
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RefreshCalculatedTotals", Err
End Sub

Private Function ValidateHeaderForCommit() As Boolean
    If m_initializedTmpOrderId <= 0 Then
        Exit Function
    End If

    If GetLongFieldValue(FIELD_CUSTOMER_ADDRESS_ID) <= 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ORDER_DETAIL_SAVE_ERROR", "Fehler beim Speichern der Bestellung."), vbExclamation, MODULE_NAME
        Exit Function
    End If

    If GetLongFieldValue(FIELD_INVOICE_ADDRESS_ID) <= 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ORDER_DETAIL_SAVE_ERROR", "Fehler beim Speichern der Bestellung."), vbExclamation, MODULE_NAME
        Exit Function
    End If

    If GetLongFieldValue(FIELD_DELIVERY_ADDRESS_ID) <= 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ORDER_DETAIL_SAVE_ERROR", "Fehler beim Speichern der Bestellung."), vbExclamation, MODULE_NAME
        Exit Function
    End If

    ValidateHeaderForCommit = True
End Function

Private Function LeaveCurrentView() As Boolean
    On Error GoTo ErrorHandler

    If ResolveHostingShellForm() Is Nothing Then
        DoCmd.Close acForm, Me.Name
        LeaveCurrentView = True
        Exit Function
    End If

    LeaveCurrentView = modAppWorkspaceService.GoBack(ResolveHostingShellForm())
    If Not LeaveCurrentView Then
        LeaveCurrentView = modAppWorkspaceService.OpenWorkspaceForm(ResolveHostingShellForm(), DASHBOARD_FORM_NAME)
    End If
    Exit Function

ErrorHandler:
    LeaveCurrentView = False
    modErrorHandler.HandleError MODULE_NAME, "LeaveCurrentView", Err
End Function

Private Sub CleanupTemporarySessionIfNeeded()
    On Error GoTo ErrorHandler

    If m_cleanupCompleted Then
        Exit Sub
    End If

    If Not m_deleteSessionOnClose Then
        Exit Sub
    End If

    If m_initializedTmpOrderId <= 0 Then
        Exit Sub
    End If

    If modOrderRepository.DeleteTemporaryOrder(m_initializedTmpOrderId) Then
        m_cleanupCompleted = True
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CleanupTemporarySessionIfNeeded", Err
End Sub

Private Sub HandleInvalidOrderContext(ByVal messageKey As String, ByVal fallbackText As String)
    MsgBox modFwTranslationRuntime.ResolveText(messageKey, fallbackText), vbExclamation, MODULE_NAME

    If ResolveHostingShellForm() Is Nothing Then
        DoCmd.Close acForm, Me.Name
    Else
        ReturnToPreviousWorkspace "HandleInvalidOrderContext"
    End If
End Sub

Private Sub NavigateHome()
    On Error GoTo ErrorHandler

    If ResolveHostingShellForm() Is Nothing Then
        Exit Sub
    End If

    If Not modAppWorkspaceService.OpenWorkspaceForm(ResolveHostingShellForm(), DASHBOARD_FORM_NAME) Then
        Err.Raise vbObjectError + 5610, MODULE_NAME & ".NavigateHome", "Workspace navigation to dashboard failed."
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

    If ResolveHostingShellForm() Is Nothing Then
        Exit Sub
    End If

    If modAppWorkspaceService.GoBack(ResolveHostingShellForm()) Then
        Exit Sub
    End If

    If Not modAppWorkspaceService.OpenWorkspaceForm(ResolveHostingShellForm(), DASHBOARD_FORM_NAME) Then
        modLoggingHandler.LogWarning MODULE_NAME & "." & callerName, _
            "Workspace history was not available and dashboard fallback could not be opened."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, callerName & ".ReturnToPreviousWorkspace", Err
End Sub

Private Function ResolveHostingShellForm() As Access.Form
    On Error GoTo SafeExit

    If Me.Parent Is Nothing Then
        Exit Function
    End If

    If StrComp(Me.Parent.Name, "frmAppShell", vbTextCompare) = 0 Then
        Set ResolveHostingShellForm = Me.Parent
    End If

SafeExit:
End Function

Private Sub RefreshCommandBar()
    On Error GoTo SafeExit

    If ResolveHostingShellForm() Is Nothing Then
        Exit Sub
    End If

    modAppShell.RefreshWorkspaceCommandBar ResolveHostingShellForm()

SafeExit:
End Sub

Private Function HasUnsavedChanges() As Boolean
    On Error GoTo SafeExit

    HasUnsavedChanges = Me.Dirty

SafeExit:
End Function

Private Function IsExistingOrderSession() As Boolean
    IsExistingOrderSession = (GetLongFieldValue(FIELD_SOURCE_ORDER_ID) > 0)
End Function

Private Function ResolveCurrentDocumentLanguageCode() As String
    Dim languageCode As String

    languageCode = Trim$(GetTextFieldValue(FIELD_LANGUAGE_CODE))
    If LenB(languageCode) = 0 Then
        languageCode = Trim$(modTenantRepository.GetTenantParameter("DEFAULT_LANGUAGE", "en-US"))
    End If

    languageCode = Trim$(modFwTranslationRuntime.NormalizeProjectLanguageCode(languageCode))
    If LenB(languageCode) = 0 Then
        languageCode = "en-US"
    End If

    ResolveCurrentDocumentLanguageCode = languageCode
End Function

Private Function GetLongFieldValue(ByVal fieldName As String) As Long
    On Error GoTo SafeExit

    If HasControl(fieldName) Then
        GetLongFieldValue = modDaoHelper.NzLong(Me.Controls(fieldName).Value, 0)
        Exit Function
    End If

    If modDaoHelper.RecordsetHasField(Me.Recordset, fieldName) Then
        GetLongFieldValue = modDaoHelper.NzLong(Me.Recordset.Fields(fieldName).Value, 0)
    End If

SafeExit:
End Function

Private Function GetTextFieldValue(ByVal fieldName As String) As String
    On Error GoTo SafeExit

    If HasControl(fieldName) Then
        GetTextFieldValue = Trim$(modDaoHelper.NzString(Me.Controls(fieldName).Value, vbNullString))
        Exit Function
    End If

    If modDaoHelper.RecordsetHasField(Me.Recordset, fieldName) Then
        GetTextFieldValue = Trim$(modDaoHelper.NzString(Me.Recordset.Fields(fieldName).Value, vbNullString))
    End If

SafeExit:
End Function

Private Function GetDoubleFieldValue(ByVal fieldName As String) As Double
    On Error GoTo SafeExit

    Dim rawValue As Variant

    If HasControl(fieldName) Then
        rawValue = Me.Controls(fieldName).Value
        If IsNumeric(Nz(rawValue, 0)) Then
            GetDoubleFieldValue = CDbl(Nz(rawValue, 0))
        End If
        Exit Function
    End If

    If modDaoHelper.RecordsetHasField(Me.Recordset, fieldName) Then
        rawValue = Me.Recordset.Fields(fieldName).Value
        If IsNumeric(Nz(rawValue, 0)) Then
            GetDoubleFieldValue = CDbl(Nz(rawValue, 0))
        End If
    End If

SafeExit:
End Function

Private Sub SetControlVisibleIfPresent(ByVal controlName As String, ByVal visibleState As Boolean)
    On Error GoTo SafeExit

    If HasControl(controlName) Then
        Me.Controls(controlName).Visible = visibleState
    End If

SafeExit:
End Sub

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
