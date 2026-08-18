VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmOrderLinesNext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ExportKind: CODE_BEHIND_ONLY
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_sfrmOrderLinesNext
' Purpose   : Local tmp_order_line editor for frmOrderDetailNext.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_sfrmOrderLinesNext"
Private Const TABLE_TMP_ORDER_LINE As String = "tmp_order_line"

Private Const FIELD_TMP_ORDER_ID As String = "tmp_order_id"
Private Const FIELD_ORDER_ID As String = "order_id"
Private Const FIELD_ORDER_LINE_ID As String = "order_line_id"
Private Const FIELD_LINE_NO As String = "line_no"
Private Const FIELD_SORT_ORDER As String = "sort_order"
Private Const FIELD_LINE_TYPE_CODE As String = "line_type_code"
Private Const FIELD_UNIT_CODE As String = "unit_code"
Private Const FIELD_QUANTITY As String = "quantity"
Private Const FIELD_UNIT_PRICE As String = "unit_price"
Private Const FIELD_VAT_CODE As String = "vat_code"
Private Const FIELD_VAT_RATE As String = "vat_rate"
Private Const FIELD_VAT_MODE As String = "vat_mode"
Private Const FIELD_DISCOUNT_TYPE As String = "discount_type"
Private Const FIELD_DISCOUNT_VALUE As String = "discount_value"
Private Const FIELD_SURCHARGE_TYPE As String = "surcharge_type"
Private Const FIELD_SURCHARGE_VALUE As String = "surcharge_value"
Private Const FIELD_LINE_BASE_AMOUNT As String = "line_base_amount"
Private Const FIELD_LINE_DISCOUNT_AMOUNT As String = "line_discount_amount"
Private Const FIELD_LINE_SURCHARGE_AMOUNT As String = "line_surcharge_amount"
Private Const FIELD_LINE_NET_AMOUNT As String = "line_net_amount"
Private Const FIELD_LINE_VAT_AMOUNT As String = "line_vat_amount"
Private Const FIELD_LINE_GROSS_AMOUNT As String = "line_gross_amount"

Private Const ADJUSTMENT_TYPE_NONE As String = "NONE"
Private Const EVENT_PROCEDURE As String = "[Event Procedure]"

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    EnsureFixedRecordSource
    ConfigureEmptyLinesBehavior
    EnsureControlEventBindings
    ConfigureAdjustmentCombo "cboDiscountType", FIELD_DISCOUNT_TYPE
    ConfigureAdjustmentCombo "cboSurchargeType", FIELD_SURCHARGE_TYPE
    ConfigureVatCodeCombo
    EnsureControlTranslationTag "cboVatCode", "FORM.SFRMORDERLINESNEXT.CBOVATCODE"
    EnsureControlTranslationTag "lblVatCode", "FORM.SFRMORDERLINESNEXT.LBLVATCODE"
    modFwTranslationRuntime.ApplyTranslations Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    ConfigureVatCodeCombo
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo ErrorHandler

    EnsureNewLineDefaults Cancel
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_BeforeInsert", Err
End Sub

Private Sub Form_AfterInsert()
    On Error GoTo ErrorHandler

    FocusDescriptionControl
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_AfterInsert", Err
End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)
    On Error GoTo ErrorHandler

    If Status <> acDeleteOK Then
        Exit Sub
    End If

    RecalculateParentOrderTotals
    RefreshParentTotals
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_AfterDelConfirm", Err
End Sub

Private Sub txtQuantity_AfterUpdate()
    LogAfterUpdateEvent "txtQuantity_AfterUpdate"
    RecalculateCurrentLineAndOrder "txtQuantity_AfterUpdate"
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    LogAfterUpdateEvent "txtUnitPrice_AfterUpdate"
    RecalculateCurrentLineAndOrder "txtUnitPrice_AfterUpdate"
End Sub

Private Sub txtDiscountValue_AfterUpdate()
    LogAfterUpdateEvent "txtDiscountValue_AfterUpdate"
    RecalculateCurrentLineAndOrder "txtDiscountValue_AfterUpdate"
End Sub

Private Sub cboDiscountType_AfterUpdate()
    LogAfterUpdateEvent "cboDiscountType_AfterUpdate"
    RecalculateCurrentLineAndOrder "cboDiscountType_AfterUpdate"
End Sub

Private Sub txtSurchargeValue_AfterUpdate()
    LogAfterUpdateEvent "txtSurchargeValue_AfterUpdate"
    RecalculateCurrentLineAndOrder "txtSurchargeValue_AfterUpdate"
End Sub

Private Sub cboSurchargeType_AfterUpdate()
    LogAfterUpdateEvent "cboSurchargeType_AfterUpdate"
    RecalculateCurrentLineAndOrder "cboSurchargeType_AfterUpdate"
End Sub

Private Sub cboVatCode_AfterUpdate()
    On Error GoTo ErrorHandler

    LogAfterUpdateEvent "cboVatCode_AfterUpdate"
    UpdateVatRateFromVatCodeCombo
    RecalculateCurrentLineAndOrder "cboVatCode_AfterUpdate"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboVatCode_AfterUpdate", Err
End Sub

Private Sub EnsureFixedRecordSource()
    On Error GoTo ErrorHandler

    If StrComp(modDaoHelper.NzString(Me.RecordSource), TABLE_TMP_ORDER_LINE, vbTextCompare) <> 0 Then
        Me.RecordSource = TABLE_TMP_ORDER_LINE
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureFixedRecordSource", Err
End Sub

Private Sub ConfigureEmptyLinesBehavior()
    On Error GoTo ErrorHandler

    Me.AllowAdditions = True
    Me.DataEntry = False
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureEmptyLinesBehavior", Err
End Sub

Private Sub EnsureControlEventBindings()
    On Error GoTo ErrorHandler

    EnsureAfterUpdateEventProcedure "txtQuantity"
    EnsureAfterUpdateEventProcedure "txtUnitPrice"
    EnsureAfterUpdateEventProcedure "cboVatCode"
    EnsureAfterUpdateEventProcedure "txtDiscountValue"
    EnsureAfterUpdateEventProcedure "cboDiscountType"
    EnsureAfterUpdateEventProcedure "txtSurchargeValue"
    EnsureAfterUpdateEventProcedure "cboSurchargeType"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureControlEventBindings", Err
End Sub

Private Sub EnsureAfterUpdateEventProcedure(ByVal controlName As String)
    On Error GoTo SafeExit

    If Not HasControl(controlName) Then
        Exit Sub
    End If

    If StrComp(modDaoHelper.NzString(Me.Controls(controlName).OnAfterUpdate), EVENT_PROCEDURE, vbTextCompare) <> 0 Then
        Me.Controls(controlName).OnAfterUpdate = EVENT_PROCEDURE
    End If

SafeExit:
End Sub

Private Sub ConfigureAdjustmentCombo(ByVal controlName As String, ByVal controlSource As String)
    On Error GoTo SafeExit

    If Not HasControl(controlName) Then
        Exit Sub
    End If

    With Me.Controls(controlName)
        .ControlSource = controlSource
        .RowSourceType = "Value List"
        .RowSource = "NONE;PERCENT;AMOUNT"
        .BoundColumn = 1
        .ColumnCount = 1
        .LimitToList = True
    End With

SafeExit:
End Sub

Private Sub ConfigureVatCodeCombo()
    On Error GoTo ErrorHandler

    If Not HasControl("cboVatCode") Then
        Exit Sub
    End If

    With Me.Controls("cboVatCode")
        .ControlSource = FIELD_VAT_CODE
        .RowSourceType = "Value List"
        .RowSource = BuildVatCodeValueList()
        .BoundColumn = 1
        .ColumnCount = 3
        .ColumnWidths = "0cm;5cm;0cm"
        .LimitToList = True
    End With
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureVatCodeCombo", Err
End Sub

Private Function BuildVatCodeValueList() As String
    On Error GoTo ErrorHandler

    Dim frontendDb As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlStatement As String
    Dim vatCode As String
    Dim vatRate As Double
    Dim translationKey As String
    Dim displayText As String
    Dim itemText As String

    Set frontendDb = modDb.GetFrontendDatabase()
    If frontendDb Is Nothing Then
        Exit Function
    End If

    sqlStatement = "SELECT vat_code, translation_key, vat_rate " & _
                   "FROM ref_vat_code " & _
                   "WHERE Nz(is_active, True)=True " & _
                   "ORDER BY Nz(sort_order, 0), Nz(vat_code, '');"
    Set rs = frontendDb.OpenRecordset(sqlStatement, dbOpenSnapshot)

    Do While Not rs.EOF
        vatCode = Trim$(modDaoHelper.NzString(rs.Fields("vat_code").Value))
        translationKey = Trim$(modDaoHelper.NzString(rs.Fields("translation_key").Value))
        vatRate = NzDouble(rs.Fields("vat_rate").Value, 0)
        displayText = ResolveVatCodeDisplayText(translationKey, vatCode, vatRate)
        itemText = vatCode & ";" & displayText & ";" & Replace(Format$(vatRate, "0.0"), ",", ".")

        If LenB(BuildVatCodeValueList) > 0 Then
            BuildVatCodeValueList = BuildVatCodeValueList & ";"
        End If
        BuildVatCodeValueList = BuildVatCodeValueList & itemText
        rs.MoveNext
    Loop

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set frontendDb = Nothing
    Exit Function

ErrorHandler:
    BuildVatCodeValueList = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "BuildVatCodeValueList", Err
    Resume CleanExit
End Function

Private Function ResolveVatCodeDisplayText( _
    ByVal translationKey As String, _
    ByVal vatCode As String, _
    ByVal vatRate As Double) As String

    Dim fallbackText As String
    Dim translatedText As String

    fallbackText = vatCode
    If LenB(translationKey) > 0 Then
        fallbackText = modFwTranslationRuntime.ResolveText(translationKey, vatCode)
    End If

    translatedText = Trim$(fallbackText)
    If LenB(translatedText) = 0 Then
        translatedText = vatCode
    End If

    ResolveVatCodeDisplayText = translatedText & " (" & Replace(Format$(vatRate, "0.0"), ",", ".") & "%)"
End Function

Private Sub UpdateVatRateFromVatCodeCombo()
    On Error GoTo ErrorHandler

    If Not HasControl("cboVatCode") Then
        Exit Sub
    End If

    If LenB(Trim$(modDaoHelper.NzString(Me.Controls("cboVatCode").Value))) = 0 Then
        Exit Sub
    End If

    SetBoundFieldValue FIELD_VAT_RATE, NzDouble(Me.Controls("cboVatCode").Column(2), 0)
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "UpdateVatRateFromVatCodeCombo", Err
End Sub

Private Sub EnsureNewLineDefaults(ByRef Cancel As Integer)
    On Error GoTo ErrorHandler

    Dim tmpOrderId As Long
    Dim nextLineNo As Long
    Dim persistedOrderId As Long

    tmpOrderId = GetParentTmpOrderId()
    If tmpOrderId <= 0 Then
        Cancel = True
        Exit Sub
    End If

    If GetBoundLongValue(FIELD_TMP_ORDER_ID, 0) <= 0 Then
        SetBoundFieldValue FIELD_TMP_ORDER_ID, tmpOrderId
    End If

    persistedOrderId = GetParentPersistedOrderId()
    If GetBoundLongValue(FIELD_ORDER_ID, 0) <= 0 And persistedOrderId > 0 Then
        SetBoundFieldValue FIELD_ORDER_ID, persistedOrderId
    End If

    nextLineNo = GetNextLineNo(tmpOrderId)
    If GetBoundLongValue(FIELD_LINE_NO, 0) <= 0 Then
        SetBoundFieldValue FIELD_LINE_NO, nextLineNo
    End If
    If GetBoundLongValue(FIELD_SORT_ORDER, 0) <= 0 Then
        SetBoundFieldValue FIELD_SORT_ORDER, nextLineNo
    End If

    If LenB(GetBoundTextValue(FIELD_DISCOUNT_TYPE, vbNullString)) = 0 Then
        SetBoundFieldValue FIELD_DISCOUNT_TYPE, ADJUSTMENT_TYPE_NONE
    End If

    If LenB(GetBoundTextValue(FIELD_SURCHARGE_TYPE, vbNullString)) = 0 Then
        SetBoundFieldValue FIELD_SURCHARGE_TYPE, ADJUSTMENT_TYPE_NONE
    End If

    ApplyParentVatDefaultsToCurrentLine
    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureNewLineDefaults", _
        "tmp_order_id=" & CStr(GetBoundLongValue(FIELD_TMP_ORDER_ID, 0)) & "; order_line_id=" & CStr(GetBoundLongValue(FIELD_ORDER_LINE_ID, 0)) & "; line_no=" & CStr(GetBoundLongValue(FIELD_LINE_NO, 0)) & "; vat_code='" & Replace(GetBoundTextValue(FIELD_VAT_CODE, vbNullString), "'", "''") & "'; vat_rate=" & Replace(CStr(NzDouble(GetBoundValue(FIELD_VAT_RATE), 0)), ",", ".") & "."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureNewLineDefaults", Err
End Sub

Private Function GetNextLineNo(ByVal tmpOrderId As Long) As Long
    On Error GoTo ErrorHandler

    Dim frontendDb As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlStatement As String

    GetNextLineNo = 10
    If tmpOrderId <= 0 Then
        Exit Function
    End If

    Set frontendDb = modDb.GetFrontendDatabase()
    If frontendDb Is Nothing Then
        Exit Function
    End If

    sqlStatement = "SELECT MAX([" & FIELD_LINE_NO & "]) AS max_line_no " & _
                   "FROM [" & TABLE_TMP_ORDER_LINE & "] " & _
                   "WHERE [" & FIELD_TMP_ORDER_ID & "]=" & CStr(tmpOrderId) & ";"
    Set rs = frontendDb.OpenRecordset(sqlStatement, dbOpenSnapshot)

    If Not (rs.BOF And rs.EOF) Then
        GetNextLineNo = modDaoHelper.NzLong(rs.Fields("max_line_no").Value, 0) + 10
        If GetNextLineNo <= 0 Then
            GetNextLineNo = 10
        End If
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set frontendDb = Nothing
    Exit Function

ErrorHandler:
    GetNextLineNo = 10
    modErrorHandler.HandleError MODULE_NAME, "GetNextLineNo", Err
    Resume CleanExit
End Function

Private Sub ApplyParentVatDefaultsToCurrentLine()
    On Error GoTo ErrorHandler

    Dim parentVatCode As String
    Dim parentVatRate As Double

    parentVatCode = GetParentVatCode()
    parentVatRate = GetParentVatRate()

    If LenB(GetBoundTextValue(FIELD_VAT_CODE, vbNullString)) = 0 And LenB(parentVatCode) > 0 Then
        SetBoundFieldValue FIELD_VAT_CODE, parentVatCode
    End If

    If NzDouble(GetBoundValue(FIELD_VAT_RATE), 0) = 0 Then
        If parentVatRate <> 0 Or LenB(parentVatCode) > 0 Then
            SetBoundFieldValue FIELD_VAT_RATE, parentVatRate
        End If
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyParentVatDefaultsToCurrentLine", Err
End Sub

Private Sub RecalculateCurrentLineAndOrder(ByVal sourceEventName As String)
    On Error GoTo ErrorHandler

    Dim tmpOrderId As Long
    Dim orderLineId As Long
    Dim lineNo As Long
    Dim lineCalculated As Boolean
    Dim cancel As Integer

    tmpOrderId = GetParentTmpOrderId()
    If tmpOrderId <= 0 Then
        modLoggingHandler.LogInfo MODULE_NAME & ".RecalculateCurrentLineAndOrder", _
            "source_event='" & sourceEventName & "'; skipped=True; reason='missing_tmp_order_id'."
        Exit Sub
    End If

    EnsureNewLineDefaults cancel
    If cancel <> 0 Then
        modLoggingHandler.LogInfo MODULE_NAME & ".RecalculateCurrentLineAndOrder", _
            "source_event='" & sourceEventName & "'; skipped=True; reason='defaults_cancelled'."
        Exit Sub
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".RecalculateCurrentLineAndOrder", _
        "source_event='" & sourceEventName & "'; tmp_order_id=" & CStr(tmpOrderId) & "; order_line_id=" & CStr(GetCurrentOrderLineId()) & "; line_no=" & CStr(GetBoundLongValue(FIELD_LINE_NO, 0)) & "; dirty=" & CStr(Me.Dirty) & "; quantity=" & Replace(CStr(NzDouble(GetBoundValue(FIELD_QUANTITY), 0)), ",", ".") & "; unit_price=" & Replace(CStr(NzDouble(GetBoundValue(FIELD_UNIT_PRICE), 0)), ",", ".") & "; vat_code='" & Replace(GetBoundTextValue(FIELD_VAT_CODE, vbNullString), "'", "''") & "'; vat_rate=" & Replace(CStr(NzDouble(GetBoundValue(FIELD_VAT_RATE), 0)), ",", ".") & "."

    If Me.Dirty Then
        Me.Dirty = False
    End If

    orderLineId = GetCurrentOrderLineId()
    lineNo = GetBoundLongValue(FIELD_LINE_NO, 0)

    If orderLineId > 0 Then
        lineCalculated = modOrderCalculationService.CalculateTemporaryOrderLineAmounts(orderLineId)
    ElseIf lineNo > 0 Then
        lineCalculated = modOrderCalculationService.CalculateTemporaryOrderLineAmountsByOrderAndLineNo(tmpOrderId, lineNo)
    End If

    If Not lineCalculated Then
        modLoggingHandler.LogInfo MODULE_NAME & ".RecalculateCurrentLineAndOrder", _
            "source_event='" & sourceEventName & "'; line_calculated=False; tmp_order_id=" & CStr(tmpOrderId) & "; order_line_id=" & CStr(orderLineId) & "; line_no=" & CStr(lineNo) & "."
        Exit Sub
    End If

    Me.Refresh

    modLoggingHandler.LogInfo MODULE_NAME & ".RecalculateCurrentLineAndOrder", _
        "source_event='" & sourceEventName & "'; line_calculated=True; line_base_amount=" & Replace(CStr(NzDouble(GetBoundValue(FIELD_LINE_BASE_AMOUNT), 0)), ",", ".") & "; line_net_amount=" & Replace(CStr(NzDouble(GetBoundValue(FIELD_LINE_NET_AMOUNT), 0)), ",", ".") & "; line_vat_amount=" & Replace(CStr(NzDouble(GetBoundValue(FIELD_LINE_VAT_AMOUNT), 0)), ",", ".") & "; line_gross_amount=" & Replace(CStr(NzDouble(GetBoundValue(FIELD_LINE_GROSS_AMOUNT), 0)), ",", ".") & "."

    If Not RecalculateParentOrderTotals() Then
        modLoggingHandler.LogInfo MODULE_NAME & ".RecalculateCurrentLineAndOrder", _
            "source_event='" & sourceEventName & "'; parent_totals_calculated=False; tmp_order_id=" & CStr(tmpOrderId) & "."
        Exit Sub
    End If

    RefreshParentTotals
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, sourceEventName, Err
End Sub

Private Function RecalculateParentOrderTotals() As Boolean
    Dim tmpOrderId As Long

    tmpOrderId = GetParentTmpOrderId()
    If tmpOrderId <= 0 Then
        Exit Function
    End If

    RecalculateParentOrderTotals = modOrderCalculationService.CalculateTemporaryOrderTotals(tmpOrderId)
    modLoggingHandler.LogInfo MODULE_NAME & ".RecalculateParentOrderTotals", _
        "tmp_order_id=" & CStr(tmpOrderId) & "; recalculated=" & CStr(RecalculateParentOrderTotals) & "."
End Function

Private Sub RefreshParentTotals()
    On Error GoTo SafeExit

    If ResolveParentOrderForm() Is Nothing Then
        Exit Sub
    End If

    CallByName ResolveParentOrderForm(), "RefreshCalculatedTotals", VbMethod

SafeExit:
End Sub

Private Function ResolveParentOrderForm() As Access.Form
    On Error GoTo SafeExit

    If Me.Parent Is Nothing Then
        Exit Function
    End If

    If StrComp(Me.Parent.Name, "frmOrderDetailNext", vbTextCompare) = 0 Then
        Set ResolveParentOrderForm = Me.Parent
    End If

SafeExit:
End Function

Private Function GetParentTmpOrderId() As Long
    On Error GoTo SafeExit

    If ResolveParentOrderForm() Is Nothing Then
        Exit Function
    End If

    GetParentTmpOrderId = modDaoHelper.NzLong(CallByName(ResolveParentOrderForm(), "GetCurrentTemporaryOrderId", VbMethod), 0)

SafeExit:
End Function

Private Function GetParentPersistedOrderId() As Long
    On Error GoTo SafeExit

    If ResolveParentOrderForm() Is Nothing Then
        Exit Function
    End If

    GetParentPersistedOrderId = modDaoHelper.NzLong(CallByName(ResolveParentOrderForm(), "GetCurrentPersistedOrderId", VbMethod), 0)

SafeExit:
End Function

Private Function GetParentVatMode() As String
    On Error GoTo SafeExit

    If ResolveParentOrderForm() Is Nothing Then
        Exit Function
    End If

    GetParentVatMode = Trim$(modDaoHelper.NzString(CallByName(ResolveParentOrderForm(), "GetCurrentOrderVatMode", VbMethod), modVatHandler.GetVatMode()))

SafeExit:
    If LenB(GetParentVatMode) = 0 Then
        GetParentVatMode = modVatHandler.GetVatMode()
    End If
End Function

Private Function GetParentVatCode() As String
    On Error GoTo SafeExit

    If ResolveParentOrderForm() Is Nothing Then
        Exit Function
    End If

    GetParentVatCode = Trim$(modDaoHelper.NzString(CallByName(ResolveParentOrderForm(), "GetCurrentOrderVatCode", VbMethod), vbNullString))

SafeExit:
End Function

Private Function GetParentVatRate() As Double
    On Error GoTo SafeExit

    If ResolveParentOrderForm() Is Nothing Then
        Exit Function
    End If

    GetParentVatRate = NzDouble(CallByName(ResolveParentOrderForm(), "GetCurrentOrderVatRate", VbMethod), 0)

SafeExit:
End Function

Private Function GetCurrentOrderLineId() As Long
    GetCurrentOrderLineId = GetBoundLongValue(FIELD_ORDER_LINE_ID, 0)
End Function

Private Function GetBoundLongValue(ByVal fieldName As String, ByVal defaultValue As Long) As Long
    GetBoundLongValue = modDaoHelper.NzLong(GetBoundValue(fieldName), defaultValue)
End Function

Private Function GetBoundTextValue(ByVal fieldName As String, ByVal defaultValue As String) As String
    GetBoundTextValue = Trim$(modDaoHelper.NzString(GetBoundValue(fieldName), defaultValue))
End Function

Private Function GetBoundValue(ByVal fieldName As String) As Variant
    On Error GoTo SafeExit

    If HasControl(fieldName) Then
        GetBoundValue = Me.Controls(fieldName).Value
        Exit Function
    End If

    Select Case UCase$(fieldName)
        Case UCase$(FIELD_LINE_NO)
            If HasControl("txtLineNo") Then
                GetBoundValue = Me.Controls("txtLineNo").Value
                Exit Function
            End If
        Case UCase$(FIELD_QUANTITY)
            If HasControl("txtQuantity") Then
                GetBoundValue = Me.Controls("txtQuantity").Value
                Exit Function
            End If
        Case UCase$(FIELD_UNIT_PRICE)
            If HasControl("txtUnitPrice") Then
                GetBoundValue = Me.Controls("txtUnitPrice").Value
                Exit Function
            End If
        Case UCase$(FIELD_VAT_CODE)
            If HasControl("cboVatCode") Then
                GetBoundValue = Me.Controls("cboVatCode").Value
                Exit Function
            End If
        Case UCase$(FIELD_VAT_RATE)
            If HasControl("txtVatRate") Then
                GetBoundValue = Me.Controls("txtVatRate").Value
                Exit Function
            End If
        Case UCase$(FIELD_DISCOUNT_TYPE)
            If HasControl("cboDiscountType") Then
                GetBoundValue = Me.Controls("cboDiscountType").Value
                Exit Function
            End If
        Case UCase$(FIELD_SURCHARGE_TYPE)
            If HasControl("cboSurchargeType") Then
                GetBoundValue = Me.Controls("cboSurchargeType").Value
                Exit Function
            End If
        Case UCase$(FIELD_DISCOUNT_VALUE)
            If HasControl("txtDiscountValue") Then
                GetBoundValue = Me.Controls("txtDiscountValue").Value
                Exit Function
            End If
        Case UCase$(FIELD_SURCHARGE_VALUE)
            If HasControl("txtSurchargeValue") Then
                GetBoundValue = Me.Controls("txtSurchargeValue").Value
                Exit Function
            End If
        Case UCase$(FIELD_LINE_NET_AMOUNT)
            If HasControl("txtLineNetAmount") Then
                GetBoundValue = Me.Controls("txtLineNetAmount").Value
                Exit Function
            End If
        Case UCase$(FIELD_LINE_VAT_AMOUNT)
            If HasControl("txtLineVatAmount") Then
                GetBoundValue = Me.Controls("txtLineVatAmount").Value
                Exit Function
            End If
        Case UCase$(FIELD_LINE_GROSS_AMOUNT)
            If HasControl("txtLineGrossAmount") Then
                GetBoundValue = Me.Controls("txtLineGrossAmount").Value
                Exit Function
            End If
        Case UCase$(FIELD_LINE_BASE_AMOUNT)
            If HasControl("txtLineBaseAmount") Then
                GetBoundValue = Me.Controls("txtLineBaseAmount").Value
                Exit Function
            End If
        Case UCase$(FIELD_LINE_DISCOUNT_AMOUNT)
            If HasControl("txtLineDiscountAmount") Then
                GetBoundValue = Me.Controls("txtLineDiscountAmount").Value
                Exit Function
            End If
        Case UCase$(FIELD_LINE_SURCHARGE_AMOUNT)
            If HasControl("txtLineSurchargeAmount") Then
                GetBoundValue = Me.Controls("txtLineSurchargeAmount").Value
                Exit Function
            End If
        Case UCase$(FIELD_ORDER_LINE_ID)
            If HasControl("txtOrderLineId") Then
                GetBoundValue = Me.Controls("txtOrderLineId").Value
                Exit Function
            End If
    End Select

    If modDaoHelper.RecordsetHasField(Me.Recordset, fieldName) Then
        GetBoundValue = Me.Recordset.Fields(fieldName).Value
    End If

SafeExit:
End Function

Private Sub SetBoundFieldValue(ByVal fieldName As String, ByVal fieldValue As Variant)
    On Error GoTo SafeExit

    If HasControl(fieldName) Then
        Me.Controls(fieldName).Value = fieldValue
        Exit Sub
    End If

    Select Case UCase$(fieldName)
        Case UCase$(FIELD_LINE_NO)
            If HasControl("txtLineNo") Then
                Me.Controls("txtLineNo").Value = fieldValue
                Exit Sub
            End If
        Case UCase$(FIELD_QUANTITY)
            If HasControl("txtQuantity") Then
                Me.Controls("txtQuantity").Value = fieldValue
                Exit Sub
            End If
        Case UCase$(FIELD_UNIT_PRICE)
            If HasControl("txtUnitPrice") Then
                Me.Controls("txtUnitPrice").Value = fieldValue
                Exit Sub
            End If
        Case UCase$(FIELD_VAT_CODE)
            If HasControl("cboVatCode") Then
                Me.Controls("cboVatCode").Value = fieldValue
                Exit Sub
            End If
        Case UCase$(FIELD_VAT_RATE)
            If HasControl("txtVatRate") Then
                Me.Controls("txtVatRate").Value = fieldValue
                Exit Sub
            End If
        Case UCase$(FIELD_DISCOUNT_TYPE)
            If HasControl("cboDiscountType") Then
                Me.Controls("cboDiscountType").Value = fieldValue
                Exit Sub
            End If
        Case UCase$(FIELD_SURCHARGE_TYPE)
            If HasControl("cboSurchargeType") Then
                Me.Controls("cboSurchargeType").Value = fieldValue
                Exit Sub
            End If
        Case UCase$(FIELD_DISCOUNT_VALUE)
            If HasControl("txtDiscountValue") Then
                Me.Controls("txtDiscountValue").Value = fieldValue
                Exit Sub
            End If
        Case UCase$(FIELD_SURCHARGE_VALUE)
            If HasControl("txtSurchargeValue") Then
                Me.Controls("txtSurchargeValue").Value = fieldValue
                Exit Sub
            End If
        Case UCase$(FIELD_LINE_NET_AMOUNT)
            If HasControl("txtLineNetAmount") Then
                Me.Controls("txtLineNetAmount").Value = fieldValue
                Exit Sub
            End If
        Case UCase$(FIELD_LINE_VAT_AMOUNT)
            If HasControl("txtLineVatAmount") Then
                Me.Controls("txtLineVatAmount").Value = fieldValue
                Exit Sub
            End If
        Case UCase$(FIELD_LINE_GROSS_AMOUNT)
            If HasControl("txtLineGrossAmount") Then
                Me.Controls("txtLineGrossAmount").Value = fieldValue
                Exit Sub
            End If
        Case UCase$(FIELD_LINE_BASE_AMOUNT)
            If HasControl("txtLineBaseAmount") Then
                Me.Controls("txtLineBaseAmount").Value = fieldValue
                Exit Sub
            End If
        Case UCase$(FIELD_LINE_DISCOUNT_AMOUNT)
            If HasControl("txtLineDiscountAmount") Then
                Me.Controls("txtLineDiscountAmount").Value = fieldValue
                Exit Sub
            End If
        Case UCase$(FIELD_LINE_SURCHARGE_AMOUNT)
            If HasControl("txtLineSurchargeAmount") Then
                Me.Controls("txtLineSurchargeAmount").Value = fieldValue
                Exit Sub
            End If
        Case UCase$(FIELD_ORDER_LINE_ID)
            If HasControl("txtOrderLineId") Then
                Me.Controls("txtOrderLineId").Value = fieldValue
                Exit Sub
            End If
    End Select

    If modDaoHelper.RecordsetHasField(Me.Recordset, fieldName) Then
        Me.Recordset.Fields(fieldName).Value = fieldValue
    End If

SafeExit:
End Sub

Private Sub FocusDescriptionControl()
    On Error GoTo SafeExit

    If HasControl("txtDescriptionText") Then
        Me.Controls("txtDescriptionText").SetFocus
    End If

SafeExit:
End Sub

Private Sub EnsureControlTranslationTag(ByVal controlName As String, ByVal translationKey As String)
    On Error GoTo SafeExit

    If Not HasControl(controlName) Then
        Exit Sub
    End If

    If LenB(Trim$(modDaoHelper.NzString(Me.Controls(controlName).Tag))) = 0 Then
        Me.Controls(controlName).Tag = "TR:" & translationKey
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

Private Function NzDouble(ByVal valueVariant As Variant, Optional ByVal defaultValue As Double = 0) As Double
    Dim rawValue As String

    rawValue = modDaoHelper.NzString(valueVariant, CStr(defaultValue))
    If IsNumeric(rawValue) Then
        NzDouble = CDbl(rawValue)
    Else
        NzDouble = defaultValue
    End If
End Function

Private Sub LogAfterUpdateEvent(ByVal eventName As String)
    On Error GoTo SafeExit

    modLoggingHandler.LogInfo MODULE_NAME & ".AfterUpdate", _
        "event='" & eventName & "'; tmp_order_id=" & CStr(GetParentTmpOrderId()) & "; order_line_id=" & CStr(GetCurrentOrderLineId()) & "; line_no=" & CStr(GetBoundLongValue(FIELD_LINE_NO, 0)) & "; dirty=" & CStr(Me.Dirty) & "."

SafeExit:
End Sub
