VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmOrderLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ExportKind: CODE_BEHIND_ONLY
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_sfrmOrderLines
' Purpose   : Editable sales-order line subform with lightweight recalculation.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_sfrmOrderLines"
Private Const FIELD_ORDER_ID As String = "order_id"
Private Const FIELD_LINE_NO As String = "line_no"
Private Const FIELD_QUANTITY As String = "quantity"
Private Const FIELD_UNIT_PRICE As String = "unit_price"
Private Const FIELD_DISCOUNT_TYPE As String = "discount_type"
Private Const FIELD_DISCOUNT_VALUE As String = "discount_value"
Private Const FIELD_SURCHARGE_TYPE As String = "surcharge_type"
Private Const FIELD_SURCHARGE_VALUE As String = "surcharge_value"
Private Const FIELD_VAT_CODE As String = "vat_code"
Private Const FIELD_VAT_RATE As String = "vat_rate"
Private Const FIELD_LINE_BASE_AMOUNT As String = "line_base_amount"
Private Const FIELD_LINE_DISCOUNT_AMOUNT As String = "line_discount_amount"
Private Const FIELD_LINE_SURCHARGE_AMOUNT As String = "line_surcharge_amount"
Private Const FIELD_LINE_NET_AMOUNT As String = "line_net_amount"
Private Const FIELD_LINE_VAT_AMOUNT As String = "line_vat_amount"
Private Const FIELD_LINE_GROSS_AMOUNT As String = "line_gross_amount"
Private Const FIELD_VAT_MODE As String = "vat_mode"
Private Const ADJUSTMENT_TYPE_NONE As String = "NONE"

Private m_pendingDeleteOrderId As Long
Private m_lineTableName As String
Private m_parentKeyFieldName As String
Private m_lineKeyFieldName As String

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "parent_name='" & ResolveParentNameForLogging() & _
        "'; parent_order_id=" & CStr(GetParentOrderIdForLogging()) & "."

    InitializeRecordSource
    ConfigureEmptyLinesBehavior
    ConfigureAdjustmentCombo "cboDiscountType", FIELD_DISCOUNT_TYPE
    ConfigureAdjustmentCombo "cboSurchargeType", FIELD_SURCHARGE_TYPE
    ConfigureVatCodeCombo
    EnsureControlTranslationTag "cboVatCode", "FORM.SFRMORDERLINES.CBOVATCODE"
    EnsureControlTranslationTag "lblVatCode", "FORM.SFRMORDERLINES.LBLVATCODE"
    modFwTranslationRuntime.ApplyTranslations Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
End Sub

Public Sub ConfigureForOrderContext( _
    ByVal lineTableName As String, _
    ByVal parentKeyFieldName As String, _
    ByVal lineKeyFieldName As String)
    On Error GoTo ErrorHandler

    m_lineTableName = Trim$(lineTableName)
    m_parentKeyFieldName = Trim$(parentKeyFieldName)
    m_lineKeyFieldName = Trim$(lineKeyFieldName)
    InitializeRecordSource
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureForOrderContext", Err
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    m_pendingDeleteOrderId = GetCurrentOrderId()
    ConfigureVatCodeCombo
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub cboVatCode_AfterUpdate()
    On Error GoTo ErrorHandler

    UpdateVatRateFromVatCodeCombo
    CalcLineTotals "cboVatCode"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboVatCode_AfterUpdate", Err
End Sub

Private Sub txtQuantity_AfterUpdate()
    On Error GoTo ErrorHandler

    CalcLineTotals "txtQuantity"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtQuantity_AfterUpdate", Err
End Sub

Private Sub txtUnitPrice_AfterUpdate()
    On Error GoTo ErrorHandler

    CalcLineTotals "txtUnitPrice"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtUnitPrice_AfterUpdate", Err
End Sub

Private Sub txtDiscountValue_AfterUpdate()
    On Error GoTo ErrorHandler

    CalcLineTotals "txtDiscountValue"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtDiscountValue_AfterUpdate", Err
End Sub

Private Sub cboDiscountType_AfterUpdate()
    On Error GoTo ErrorHandler

    CalcLineTotals "cboDiscountType"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboDiscountType_AfterUpdate", Err
End Sub

Private Sub txtSurchargeValue_AfterUpdate()
    On Error GoTo ErrorHandler

    CalcLineTotals "txtSurchargeValue"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtSurchargeValue_AfterUpdate", Err
End Sub

Private Sub cboSurchargeType_AfterUpdate()
    On Error GoTo ErrorHandler

    CalcLineTotals "cboSurchargeType"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboSurchargeType_AfterUpdate", Err
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

Private Sub Form_BeforeUpdate(Cancel As Integer)
    ' Local line calculation happens directly in field AfterUpdate handlers.
End Sub

Private Sub Form_AfterUpdate()
    ' Header totals are intentionally not recalculated on line field edits.
End Sub

Private Sub Form_Delete(Cancel As Integer)
    On Error GoTo ErrorHandler

    m_pendingDeleteOrderId = GetCurrentOrderId()
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Delete", Err
End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)
    ' Header totals are intentionally not recalculated automatically here.
End Sub

Private Sub InitializeRecordSource()
    On Error GoTo ErrorHandler

    If LenB(m_lineTableName) = 0 Then
        m_lineTableName = "ord_order_line"
    End If
    If LenB(m_parentKeyFieldName) = 0 Then
        m_parentKeyFieldName = FIELD_ORDER_ID
    End If
    If LenB(m_lineKeyFieldName) = 0 Then
        m_lineKeyFieldName = "order_line_id"
    End If

    If StrComp(Trim$(modDaoHelper.NzString(Me.RecordSource)), m_lineTableName, vbTextCompare) <> 0 Then
        Me.RecordSource = m_lineTableName
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "InitializeRecordSource", Err
End Sub

Private Sub ConfigureEmptyLinesBehavior()
    On Error GoTo ErrorHandler

    ' A sales order without lines is valid. Keep the subform bound and editable
    ' so the user can add the first line directly into an empty detail section.
    Me.AllowAdditions = True
    Me.DataEntry = False
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureEmptyLinesBehavior", Err
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

Private Sub EnsureNewLineDefaults(ByRef Cancel As Integer)
    On Error GoTo ErrorHandler

    Dim parentOrderId As Long

    parentOrderId = GetParentOrderId()
    If parentOrderId <= 0 Then
        Cancel = True
        modLoggingHandler.LogWarning MODULE_NAME & ".EnsureNewLineDefaults", _
            "Order line insert was blocked because the parent order is not persisted yet."
        Exit Sub
    End If

    If GetBoundLongValue(m_parentKeyFieldName, 0) <= 0 Then
        SetBoundFieldValue m_parentKeyFieldName, parentOrderId
    End If

    If GetBoundLongValue(FIELD_LINE_NO, 0) <= 0 Then
        SetBoundFieldValue FIELD_LINE_NO, GetNextLineNo(parentOrderId)
    End If

    ApplyOrderVatDefaultsToCurrentLine
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureNewLineDefaults", Err
End Sub

Private Function GetNextLineNo(ByVal OrderId As Long) As Long
    On Error GoTo ErrorHandler

    Dim nextValue As Variant

    If OrderId <= 0 Then
        GetNextLineNo = 10
        Exit Function
    End If

    nextValue = DMax(FIELD_LINE_NO, m_lineTableName, "[" & m_parentKeyFieldName & "]=" & CStr(OrderId))
    GetNextLineNo = modDaoHelper.NzLong(nextValue, 0) + 10
    Exit Function

ErrorHandler:
    GetNextLineNo = 10
    modErrorHandler.HandleError MODULE_NAME, "GetNextLineNo", Err
End Function

Private Sub CalcLineTotals(ByVal source_control_name As String)
    On Error GoTo ErrorHandler

    Dim quantity As Double
    Dim unitPrice As Currency
    Dim vatRate As Double
    Dim vatMode As String
    Dim discountType As String
    Dim discountValue As Currency
    Dim surchargeType As String
    Dim surchargeValue As Currency
    Dim baseAmount As Currency
    Dim discountAmount As Currency
    Dim surchargeAmount As Currency
    Dim lineNetAmount As Currency
    Dim lineVatAmount As Currency
    Dim lineGrossAmount As Currency
    Dim discountedBaseAmount As Currency
    Dim calculationBaseAmount As Currency
    Dim effectiveVatRate As Double

    quantity = NzDouble(GetControlValue("txtQuantity", FIELD_QUANTITY), 0)
    unitPrice = NzCurrency(GetControlValue("txtUnitPrice", FIELD_UNIT_PRICE), 0)
    vatRate = NzDouble(GetControlValue("txtVatRate", FIELD_VAT_RATE), 0)
    vatMode = ResolveCurrentVatMode()
    discountType = Trim$(modDaoHelper.NzString(GetControlValue("cboDiscountType", FIELD_DISCOUNT_TYPE), ADJUSTMENT_TYPE_NONE))
    surchargeType = Trim$(modDaoHelper.NzString(GetControlValue("cboSurchargeType", FIELD_SURCHARGE_TYPE), ADJUSTMENT_TYPE_NONE))
    discountValue = NzCurrency(GetControlValue("txtDiscountValue", FIELD_DISCOUNT_VALUE), 0)
    surchargeValue = NzCurrency(GetControlValue("txtSurchargeValue", FIELD_SURCHARGE_VALUE), 0)

    effectiveVatRate = vatRate
    If StrComp(modVatHandler.NormalizeVatMode(vatMode), "NONE", vbTextCompare) = 0 Then
        effectiveVatRate = 0
    End If

    baseAmount = RoundCurrency(CCur(quantity * CDbl(unitPrice)))
    Select Case UCase$(discountType)
        Case "", ADJUSTMENT_TYPE_NONE
            discountAmount = 0
        Case "PERCENT"
            discountAmount = RoundCurrency(CCur(CDbl(baseAmount) * CDbl(discountValue) / 100#))
        Case "AMOUNT"
            discountAmount = RoundCurrency(discountValue)
        Case Else
            discountAmount = 0
    End Select
    If discountAmount > baseAmount Then discountAmount = baseAmount
    discountedBaseAmount = RoundCurrency(baseAmount - discountAmount)
    Select Case UCase$(surchargeType)
        Case "", ADJUSTMENT_TYPE_NONE
            surchargeAmount = 0
        Case "PERCENT"
            surchargeAmount = RoundCurrency(CCur(CDbl(discountedBaseAmount) * CDbl(surchargeValue) / 100#))
        Case "AMOUNT"
            surchargeAmount = RoundCurrency(surchargeValue)
        Case Else
            surchargeAmount = 0
    End Select
    calculationBaseAmount = RoundCurrency(discountedBaseAmount + surchargeAmount)

    Select Case modVatHandler.NormalizeVatMode(vatMode)
        Case "INCLUSIVE"
            lineGrossAmount = calculationBaseAmount
            lineNetAmount = modVatHandler.CalculateNetFromGross(lineGrossAmount, effectiveVatRate)
            lineVatAmount = RoundCurrency(lineGrossAmount - lineNetAmount)
        Case "NONE"
            lineNetAmount = calculationBaseAmount
            lineVatAmount = 0
            lineGrossAmount = lineNetAmount
        Case Else
            lineNetAmount = calculationBaseAmount
            lineVatAmount = modVatHandler.CalculateVatAmount(lineNetAmount, effectiveVatRate, vatMode)
            lineGrossAmount = ResolveGrossAmount(lineNetAmount, lineVatAmount, effectiveVatRate, vatMode)
    End Select

    If lineNetAmount < 0 Then lineNetAmount = 0
    If lineGrossAmount < 0 Then lineGrossAmount = 0

    SetControlValue "txtLineBaseAmount", FIELD_LINE_BASE_AMOUNT, baseAmount
    SetControlValue "txtLineDiscountAmount", FIELD_LINE_DISCOUNT_AMOUNT, discountAmount
    SetControlValue "txtLineSurchargeAmount", FIELD_LINE_SURCHARGE_AMOUNT, surchargeAmount
    SetControlValue "txtLineNetAmount", FIELD_LINE_NET_AMOUNT, lineNetAmount
    SetControlValue "txtLineVatAmount", FIELD_LINE_VAT_AMOUNT, lineVatAmount
    SetControlValue "txtLineGrossAmount", FIELD_LINE_GROSS_AMOUNT, lineGrossAmount
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CalcLineTotals", Err
End Sub

Private Sub ConfigureVatCodeCombo()
    On Error GoTo SafeExit

    If Not HasControl("cboVatCode") Then
        Exit Sub
    End If

    With Me.Controls("cboVatCode")
        .ControlSource = FIELD_VAT_CODE
        .RowSourceType = "Table/Query"
        .RowSource = modOrderRepository.GetVatCodeRowSourceSql()
        .BoundColumn = 1
        .ColumnCount = 3
        .ColumnWidths = "0cm;5cm;0cm"
        .LimitToList = True
    End With

SafeExit:
End Sub

Private Sub UpdateVatRateFromVatCodeCombo()
    On Error GoTo ErrorHandler

    If Not HasControl("cboVatCode") Then
        Exit Sub
    End If

    If LenB(Trim$(modDaoHelper.NzString(Me.Controls("cboVatCode").Value))) = 0 Then
        Exit Sub
    End If

    SetControlValue "txtVatRate", FIELD_VAT_RATE, NzDouble(Me.Controls("cboVatCode").Column(2), 0)

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "UpdateVatRateFromVatCodeCombo", Err
End Sub

Private Sub ApplyOrderVatDefaultsToCurrentLine()
    On Error GoTo ErrorHandler

    Dim parentForm As Access.Form
    Dim parentVatCode As String
    Dim parentVatRate As Double

    Set parentForm = ResolveParentOrderForm()
    If parentForm Is Nothing Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ApplyOrderVatDefaultsToCurrentLine", _
            "Parent order form was not available during line insert."
        Exit Sub
    End If

    parentVatCode = GetParentOrderVatCode(parentForm)
    parentVatRate = GetParentOrderVatRate(parentForm)

    If LenB(Trim$(GetBoundTextValue(FIELD_VAT_CODE, vbNullString))) = 0 Then
        If LenB(parentVatCode) > 0 Then
            SetBoundFieldValue FIELD_VAT_CODE, parentVatCode
        Else
            modLoggingHandler.LogWarning MODULE_NAME & ".ApplyOrderVatDefaultsToCurrentLine", _
                "Parent vat_code was not available for current line."
        End If
    End If

    If NzDouble(GetBoundValue(FIELD_VAT_RATE), 0) = 0 Then
        If parentVatRate <> 0 Or LenB(parentVatCode) > 0 Then
            SetBoundFieldValue FIELD_VAT_RATE, parentVatRate
        Else
            modLoggingHandler.LogWarning MODULE_NAME & ".ApplyOrderVatDefaultsToCurrentLine", _
                "Parent vat_rate was not available for current line."
        End If
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyOrderVatDefaultsToCurrentLine", Err
End Sub

Private Function GetCurrentOrderId() As Long
    On Error GoTo SafeExit

    GetCurrentOrderId = GetBoundLongValue(m_parentKeyFieldName, 0)

SafeExit:
End Function

Private Function GetParentOrderId() As Long
    On Error GoTo SafeExit

    Dim parentForm As Access.Form

    Set parentForm = ResolveParentOrderForm()
    If parentForm Is Nothing Then
        Exit Function
    End If

    GetParentOrderId = modDaoHelper.NzLong(CallByName(parentForm, "GetCurrentOrderHeaderId", VbMethod), 0)

SafeExit:
End Function

Private Function ResolveParentOrderForm() As Access.Form
    On Error GoTo SafeExit

    Dim parentForm As Access.Form

    Set parentForm = Me.Parent
    If parentForm Is Nothing Then
        Exit Function
    End If

    If StrComp(parentForm.Name, "frmOrderDetail", vbTextCompare) = 0 Then
        Set ResolveParentOrderForm = parentForm
    End If

SafeExit:
End Function

Private Function GetParentOrderVatCode(ByVal parentForm As Access.Form) As String
    On Error GoTo ErrorHandler

    If parentForm Is Nothing Then
        Exit Function
    End If

    If HasFormControl(parentForm, FIELD_VAT_CODE) Then
        GetParentOrderVatCode = modDaoHelper.NzString(parentForm.Controls(FIELD_VAT_CODE).Value)
    ElseIf HasFormControl(parentForm, "txtVatCode") Then
        GetParentOrderVatCode = modDaoHelper.NzString(parentForm.Controls("txtVatCode").Value)
    End If
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetParentOrderVatCode", Err
End Function

Private Function GetParentOrderVatRate(ByVal parentForm As Access.Form) As Double
    On Error GoTo ErrorHandler

    If parentForm Is Nothing Then
        Exit Function
    End If

    If HasFormControl(parentForm, FIELD_VAT_RATE) Then
        GetParentOrderVatRate = NzDouble(parentForm.Controls(FIELD_VAT_RATE).Value, 0)
    ElseIf HasFormControl(parentForm, "txtVatRate") Then
        GetParentOrderVatRate = NzDouble(parentForm.Controls("txtVatRate").Value, 0)
    End If
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetParentOrderVatRate", Err
End Function

Private Function GetControlValue(ByVal controlName As String, ByVal fallbackFieldName As String) As Variant
    On Error GoTo ErrorHandler

    If HasControl(controlName) Then
        GetControlValue = Me.Controls(controlName).Value
        Exit Function
    End If

    GetControlValue = GetBoundValue(fallbackFieldName)
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetControlValue", Err
End Function

Private Sub SetControlValue(ByVal controlName As String, ByVal fallbackFieldName As String, ByVal fieldValue As Variant)
    On Error GoTo ErrorHandler

    If HasControl(controlName) Then
        Me.Controls(controlName).Value = fieldValue
    Else
        SetBoundFieldValue fallbackFieldName, fieldValue
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SetControlValue", Err
End Sub

Private Function ResolveCurrentVatMode() As String
    On Error GoTo SafeExit

    Dim parentForm As Access.Form

    Set parentForm = ResolveParentOrderForm()
    If parentForm Is Nothing Then
        ResolveCurrentVatMode = modVatHandler.GetVatMode()
        Exit Function
    End If

    If HasFormControl(parentForm, FIELD_VAT_MODE) Then
        ResolveCurrentVatMode = modDaoHelper.NzString(parentForm.Controls(FIELD_VAT_MODE).Value, modVatHandler.GetVatMode())
    ElseIf HasFormControl(parentForm, "cboVatMode") Then
        ResolveCurrentVatMode = modDaoHelper.NzString(parentForm.Controls("cboVatMode").Value, modVatHandler.GetVatMode())
    Else
        ResolveCurrentVatMode = modVatHandler.GetVatMode()
    End If

SafeExit:
    If LenB(ResolveCurrentVatMode) = 0 Then
        ResolveCurrentVatMode = modVatHandler.GetVatMode()
    End If
End Function

Public Sub SetOrderEntryEnabled(ByVal isEnabled As Boolean)
    On Error GoTo ErrorHandler

    Me.AllowAdditions = isEnabled
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SetOrderEntryEnabled", Err
End Sub

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
    End Select

    If modDaoHelper.RecordsetHasField(Me.Recordset, fieldName) Then
        GetBoundValue = Me.Recordset.Fields(fieldName).Value
    End If

SafeExit:
End Function

Private Function GetBoundLongValue(ByVal fieldName As String, ByVal defaultValue As Long) As Long
    GetBoundLongValue = modDaoHelper.NzLong(GetBoundValue(fieldName), defaultValue)
End Function

Private Function GetBoundTextValue(ByVal fieldName As String, ByVal defaultValue As String) As String
    GetBoundTextValue = Trim$(modDaoHelper.NzString(GetBoundValue(fieldName), defaultValue))
End Function

Private Sub SetBoundFieldValue(ByVal fieldName As String, ByVal fieldValue As Variant)
    On Error GoTo SafeExit

    If HasControl(fieldName) Then
        Me.Controls(fieldName).Value = fieldValue
    End If

    Select Case UCase$(fieldName)
        Case UCase$(FIELD_LINE_NO)
            If HasControl("txtLineNo") Then
                Me.Controls("txtLineNo").Value = fieldValue
            End If
        Case UCase$(FIELD_VAT_CODE)
            If HasControl("cboVatCode") Then
                Me.Controls("cboVatCode").Value = fieldValue
            End If
        Case UCase$(FIELD_VAT_RATE)
            If HasControl("txtVatRate") Then
                Me.Controls("txtVatRate").Value = fieldValue
            End If
    End Select

    If modDaoHelper.RecordsetHasField(Me.Recordset, fieldName) Then
        Me.Recordset.Fields(fieldName).Value = fieldValue
    End If

SafeExit:
End Sub

Private Sub FocusDescriptionControl()
    On Error GoTo ErrorHandler

    If HasControl("txtDescriptionText") Then
        Me.Controls("txtDescriptionText").SetFocus
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "FocusDescriptionControl", Err
End Sub

Private Function NzDouble(ByVal Value As Variant, Optional ByVal defaultValue As Double = 0) As Double
    If IsNull(Value) Or IsEmpty(Value) Or LenB(Trim$(CStr(Value))) = 0 Then
        NzDouble = defaultValue
    ElseIf IsNumeric(Value) Then
        NzDouble = CDbl(Value)
    Else
        NzDouble = defaultValue
    End If
End Function

Private Function NzCurrency(ByVal Value As Variant, Optional ByVal defaultValue As Currency = 0) As Currency
    If IsNull(Value) Or IsEmpty(Value) Or LenB(Trim$(CStr(Value))) = 0 Then
        NzCurrency = defaultValue
    ElseIf IsNumeric(Value) Then
        NzCurrency = CCur(Value)
    Else
        NzCurrency = defaultValue
    End If
End Function

Private Function RoundCurrency(ByVal Amount As Currency) As Currency
    RoundCurrency = CCur(Round(CDbl(Amount), 2))
End Function

Private Function ResolveGrossAmount(ByVal netAmount As Currency, ByVal vatAmount As Currency, ByVal vatRate As Double, ByVal vatMode As String) As Currency
    Select Case modVatHandler.NormalizeVatMode(vatMode)
        Case "EXCLUSIVE"
            ResolveGrossAmount = modVatHandler.CalculateGrossFromNet(netAmount, vatRate)
        Case "INCLUSIVE", "NONE"
            ResolveGrossAmount = RoundCurrency(netAmount)
        Case Else
            ResolveGrossAmount = RoundCurrency(netAmount + vatAmount)
    End Select
End Function

Private Sub EnsureControlTranslationTag(ByVal ControlName As String, ByVal translationKey As String)
    On Error GoTo ErrorHandler

    If HasControl(ControlName) Then
        Me.Controls(ControlName).Tag = modFwTranslationRuntime.SetTranslationKeyInTag( _
            modDaoHelper.NzString(Me.Controls(ControlName).Tag), _
            translationKey)
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureControlTranslationTag", Err
End Sub

Private Function ResolveParentNameForLogging() As String
    On Error GoTo SafeExit

    If Not Me.Parent Is Nothing Then
        ResolveParentNameForLogging = Me.Parent.Name
    End If

SafeExit:
    If LenB(ResolveParentNameForLogging) = 0 Then
        ResolveParentNameForLogging = "<none>"
    End If
End Function

Private Function GetParentOrderIdForLogging() As Long
    On Error GoTo SafeExit

    GetParentOrderIdForLogging = GetParentOrderId()

SafeExit:
End Function

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

Private Function HasFormControl(ByVal formInstance As Access.Form, ByVal controlName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentControl As Control

    For Each currentControl In formInstance.Controls
        If StrComp(currentControl.Name, controlName, vbTextCompare) = 0 Then
            HasFormControl = True
            Exit Function
        End If
    Next currentControl

SafeExit:
End Function
