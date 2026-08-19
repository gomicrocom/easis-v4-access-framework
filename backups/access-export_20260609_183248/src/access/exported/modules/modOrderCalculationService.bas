Attribute VB_Name = "modOrderCalculationService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modOrderCalculationService
' Purpose   : Calculates Phase 1 sales-order line and header amounts.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modOrderCalculationService"

Private Const TABLE_ORD_ORDER As String = "ord_order"
Private Const TABLE_ORD_ORDER_LINE As String = "ord_order_line"

Private Const FIELD_ORDER_ID As String = "order_id"
Private Const FIELD_ORDER_LINE_ID As String = "order_line_id"
Private Const FIELD_QUANTITY As String = "quantity"
Private Const FIELD_UNIT_PRICE As String = "unit_price"
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

Private Const FIELD_HEADER_DISCOUNT_TYPE As String = "header_discount_type"
Private Const FIELD_HEADER_DISCOUNT_VALUE As String = "header_discount_value"
Private Const FIELD_HEADER_SURCHARGE_TYPE As String = "header_surcharge_type"
Private Const FIELD_HEADER_SURCHARGE_VALUE As String = "header_surcharge_value"
Private Const FIELD_SUBTOTAL_NET_AMOUNT As String = "subtotal_net_amount"
Private Const FIELD_HEADER_DISCOUNT_AMOUNT As String = "header_discount_amount"
Private Const FIELD_HEADER_SURCHARGE_AMOUNT As String = "header_surcharge_amount"
Private Const FIELD_NET_AMOUNT As String = "net_amount"
Private Const FIELD_VAT_AMOUNT As String = "vat_amount"
Private Const FIELD_GROSS_AMOUNT As String = "gross_amount"

Private Const ADJUSTMENT_TYPE_NONE As String = "NONE"

Public Function EnsureOrderCalculationSchema() As Boolean
    On Error GoTo ErrorHandler

    EnsureOrderCalculationSchema = modBasicModuleSchema.EnsureOrderPhase1Schema()
    Exit Function

ErrorHandler:
    EnsureOrderCalculationSchema = False
    modErrorHandler.HandleError MODULE_NAME, "EnsureOrderCalculationSchema", Err
End Function

Public Function CalculateOrderLineAmounts(ByVal OrderLineId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsLine As DAO.Recordset
    Dim SqlText As String
    Dim OrderId As Long
    Dim quantity As Double
    Dim UnitPrice As Currency
    Dim vatRate As Double
    Dim VatMode As String
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

    CalculateOrderLineAmounts = False

    If OrderLineId <= 0 Then
        Exit Function
    End If

    If Not EnsureOrderCalculationSchema() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Exit Function
    End If

    SqlText = "SELECT * FROM [" & TABLE_ORD_ORDER_LINE & "] WHERE [" & FIELD_ORDER_LINE_ID & "]=" & CStr(OrderLineId) & ";"
    Set rsLine = db.OpenRecordset(SqlText, dbOpenDynaset)

    If rsLine.BOF And rsLine.EOF Then
        GoTo CleanExit
    End If

    OrderId = GetRecordsetLongValue(rsLine, FIELD_ORDER_ID, 0)
    quantity = GetRecordsetDoubleValue(rsLine, FIELD_QUANTITY, 0)
    UnitPrice = GetRecordsetCurrencyValue(rsLine, FIELD_UNIT_PRICE, 0)
    vatRate = GetRecordsetDoubleValue(rsLine, FIELD_VAT_RATE, 0)
    VatMode = ResolveOrderVatMode(OrderId, modVatHandler.GetVatMode())
    discountType = GetRecordsetStringValue(rsLine, FIELD_DISCOUNT_TYPE, ADJUSTMENT_TYPE_NONE)
    discountValue = GetRecordsetCurrencyValue(rsLine, FIELD_DISCOUNT_VALUE, 0)
    surchargeType = GetRecordsetStringValue(rsLine, FIELD_SURCHARGE_TYPE, ADJUSTMENT_TYPE_NONE)
    surchargeValue = GetRecordsetCurrencyValue(rsLine, FIELD_SURCHARGE_VALUE, 0)

    baseAmount = RoundCurrency(CCur(quantity * CDbl(UnitPrice)))
    discountAmount = modDocumentCalculationService.CalculateAdjustmentAmount(baseAmount, discountType, discountValue)
    If discountAmount > baseAmount Then
        discountAmount = baseAmount
    End If

    discountedBaseAmount = RoundCurrency(baseAmount - discountAmount)
    surchargeAmount = modDocumentCalculationService.CalculateAdjustmentAmount(discountedBaseAmount, surchargeType, surchargeValue)

    lineNetAmount = RoundCurrency(discountedBaseAmount + surchargeAmount)
    If lineNetAmount < 0 Then
        lineNetAmount = 0
    End If

    lineVatAmount = modVatHandler.CalculateVatAmount(lineNetAmount, vatRate, VatMode)
    lineGrossAmount = ResolveGrossAmount(lineNetAmount, lineVatAmount, vatRate, VatMode)

    rsLine.Edit
    SetRecordsetValue rsLine, FIELD_LINE_BASE_AMOUNT, baseAmount
    SetRecordsetValue rsLine, FIELD_LINE_DISCOUNT_AMOUNT, discountAmount
    SetRecordsetValue rsLine, FIELD_LINE_SURCHARGE_AMOUNT, surchargeAmount
    SetRecordsetValue rsLine, FIELD_LINE_NET_AMOUNT, lineNetAmount
    SetRecordsetValue rsLine, FIELD_LINE_VAT_AMOUNT, lineVatAmount
    SetRecordsetValue rsLine, FIELD_LINE_GROSS_AMOUNT, lineGrossAmount
    rsLine.Update

    CalculateOrderLineAmounts = True

CleanExit:
    On Error Resume Next
    If Not rsLine Is Nothing Then rsLine.Close
    Set rsLine = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    CalculateOrderLineAmounts = False
    modErrorHandler.HandleError MODULE_NAME, "CalculateOrderLineAmounts", Err
    Resume CleanExit
End Function

Public Function CalculateOrderTotals(ByVal OrderId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsOrder As DAO.Recordset
    Dim rsLines As DAO.Recordset
    Dim SqlText As String
    Dim subtotalNetAmount As Currency
    Dim totalVatAmount As Currency
    Dim totalGrossAmount As Currency
    Dim headerDiscountType As String
    Dim headerDiscountValue As Currency
    Dim headerSurchargeType As String
    Dim headerSurchargeValue As Currency

    CalculateOrderTotals = False

    If OrderId <= 0 Then
        Exit Function
    End If

    If Not EnsureOrderCalculationSchema() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Exit Function
    End If

    SqlText = "SELECT * FROM [" & TABLE_ORD_ORDER & "] WHERE [" & FIELD_ORDER_ID & "]=" & CStr(OrderId) & ";"
    Set rsOrder = db.OpenRecordset(SqlText, dbOpenDynaset)
    If rsOrder.BOF And rsOrder.EOF Then
        GoTo CleanExit
    End If

    SqlText = "SELECT * FROM [" & TABLE_ORD_ORDER_LINE & "] WHERE [" & FIELD_ORDER_ID & "]=" & CStr(OrderId) & ";"
    Set rsLines = db.OpenRecordset(SqlText, dbOpenSnapshot)

    If Not (rsLines.BOF And rsLines.EOF) Then
        rsLines.MoveFirst
        Do Until rsLines.EOF
            subtotalNetAmount = subtotalNetAmount + GetRecordsetCurrencyValue(rsLines, FIELD_LINE_NET_AMOUNT, 0)
            totalVatAmount = totalVatAmount + GetRecordsetCurrencyValue(rsLines, FIELD_LINE_VAT_AMOUNT, 0)
            totalGrossAmount = totalGrossAmount + GetRecordsetCurrencyValue(rsLines, FIELD_LINE_GROSS_AMOUNT, 0)
            rsLines.MoveNext
        Loop
    End If

    headerDiscountType = GetRecordsetStringValue(rsOrder, FIELD_HEADER_DISCOUNT_TYPE, ADJUSTMENT_TYPE_NONE)
    headerDiscountValue = GetRecordsetCurrencyValue(rsOrder, FIELD_HEADER_DISCOUNT_VALUE, 0)
    headerSurchargeType = GetRecordsetStringValue(rsOrder, FIELD_HEADER_SURCHARGE_TYPE, ADJUSTMENT_TYPE_NONE)
    headerSurchargeValue = GetRecordsetCurrencyValue(rsOrder, FIELD_HEADER_SURCHARGE_VALUE, 0)

    If Not HeaderAdjustmentsAreInactive(headerDiscountType, headerDiscountValue, headerSurchargeType, headerSurchargeValue) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CalculateOrderTotals", _
            "Header adjustments are stored but not applied yet for OrderId=" & CStr(OrderId) & "."
    End If

    rsOrder.Edit
    SetRecordsetValue rsOrder, FIELD_SUBTOTAL_NET_AMOUNT, subtotalNetAmount
    SetRecordsetValue rsOrder, FIELD_HEADER_DISCOUNT_AMOUNT, 0
    SetRecordsetValue rsOrder, FIELD_HEADER_SURCHARGE_AMOUNT, 0
    SetRecordsetValue rsOrder, FIELD_NET_AMOUNT, subtotalNetAmount
    SetRecordsetValue rsOrder, FIELD_VAT_AMOUNT, totalVatAmount
    SetRecordsetValue rsOrder, FIELD_GROSS_AMOUNT, totalGrossAmount
    rsOrder.Update

    CalculateOrderTotals = True

CleanExit:
    On Error Resume Next
    If Not rsLines Is Nothing Then rsLines.Close
    If Not rsOrder Is Nothing Then rsOrder.Close
    Set rsLines = Nothing
    Set rsOrder = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    CalculateOrderTotals = False
    modErrorHandler.HandleError MODULE_NAME, "CalculateOrderTotals", Err
    Resume CleanExit
End Function

Public Function RecalculateOrder(ByVal OrderId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsLines As DAO.Recordset
    Dim SqlText As String
    Dim OrderLineId As Long

    RecalculateOrder = False

    If OrderId <= 0 Then
        Exit Function
    End If

    If Not EnsureOrderCalculationSchema() Then
        Exit Function
    End If

    If Not modOrderRepository.OrderExists(OrderId) Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Exit Function
    End If

    SqlText = "SELECT [" & FIELD_ORDER_LINE_ID & "] FROM [" & TABLE_ORD_ORDER_LINE & "] WHERE [" & FIELD_ORDER_ID & "]=" & CStr(OrderId) & ";"
    Set rsLines = db.OpenRecordset(SqlText, dbOpenSnapshot)

    If Not (rsLines.BOF And rsLines.EOF) Then
        rsLines.MoveFirst
        Do Until rsLines.EOF
            OrderLineId = GetRecordsetLongValue(rsLines, FIELD_ORDER_LINE_ID, 0)
            If OrderLineId > 0 Then
                If Not CalculateOrderLineAmounts(OrderLineId) Then
                    GoTo CleanExit
                End If
            End If
            rsLines.MoveNext
        Loop
    End If

    RecalculateOrder = CalculateOrderTotals(OrderId)

CleanExit:
    On Error Resume Next
    If Not rsLines Is Nothing Then rsLines.Close
    Set rsLines = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    RecalculateOrder = False
    modErrorHandler.HandleError MODULE_NAME, "RecalculateOrder", Err
    Resume CleanExit
End Function

Private Function HeaderAdjustmentsAreInactive( _
    ByVal headerDiscountType As String, _
    ByVal headerDiscountValue As Currency, _
    ByVal headerSurchargeType As String, _
    ByVal headerSurchargeValue As Currency) As Boolean

    HeaderAdjustmentsAreInactive = (UCase$(Trim$(headerDiscountType)) = ADJUSTMENT_TYPE_NONE Or headerDiscountValue = 0) _
        And (UCase$(Trim$(headerSurchargeType)) = ADJUSTMENT_TYPE_NONE Or headerSurchargeValue = 0)
End Function

Private Function ResolveGrossAmount(ByVal NetAmount As Currency, ByVal VatAmount As Currency, ByVal vatRate As Double, ByVal VatMode As String) As Currency
    Select Case modVatHandler.NormalizeVatMode(VatMode)
        Case "EXCLUSIVE"
            ResolveGrossAmount = modVatHandler.CalculateGrossFromNet(NetAmount, vatRate)
        Case "INCLUSIVE", "NONE"
            ResolveGrossAmount = RoundCurrency(NetAmount)
        Case Else
            ResolveGrossAmount = RoundCurrency(NetAmount + VatAmount)
    End Select
End Function

Private Function ResolveOrderVatMode(ByVal OrderId As Long, ByVal defaultValue As String) As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsOrder As DAO.Recordset
    Dim SqlText As String

    ResolveOrderVatMode = defaultValue

    If OrderId <= 0 Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Exit Function
    End If

    SqlText = "SELECT [" & FIELD_VAT_MODE & "] FROM [" & TABLE_ORD_ORDER & "] WHERE [" & FIELD_ORDER_ID & "]=" & CStr(OrderId) & ";"
    Set rsOrder = db.OpenRecordset(SqlText, dbOpenSnapshot)

    If Not (rsOrder.BOF And rsOrder.EOF) Then
        ResolveOrderVatMode = GetRecordsetStringValue(rsOrder, FIELD_VAT_MODE, defaultValue)
    End If

CleanExit:
    On Error Resume Next
    If Not rsOrder Is Nothing Then rsOrder.Close
    Set rsOrder = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    ResolveOrderVatMode = defaultValue
    modErrorHandler.HandleError MODULE_NAME, "ResolveOrderVatMode", Err
    Resume CleanExit
End Function

Private Function GetRecordsetStringValue(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal defaultValue As String) As String
    If modDaoHelper.RecordsetHasField(rs, fieldName) Then
        GetRecordsetStringValue = Trim$(modDaoHelper.NzString(rs.Fields(fieldName).Value, defaultValue))
    Else
        GetRecordsetStringValue = defaultValue
    End If
End Function

Private Function GetRecordsetCurrencyValue(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal defaultValue As Currency) As Currency
    Dim rawValue As String

    If modDaoHelper.RecordsetHasField(rs, fieldName) Then
        rawValue = modDaoHelper.NzString(rs.Fields(fieldName).Value, CStr(defaultValue))
        If IsNumeric(rawValue) Then
            GetRecordsetCurrencyValue = CCur(rawValue)
        Else
            GetRecordsetCurrencyValue = defaultValue
        End If
    Else
        GetRecordsetCurrencyValue = defaultValue
    End If
End Function

Private Function GetRecordsetDoubleValue(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal defaultValue As Double) As Double
    Dim rawValue As String

    If modDaoHelper.RecordsetHasField(rs, fieldName) Then
        rawValue = modDaoHelper.NzString(rs.Fields(fieldName).Value, CStr(defaultValue))
        If IsNumeric(rawValue) Then
            GetRecordsetDoubleValue = CDbl(rawValue)
        Else
            GetRecordsetDoubleValue = defaultValue
        End If
    Else
        GetRecordsetDoubleValue = defaultValue
    End If
End Function

Private Function GetRecordsetLongValue(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal defaultValue As Long) As Long
    If modDaoHelper.RecordsetHasField(rs, fieldName) Then
        GetRecordsetLongValue = modDaoHelper.NzLong(rs.Fields(fieldName).Value, defaultValue)
    Else
        GetRecordsetLongValue = defaultValue
    End If
End Function

Private Sub SetRecordsetValue(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal fieldValue As Variant)
    If modDaoHelper.RecordsetHasField(rs, fieldName) Then
        rs.Fields(fieldName).Value = fieldValue
    End If
End Sub

Private Function RoundCurrency(ByVal Amount As Currency) As Currency
    RoundCurrency = CCur(Round(CDbl(Amount), 2))
End Function
