Option Compare Database
Option Explicit

'===============================================================================
' Module    : modDocumentCalculationService
' Purpose   : Calculates document position and header amounts including discounts
'             and surcharges and ensures required calculation fields exist.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modDocumentCalculationService"

Private Const TABLE_DOC_DOCUMENT As String = "doc_document"
Private Const TABLE_DOC_DOCUMENT_POSITION As String = "doc_document_position"

Private Const FIELD_DOCUMENT_ID As String = "document_id"
Private Const FIELD_DOCUMENT_POSITION_ID As String = "document_position_id"
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
Private Const FIELD_NET_AMOUNT As String = "net_amount"
Private Const FIELD_VAT_AMOUNT As String = "vat_amount"
Private Const FIELD_GROSS_AMOUNT As String = "gross_amount"
Private Const FIELD_LINE_TOTAL_NET As String = "line_total_net"
Private Const FIELD_LINE_TOTAL_VAT As String = "line_total_vat"
Private Const FIELD_LINE_TOTAL_GROSS As String = "line_total_gross"

Private Const FIELD_HEADER_DISCOUNT_TYPE As String = "header_discount_type"
Private Const FIELD_HEADER_DISCOUNT_VALUE As String = "header_discount_value"
Private Const FIELD_HEADER_SURCHARGE_TYPE As String = "header_surcharge_type"
Private Const FIELD_HEADER_SURCHARGE_VALUE As String = "header_surcharge_value"
Private Const FIELD_SUBTOTAL_NET_AMOUNT As String = "subtotal_net_amount"
Private Const FIELD_HEADER_DISCOUNT_AMOUNT As String = "header_discount_amount"
Private Const FIELD_HEADER_SURCHARGE_AMOUNT As String = "header_surcharge_amount"
Private Const FIELD_TOTAL_NET As String = "total_net"
Private Const FIELD_TOTAL_VAT As String = "total_vat"
Private Const FIELD_TOTAL_GROSS As String = "total_gross"

Private Const ADJUSTMENT_TYPE_NONE As String = "NONE"
Private Const ADJUSTMENT_TYPE_PERCENT As String = "PERCENT"
Private Const ADJUSTMENT_TYPE_AMOUNT As String = "AMOUNT"

Private Const TEXT_FIELD_SIZE_TYPE As Long = 20

Public Function EnsureDocumentCalculationSchema() As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database

    EnsureDocumentCalculationSchema = False

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".EnsureDocumentCalculationSchema", _
            "Document calculation schema update skipped because backend configuration is not valid."
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Exit Function
    End If

    If Not TableExists(db, TABLE_DOC_DOCUMENT_POSITION) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".EnsureDocumentCalculationSchema", _
            "Table '" & TABLE_DOC_DOCUMENT_POSITION & "' is not available."
        Exit Function
    End If

    If Not TableExists(db, TABLE_DOC_DOCUMENT) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".EnsureDocumentCalculationSchema", _
            "Table '" & TABLE_DOC_DOCUMENT & "' is not available."
        Exit Function
    End If

    If Not EnsureTextField(db, TABLE_DOC_DOCUMENT_POSITION, FIELD_DISCOUNT_TYPE, TEXT_FIELD_SIZE_TYPE, ADJUSTMENT_TYPE_NONE) Then GoTo CleanExit
    If Not EnsureCurrencyField(db, TABLE_DOC_DOCUMENT_POSITION, FIELD_DISCOUNT_VALUE, 0) Then GoTo CleanExit
    If Not EnsureTextField(db, TABLE_DOC_DOCUMENT_POSITION, FIELD_SURCHARGE_TYPE, TEXT_FIELD_SIZE_TYPE, ADJUSTMENT_TYPE_NONE) Then GoTo CleanExit
    If Not EnsureCurrencyField(db, TABLE_DOC_DOCUMENT_POSITION, FIELD_SURCHARGE_VALUE, 0) Then GoTo CleanExit
    If Not EnsureCurrencyField(db, TABLE_DOC_DOCUMENT_POSITION, FIELD_LINE_BASE_AMOUNT, 0) Then GoTo CleanExit
    If Not EnsureCurrencyField(db, TABLE_DOC_DOCUMENT_POSITION, FIELD_LINE_DISCOUNT_AMOUNT, 0) Then GoTo CleanExit
    If Not EnsureCurrencyField(db, TABLE_DOC_DOCUMENT_POSITION, FIELD_LINE_SURCHARGE_AMOUNT, 0) Then GoTo CleanExit

    If Not EnsureTextField(db, TABLE_DOC_DOCUMENT, FIELD_HEADER_DISCOUNT_TYPE, TEXT_FIELD_SIZE_TYPE, ADJUSTMENT_TYPE_NONE) Then GoTo CleanExit
    If Not EnsureCurrencyField(db, TABLE_DOC_DOCUMENT, FIELD_HEADER_DISCOUNT_VALUE, 0) Then GoTo CleanExit
    If Not EnsureTextField(db, TABLE_DOC_DOCUMENT, FIELD_HEADER_SURCHARGE_TYPE, TEXT_FIELD_SIZE_TYPE, ADJUSTMENT_TYPE_NONE) Then GoTo CleanExit
    If Not EnsureCurrencyField(db, TABLE_DOC_DOCUMENT, FIELD_HEADER_SURCHARGE_VALUE, 0) Then GoTo CleanExit
    If Not EnsureCurrencyField(db, TABLE_DOC_DOCUMENT, FIELD_SUBTOTAL_NET_AMOUNT, 0) Then GoTo CleanExit
    If Not EnsureCurrencyField(db, TABLE_DOC_DOCUMENT, FIELD_HEADER_DISCOUNT_AMOUNT, 0) Then GoTo CleanExit
    If Not EnsureCurrencyField(db, TABLE_DOC_DOCUMENT, FIELD_HEADER_SURCHARGE_AMOUNT, 0) Then GoTo CleanExit

    EnsureDocumentCalculationSchema = True

CleanExit:
    Set db = Nothing
    Exit Function

ErrorHandler:
    EnsureDocumentCalculationSchema = False
    modErrorHandler.HandleError MODULE_NAME, "EnsureDocumentCalculationSchema", Err
    Resume CleanExit
End Function

Public Function CalculateAdjustmentAmount( _
    ByVal BaseAmount As Currency, _
    ByVal AdjustmentType As String, _
    ByVal AdjustmentValue As Currency _
) As Currency
    On Error GoTo ErrorHandler

    Dim normalizedType As String

    normalizedType = NormalizeAdjustmentType(AdjustmentType)

    If AdjustmentValue < 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CalculateAdjustmentAmount", _
            "Negative adjustment value '" & CStr(AdjustmentValue) & "' is not allowed. Falling back to 0."
        Exit Function
    End If

    Select Case normalizedType
        Case ADJUSTMENT_TYPE_NONE
            CalculateAdjustmentAmount = 0

        Case ADJUSTMENT_TYPE_PERCENT
            CalculateAdjustmentAmount = RoundCurrency(CCur(CDbl(BaseAmount) * (CDbl(AdjustmentValue) / 100#)))

        Case ADJUSTMENT_TYPE_AMOUNT
            CalculateAdjustmentAmount = RoundCurrency(AdjustmentValue)

        Case Else
            modLoggingHandler.LogWarning MODULE_NAME & ".CalculateAdjustmentAmount", _
                "Unknown adjustment type '" & Trim$(AdjustmentType) & "'. Falling back to 0."
            CalculateAdjustmentAmount = 0
    End Select
    Exit Function

ErrorHandler:
    CalculateAdjustmentAmount = 0
    modErrorHandler.HandleError MODULE_NAME, "CalculateAdjustmentAmount", Err
End Function

Public Function CalculatePositionAmounts(ByVal DocumentPositionId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsPosition As DAO.Recordset
    Dim sqlText As String
    Dim documentId As Long
    Dim quantity As Double
    Dim unitPrice As Currency
    Dim vatRate As Double
    Dim vatMode As String
    Dim baseAmount As Currency
    Dim discountAmount As Currency
    Dim surchargeAmount As Currency
    Dim netAmount As Currency
    Dim vatAmount As Currency
    Dim grossAmount As Currency
    Dim discountType As String
    Dim surchargeType As String
    Dim discountValue As Currency
    Dim surchargeValue As Currency
    Dim discountedBaseAmount As Currency

    CalculatePositionAmounts = False

    If DocumentPositionId <= 0 Then
        Exit Function
    End If

    If Not EnsureDocumentCalculationSchema() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Exit Function
    End If

    sqlText = "SELECT * FROM [" & TABLE_DOC_DOCUMENT_POSITION & "] WHERE [" & FIELD_DOCUMENT_POSITION_ID & "]=" & CStr(DocumentPositionId) & ";"
    Set rsPosition = db.OpenRecordset(sqlText, dbOpenDynaset)

    If rsPosition.BOF And rsPosition.EOF Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CalculatePositionAmounts", _
            "Position calculation skipped because DocumentPositionId=" & CStr(DocumentPositionId) & " does not exist."
        GoTo CleanExit
    End If

    quantity = GetRecordsetDoubleValue(rsPosition, FIELD_QUANTITY, 0)
    unitPrice = GetRecordsetCurrencyValue(rsPosition, FIELD_UNIT_PRICE, 0)
    vatRate = GetRecordsetDoubleValue(rsPosition, FIELD_VAT_RATE, 0)
    documentId = GetRecordsetLongValue(rsPosition, FIELD_DOCUMENT_ID, 0)
    vatMode = ResolveDocumentVatMode(documentId, modVatHandler.GetVatMode())

    discountType = GetRecordsetStringValue(rsPosition, FIELD_DISCOUNT_TYPE, ADJUSTMENT_TYPE_NONE)
    discountValue = GetRecordsetCurrencyValue(rsPosition, FIELD_DISCOUNT_VALUE, 0)
    surchargeType = GetRecordsetStringValue(rsPosition, FIELD_SURCHARGE_TYPE, ADJUSTMENT_TYPE_NONE)
    surchargeValue = GetRecordsetCurrencyValue(rsPosition, FIELD_SURCHARGE_VALUE, 0)

    baseAmount = RoundCurrency(CCur(quantity * CDbl(unitPrice)))
    discountAmount = CalculateAdjustmentAmount(baseAmount, discountType, discountValue)
    If discountAmount > baseAmount Then
        discountAmount = baseAmount
    End If

    discountedBaseAmount = RoundCurrency(baseAmount - discountAmount)
    surchargeAmount = CalculateAdjustmentAmount(discountedBaseAmount, surchargeType, surchargeValue)

    netAmount = RoundCurrency(discountedBaseAmount + surchargeAmount)
    If netAmount < 0 Then
        netAmount = 0
    End If

    vatAmount = modVatHandler.CalculateVatAmount(netAmount, vatRate, vatMode)
    grossAmount = ResolveGrossAmount(netAmount, vatAmount, vatRate, vatMode)

    rsPosition.Edit
    SetRecordsetValue rsPosition, FIELD_LINE_BASE_AMOUNT, baseAmount
    SetRecordsetValue rsPosition, FIELD_LINE_DISCOUNT_AMOUNT, discountAmount
    SetRecordsetValue rsPosition, FIELD_LINE_SURCHARGE_AMOUNT, surchargeAmount
    SetRecordsetValue rsPosition, FIELD_NET_AMOUNT, netAmount
    SetRecordsetValue rsPosition, FIELD_VAT_AMOUNT, vatAmount
    SetRecordsetValue rsPosition, FIELD_GROSS_AMOUNT, grossAmount
    SetRecordsetValue rsPosition, FIELD_LINE_TOTAL_NET, netAmount
    SetRecordsetValue rsPosition, FIELD_LINE_TOTAL_VAT, vatAmount
    SetRecordsetValue rsPosition, FIELD_LINE_TOTAL_GROSS, grossAmount
    rsPosition.Update

    CalculatePositionAmounts = True

CleanExit:
    On Error Resume Next
    If Not rsPosition Is Nothing Then rsPosition.Close
    Set rsPosition = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    CalculatePositionAmounts = False
    modErrorHandler.HandleError MODULE_NAME, "CalculatePositionAmounts", Err
    Resume CleanExit
End Function

Public Function CalculateDocumentTotals(ByVal DocumentId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsDocument As DAO.Recordset
    Dim rsPositions As DAO.Recordset
    Dim sqlText As String
    Dim subtotalNetAmount As Currency
    Dim vatAmount As Currency
    Dim grossPositionAmount As Currency
    Dim headerDiscountAmount As Currency
    Dim headerSurchargeAmount As Currency
    Dim netAmount As Currency
    Dim grossAmount As Currency
    Dim headerDiscountType As String
    Dim headerSurchargeType As String
    Dim headerDiscountValue As Currency
    Dim headerSurchargeValue As Currency

    CalculateDocumentTotals = False

    If DocumentId <= 0 Then
        Exit Function
    End If

    If Not EnsureDocumentCalculationSchema() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Exit Function
    End If

    sqlText = "SELECT * FROM [" & TABLE_DOC_DOCUMENT & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsDocument = db.OpenRecordset(sqlText, dbOpenDynaset)

    If rsDocument.BOF And rsDocument.EOF Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CalculateDocumentTotals", _
            "Document total calculation skipped because DocumentId=" & CStr(DocumentId) & " does not exist."
        GoTo CleanExit
    End If

    sqlText = "SELECT * FROM [" & TABLE_DOC_DOCUMENT_POSITION & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsPositions = db.OpenRecordset(sqlText, dbOpenSnapshot)

    If Not (rsPositions.BOF And rsPositions.EOF) Then
        rsPositions.MoveFirst
        Do Until rsPositions.EOF
            subtotalNetAmount = subtotalNetAmount + GetPositionNetAmount(rsPositions)
            vatAmount = vatAmount + GetPositionVatAmount(rsPositions)
            grossPositionAmount = grossPositionAmount + GetPositionGrossAmount(rsPositions)
            rsPositions.MoveNext
        Loop
    End If

    headerDiscountType = GetRecordsetStringValue(rsDocument, FIELD_HEADER_DISCOUNT_TYPE, ADJUSTMENT_TYPE_NONE)
    headerDiscountValue = GetRecordsetCurrencyValue(rsDocument, FIELD_HEADER_DISCOUNT_VALUE, 0)
    headerSurchargeType = GetRecordsetStringValue(rsDocument, FIELD_HEADER_SURCHARGE_TYPE, ADJUSTMENT_TYPE_NONE)
    headerSurchargeValue = GetRecordsetCurrencyValue(rsDocument, FIELD_HEADER_SURCHARGE_VALUE, 0)

    If headerDiscountValue > 0 Or headerSurchargeValue > 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CalculateDocumentTotals", _
            "Header adjustments are stored but not applied yet for DocumentId=" & CStr(DocumentId) & "."
    End If

    headerDiscountAmount = 0
    headerSurchargeAmount = 0
    netAmount = subtotalNetAmount
    grossAmount = grossPositionAmount

    rsDocument.Edit
    SetRecordsetValue rsDocument, FIELD_SUBTOTAL_NET_AMOUNT, subtotalNetAmount
    SetRecordsetValue rsDocument, FIELD_HEADER_DISCOUNT_AMOUNT, headerDiscountAmount
    SetRecordsetValue rsDocument, FIELD_HEADER_SURCHARGE_AMOUNT, headerSurchargeAmount
    SetRecordsetValue rsDocument, FIELD_NET_AMOUNT, netAmount
    SetRecordsetValue rsDocument, FIELD_VAT_AMOUNT, vatAmount
    SetRecordsetValue rsDocument, FIELD_GROSS_AMOUNT, grossAmount
    SetRecordsetValue rsDocument, FIELD_TOTAL_NET, netAmount
    SetRecordsetValue rsDocument, FIELD_TOTAL_VAT, vatAmount
    SetRecordsetValue rsDocument, FIELD_TOTAL_GROSS, grossAmount
    rsDocument.Update

    CalculateDocumentTotals = True

CleanExit:
    On Error Resume Next
    If Not rsPositions Is Nothing Then rsPositions.Close
    If Not rsDocument Is Nothing Then rsDocument.Close
    Set rsPositions = Nothing
    Set rsDocument = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    CalculateDocumentTotals = False
    modErrorHandler.HandleError MODULE_NAME, "CalculateDocumentTotals", Err
    Resume CleanExit
End Function

Public Function RecalculateDocument(ByVal DocumentId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsPositions As DAO.Recordset
    Dim sqlText As String
    Dim positionId As Long

    RecalculateDocument = False

    If DocumentId <= 0 Then
        Exit Function
    End If

    If Not EnsureDocumentCalculationSchema() Then
        Exit Function
    End If

    If Not DocumentExists(DocumentId) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".RecalculateDocument", _
            "Document recalculation skipped because DocumentId=" & CStr(DocumentId) & " does not exist."
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Exit Function
    End If

    sqlText = "SELECT [" & FIELD_DOCUMENT_POSITION_ID & "] FROM [" & TABLE_DOC_DOCUMENT_POSITION & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsPositions = db.OpenRecordset(sqlText, dbOpenSnapshot)

    If Not (rsPositions.BOF And rsPositions.EOF) Then
        rsPositions.MoveFirst
        Do Until rsPositions.EOF
            positionId = GetRecordsetLongValue(rsPositions, FIELD_DOCUMENT_POSITION_ID, 0)
            If positionId > 0 Then
                If Not CalculatePositionAmounts(positionId) Then
                    GoTo CleanExit
                End If
            End If
            rsPositions.MoveNext
        Loop
    End If

    RecalculateDocument = CalculateDocumentTotals(DocumentId)

CleanExit:
    On Error Resume Next
    If Not rsPositions Is Nothing Then rsPositions.Close
    Set rsPositions = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    RecalculateDocument = False
    modErrorHandler.HandleError MODULE_NAME, "RecalculateDocument", Err
    Resume CleanExit
End Function

Private Function NormalizeAdjustmentType(ByVal AdjustmentType As String) As String
    NormalizeAdjustmentType = UCase$(Trim$(AdjustmentType))
End Function

Private Function ResolveGrossAmount(ByVal NetAmount As Currency, ByVal VatAmount As Currency, ByVal VatRate As Double, ByVal VatMode As String) As Currency
    Dim normalizedVatMode As String

    normalizedVatMode = modVatHandler.NormalizeVatMode(VatMode)

    Select Case normalizedVatMode
        Case "EXCLUSIVE"
            ResolveGrossAmount = modVatHandler.CalculateGrossFromNet(NetAmount, VatRate)
        Case "INCLUSIVE", "NONE"
            ResolveGrossAmount = RoundCurrency(NetAmount)
        Case Else
            ResolveGrossAmount = RoundCurrency(NetAmount + VatAmount)
    End Select
End Function

Private Function ResolveDocumentVatMode(ByVal DocumentId As Long, ByVal DefaultValue As String) As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsDocument As DAO.Recordset
    Dim sqlText As String

    ResolveDocumentVatMode = DefaultValue

    If DocumentId <= 0 Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Exit Function
    End If

    sqlText = "SELECT [" & FIELD_VAT_MODE & "] FROM [" & TABLE_DOC_DOCUMENT & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsDocument = db.OpenRecordset(sqlText, dbOpenSnapshot)

    If Not (rsDocument.BOF And rsDocument.EOF) Then
        ResolveDocumentVatMode = GetRecordsetStringValue(rsDocument, FIELD_VAT_MODE, DefaultValue)
    End If

CleanExit:
    On Error Resume Next
    If Not rsDocument Is Nothing Then rsDocument.Close
    Set rsDocument = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    ResolveDocumentVatMode = DefaultValue
    modErrorHandler.HandleError MODULE_NAME, "ResolveDocumentVatMode", Err
    Resume CleanExit
End Function

Private Function GetPositionNetAmount(ByVal rsPosition As DAO.Recordset) As Currency
    If modDaoHelper.RecordsetHasField(rsPosition, FIELD_NET_AMOUNT) Then
        GetPositionNetAmount = GetRecordsetCurrencyValue(rsPosition, FIELD_NET_AMOUNT, 0)
    Else
        GetPositionNetAmount = GetRecordsetCurrencyValue(rsPosition, FIELD_LINE_TOTAL_NET, 0)
    End If
End Function

Private Function GetPositionVatAmount(ByVal rsPosition As DAO.Recordset) As Currency
    If modDaoHelper.RecordsetHasField(rsPosition, FIELD_VAT_AMOUNT) Then
        GetPositionVatAmount = GetRecordsetCurrencyValue(rsPosition, FIELD_VAT_AMOUNT, 0)
    Else
        GetPositionVatAmount = GetRecordsetCurrencyValue(rsPosition, FIELD_LINE_TOTAL_VAT, 0)
    End If
End Function

Private Function GetPositionGrossAmount(ByVal rsPosition As DAO.Recordset) As Currency
    If modDaoHelper.RecordsetHasField(rsPosition, FIELD_GROSS_AMOUNT) Then
        GetPositionGrossAmount = GetRecordsetCurrencyValue(rsPosition, FIELD_GROSS_AMOUNT, 0)
    Else
        GetPositionGrossAmount = GetRecordsetCurrencyValue(rsPosition, FIELD_LINE_TOTAL_GROSS, 0)
    End If
End Function

Private Function DocumentExists(ByVal DocumentId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsDocument As DAO.Recordset
    Dim sqlText As String

    If DocumentId <= 0 Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Exit Function
    End If

    sqlText = "SELECT [" & FIELD_DOCUMENT_ID & "] FROM [" & TABLE_DOC_DOCUMENT & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsDocument = db.OpenRecordset(sqlText, dbOpenSnapshot)

    DocumentExists = Not (rsDocument.BOF And rsDocument.EOF)

CleanExit:
    On Error Resume Next
    If Not rsDocument Is Nothing Then rsDocument.Close
    Set rsDocument = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    DocumentExists = False
    modErrorHandler.HandleError MODULE_NAME, "DocumentExists", Err
    Resume CleanExit
End Function

Private Function EnsureTextField(ByVal db As DAO.Database, ByVal TableName As String, ByVal FieldName As String, ByVal FieldSize As Long, ByVal DefaultValue As String) As Boolean
    On Error GoTo ErrorHandler

    If Not FieldExists(db, TableName, FieldName) Then
        db.Execute "ALTER TABLE [" & TableName & "] ADD COLUMN [" & FieldName & "] TEXT(" & CStr(FieldSize) & ");", dbFailOnError
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureTextField", _
            "Added field '" & FieldName & "' to table '" & TableName & "'."
    End If

    ApplyFieldDefaultValue db, TableName, FieldName, """" & Replace(DefaultValue, """", """""") & """"
    db.Execute "UPDATE [" & TableName & "] SET [" & FieldName & "]='" & Replace(DefaultValue, "'", "''") & "' WHERE [" & FieldName & "] IS NULL;", dbFailOnError

    EnsureTextField = True
    Exit Function

ErrorHandler:
    EnsureTextField = False
    modErrorHandler.HandleError MODULE_NAME, "EnsureTextField", Err
End Function

Private Function EnsureCurrencyField(ByVal db As DAO.Database, ByVal TableName As String, ByVal FieldName As String, ByVal DefaultValue As Currency) As Boolean
    On Error GoTo ErrorHandler

    If Not FieldExists(db, TableName, FieldName) Then
        db.Execute "ALTER TABLE [" & TableName & "] ADD COLUMN [" & FieldName & "] CURRENCY;", dbFailOnError
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureCurrencyField", _
            "Added field '" & FieldName & "' to table '" & TableName & "'."
    End If

    ApplyFieldDefaultValue db, TableName, FieldName, CStr(DefaultValue)
    db.Execute "UPDATE [" & TableName & "] SET [" & FieldName & "]=" & CStr(DefaultValue) & " WHERE [" & FieldName & "] IS NULL;", dbFailOnError

    EnsureCurrencyField = True
    Exit Function

ErrorHandler:
    EnsureCurrencyField = False
    modErrorHandler.HandleError MODULE_NAME, "EnsureCurrencyField", Err
End Function

Private Sub ApplyFieldDefaultValue(ByVal db As DAO.Database, ByVal TableName As String, ByVal FieldName As String, ByVal DefaultValueExpression As String)
    On Error Resume Next

    db.TableDefs(TableName).Fields(FieldName).DefaultValue = DefaultValueExpression
End Sub

Private Function TableExists(ByVal db As DAO.Database, ByVal TableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tdf As DAO.TableDef

    For Each tdf In db.TableDefs
        If UCase$(Trim$(tdf.Name)) = UCase$(Trim$(TableName)) Then
            TableExists = True
            Exit Function
        End If
    Next tdf
    Exit Function

ErrorHandler:
    TableExists = False
    modErrorHandler.HandleError MODULE_NAME, "TableExists", Err
End Function

Private Function FieldExists(ByVal db As DAO.Database, ByVal TableName As String, ByVal FieldName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    Set tdf = db.TableDefs(TableName)

    For Each fld In tdf.Fields
        If UCase$(Trim$(fld.Name)) = UCase$(Trim$(FieldName)) Then
            FieldExists = True
            Exit Function
        End If
    Next fld
    Exit Function

ErrorHandler:
    FieldExists = False
    modErrorHandler.HandleError MODULE_NAME, "FieldExists", Err
End Function

Private Function GetRecordsetStringValue(ByVal rs As DAO.Recordset, ByVal FieldName As String, ByVal DefaultValue As String) As String
    If modDaoHelper.RecordsetHasField(rs, FieldName) Then
        GetRecordsetStringValue = Trim$(modDaoHelper.NzString(rs.Fields(FieldName).Value, DefaultValue))
    Else
        GetRecordsetStringValue = DefaultValue
    End If
End Function

Private Function GetRecordsetCurrencyValue(ByVal rs As DAO.Recordset, ByVal FieldName As String, ByVal DefaultValue As Currency) As Currency
    Dim rawValue As String

    If modDaoHelper.RecordsetHasField(rs, FieldName) Then
        rawValue = modDaoHelper.NzString(rs.Fields(FieldName).Value, CStr(DefaultValue))
        If IsNumeric(rawValue) Then
            GetRecordsetCurrencyValue = CCur(rawValue)
        Else
            GetRecordsetCurrencyValue = DefaultValue
        End If
    Else
        GetRecordsetCurrencyValue = DefaultValue
    End If
End Function

Private Function GetRecordsetDoubleValue(ByVal rs As DAO.Recordset, ByVal FieldName As String, ByVal DefaultValue As Double) As Double
    Dim rawValue As String

    If modDaoHelper.RecordsetHasField(rs, FieldName) Then
        rawValue = modDaoHelper.NzString(rs.Fields(FieldName).Value, CStr(DefaultValue))
        If IsNumeric(rawValue) Then
            GetRecordsetDoubleValue = CDbl(rawValue)
        Else
            GetRecordsetDoubleValue = DefaultValue
        End If
    Else
        GetRecordsetDoubleValue = DefaultValue
    End If
End Function

Private Function GetRecordsetLongValue(ByVal rs As DAO.Recordset, ByVal FieldName As String, ByVal DefaultValue As Long) As Long
    If modDaoHelper.RecordsetHasField(rs, FieldName) Then
        GetRecordsetLongValue = modDaoHelper.NzLong(rs.Fields(FieldName).Value, DefaultValue)
    Else
        GetRecordsetLongValue = DefaultValue
    End If
End Function

Private Sub SetRecordsetValue(ByVal rs As DAO.Recordset, ByVal FieldName As String, ByVal FieldValue As Variant)
    If modDaoHelper.RecordsetHasField(rs, FieldName) Then
        rs.Fields(FieldName).Value = FieldValue
    End If
End Sub

Private Function RoundCurrency(ByVal Amount As Currency) As Currency
    RoundCurrency = CCur(Round(CDbl(Amount), 2))
End Function
