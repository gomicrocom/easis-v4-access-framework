Attribute VB_Name = "modDocumentCalculationService"
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

    If Not modDbSchema.TableExists(db, TABLE_DOC_DOCUMENT_POSITION) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".EnsureDocumentCalculationSchema", _
            "Table '" & TABLE_DOC_DOCUMENT_POSITION & "' is not available."
        Exit Function
    End If

    If Not modDbSchema.TableExists(db, TABLE_DOC_DOCUMENT) Then
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
    ByVal baseAmount As Currency, _
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
            CalculateAdjustmentAmount = RoundCurrency(CCur(CDbl(baseAmount) * (CDbl(AdjustmentValue) / 100#)))

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
    Dim SqlText As String
    Dim DocumentId As Long
    Dim quantity As Double
    Dim UnitPrice As Currency
    Dim vatRate As Double
    Dim VatMode As String
    Dim baseAmount As Currency
    Dim discountAmount As Currency
    Dim surchargeAmount As Currency
    Dim NetAmount As Currency
    Dim VatAmount As Currency
    Dim GrossAmount As Currency
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

    SqlText = "SELECT * FROM [" & TABLE_DOC_DOCUMENT_POSITION & "] WHERE [" & FIELD_DOCUMENT_POSITION_ID & "]=" & CStr(DocumentPositionId) & ";"
    Set rsPosition = db.OpenRecordset(SqlText, dbOpenDynaset)

    If rsPosition.BOF And rsPosition.EOF Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CalculatePositionAmounts", _
            "Position calculation skipped because DocumentPositionId=" & CStr(DocumentPositionId) & " does not exist."
        GoTo CleanExit
    End If

    quantity = GetRecordsetDoubleValue(rsPosition, FIELD_QUANTITY, 0)
    UnitPrice = GetRecordsetCurrencyValue(rsPosition, FIELD_UNIT_PRICE, 0)
    vatRate = GetRecordsetDoubleValue(rsPosition, FIELD_VAT_RATE, 0)
    DocumentId = GetRecordsetLongValue(rsPosition, FIELD_DOCUMENT_ID, 0)
    VatMode = ResolveDocumentVatMode(DocumentId, modVatHandler.GetVatMode())

    discountType = GetRecordsetStringValue(rsPosition, FIELD_DISCOUNT_TYPE, ADJUSTMENT_TYPE_NONE)
    discountValue = GetRecordsetCurrencyValue(rsPosition, FIELD_DISCOUNT_VALUE, 0)
    surchargeType = GetRecordsetStringValue(rsPosition, FIELD_SURCHARGE_TYPE, ADJUSTMENT_TYPE_NONE)
    surchargeValue = GetRecordsetCurrencyValue(rsPosition, FIELD_SURCHARGE_VALUE, 0)

    baseAmount = RoundCurrency(CCur(quantity * CDbl(UnitPrice)))
    discountAmount = CalculateAdjustmentAmount(baseAmount, discountType, discountValue)
    If discountAmount > baseAmount Then
        discountAmount = baseAmount
    End If

    discountedBaseAmount = RoundCurrency(baseAmount - discountAmount)
    surchargeAmount = CalculateAdjustmentAmount(discountedBaseAmount, surchargeType, surchargeValue)

    NetAmount = RoundCurrency(discountedBaseAmount + surchargeAmount)
    If NetAmount < 0 Then
        NetAmount = 0
    End If

    VatAmount = modVatHandler.CalculateVatAmount(NetAmount, vatRate, VatMode)
    GrossAmount = ResolveGrossAmount(NetAmount, VatAmount, vatRate, VatMode)

    rsPosition.Edit
    SetRecordsetValue rsPosition, FIELD_LINE_BASE_AMOUNT, baseAmount
    SetRecordsetValue rsPosition, FIELD_LINE_DISCOUNT_AMOUNT, discountAmount
    SetRecordsetValue rsPosition, FIELD_LINE_SURCHARGE_AMOUNT, surchargeAmount
    SetRecordsetValue rsPosition, FIELD_NET_AMOUNT, NetAmount
    SetRecordsetValue rsPosition, FIELD_VAT_AMOUNT, VatAmount
    SetRecordsetValue rsPosition, FIELD_GROSS_AMOUNT, GrossAmount
    SetRecordsetValue rsPosition, FIELD_LINE_TOTAL_NET, NetAmount
    SetRecordsetValue rsPosition, FIELD_LINE_TOTAL_VAT, VatAmount
    SetRecordsetValue rsPosition, FIELD_LINE_TOTAL_GROSS, GrossAmount
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
    Dim SqlText As String
    Dim subtotalNetAmount As Currency
    Dim VatAmount As Currency
    Dim grossPositionAmount As Currency
    Dim headerDiscountAmount As Currency
    Dim headerSurchargeAmount As Currency
    Dim NetAmount As Currency
    Dim GrossAmount As Currency
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

    SqlText = "SELECT * FROM [" & TABLE_DOC_DOCUMENT & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsDocument = db.OpenRecordset(SqlText, dbOpenDynaset)

    If rsDocument.BOF And rsDocument.EOF Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CalculateDocumentTotals", _
            "Document total calculation skipped because DocumentId=" & CStr(DocumentId) & " does not exist."
        GoTo CleanExit
    End If

    SqlText = "SELECT * FROM [" & TABLE_DOC_DOCUMENT_POSITION & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsPositions = db.OpenRecordset(SqlText, dbOpenSnapshot)

    If Not (rsPositions.BOF And rsPositions.EOF) Then
        rsPositions.MoveFirst
        Do Until rsPositions.EOF
            subtotalNetAmount = subtotalNetAmount + GetPositionNetAmount(rsPositions)
            VatAmount = VatAmount + GetPositionVatAmount(rsPositions)
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
    NetAmount = subtotalNetAmount
    GrossAmount = grossPositionAmount

    rsDocument.Edit
    SetRecordsetValue rsDocument, FIELD_SUBTOTAL_NET_AMOUNT, subtotalNetAmount
    SetRecordsetValue rsDocument, FIELD_HEADER_DISCOUNT_AMOUNT, headerDiscountAmount
    SetRecordsetValue rsDocument, FIELD_HEADER_SURCHARGE_AMOUNT, headerSurchargeAmount
    SetRecordsetValue rsDocument, FIELD_NET_AMOUNT, NetAmount
    SetRecordsetValue rsDocument, FIELD_VAT_AMOUNT, VatAmount
    SetRecordsetValue rsDocument, FIELD_GROSS_AMOUNT, GrossAmount
    SetRecordsetValue rsDocument, FIELD_TOTAL_NET, NetAmount
    SetRecordsetValue rsDocument, FIELD_TOTAL_VAT, VatAmount
    SetRecordsetValue rsDocument, FIELD_TOTAL_GROSS, GrossAmount
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
    Dim SqlText As String
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

    SqlText = "SELECT [" & FIELD_DOCUMENT_POSITION_ID & "] FROM [" & TABLE_DOC_DOCUMENT_POSITION & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsPositions = db.OpenRecordset(SqlText, dbOpenSnapshot)

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

Private Function ResolveGrossAmount(ByVal NetAmount As Currency, ByVal VatAmount As Currency, ByVal vatRate As Double, ByVal VatMode As String) As Currency
    Dim normalizedVatMode As String

    normalizedVatMode = modVatHandler.NormalizeVatMode(VatMode)

    Select Case normalizedVatMode
        Case "EXCLUSIVE"
            ResolveGrossAmount = modVatHandler.CalculateGrossFromNet(NetAmount, vatRate)
        Case "INCLUSIVE", "NONE"
            ResolveGrossAmount = RoundCurrency(NetAmount)
        Case Else
            ResolveGrossAmount = RoundCurrency(NetAmount + VatAmount)
    End Select
End Function

Private Function ResolveDocumentVatMode(ByVal DocumentId As Long, ByVal defaultValue As String) As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsDocument As DAO.Recordset
    Dim SqlText As String

    ResolveDocumentVatMode = defaultValue

    If DocumentId <= 0 Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Exit Function
    End If

    SqlText = "SELECT [" & FIELD_VAT_MODE & "] FROM [" & TABLE_DOC_DOCUMENT & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsDocument = db.OpenRecordset(SqlText, dbOpenSnapshot)

    If Not (rsDocument.BOF And rsDocument.EOF) Then
        ResolveDocumentVatMode = GetRecordsetStringValue(rsDocument, FIELD_VAT_MODE, defaultValue)
    End If

CleanExit:
    On Error Resume Next
    If Not rsDocument Is Nothing Then rsDocument.Close
    Set rsDocument = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    ResolveDocumentVatMode = defaultValue
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
    Dim SqlText As String

    If DocumentId <= 0 Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Exit Function
    End If

    SqlText = "SELECT [" & FIELD_DOCUMENT_ID & "] FROM [" & TABLE_DOC_DOCUMENT & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsDocument = db.OpenRecordset(SqlText, dbOpenSnapshot)

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

Private Function EnsureTextField(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByVal FieldSize As Long, ByVal defaultValue As String) As Boolean
    On Error GoTo ErrorHandler

    If Not modDbSchema.FieldExists(db, tableName, fieldName) Then
        db.Execute "ALTER TABLE [" & tableName & "] ADD COLUMN [" & fieldName & "] TEXT(" & CStr(FieldSize) & ");", dbFailOnError
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureTextField", _
            "Added field '" & fieldName & "' to table '" & tableName & "'."
    End If

    ApplyFieldDefaultValue db, tableName, fieldName, """" & Replace(defaultValue, """", """""") & """"
    db.Execute "UPDATE [" & tableName & "] SET [" & fieldName & "]='" & Replace(defaultValue, "'", "''") & "' WHERE [" & fieldName & "] IS NULL;", dbFailOnError

    EnsureTextField = True
    Exit Function

ErrorHandler:
    EnsureTextField = False
    modErrorHandler.HandleError MODULE_NAME, "EnsureTextField", Err
End Function

Private Function EnsureCurrencyField(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByVal defaultValue As Currency) As Boolean
    On Error GoTo ErrorHandler

    If Not modDbSchema.FieldExists(db, tableName, fieldName) Then
        db.Execute "ALTER TABLE [" & tableName & "] ADD COLUMN [" & fieldName & "] CURRENCY;", dbFailOnError
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureCurrencyField", _
            "Added field '" & fieldName & "' to table '" & tableName & "'."
    End If

    ApplyFieldDefaultValue db, tableName, fieldName, CStr(defaultValue)
    db.Execute "UPDATE [" & tableName & "] SET [" & fieldName & "]=" & CStr(defaultValue) & " WHERE [" & fieldName & "] IS NULL;", dbFailOnError

    EnsureCurrencyField = True
    Exit Function

ErrorHandler:
    EnsureCurrencyField = False
    modErrorHandler.HandleError MODULE_NAME, "EnsureCurrencyField", Err
End Function

Private Sub ApplyFieldDefaultValue(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByVal DefaultValueExpression As String)
    On Error Resume Next

    db.TableDefs(tableName).Fields(fieldName).defaultValue = DefaultValueExpression
End Sub



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


