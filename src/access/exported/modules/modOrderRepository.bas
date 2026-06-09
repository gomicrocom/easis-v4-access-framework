Attribute VB_Name = "modOrderRepository"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modOrderRepository
' Purpose   : DAO persistence helpers for Phase 1 sales orders and order lines.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modOrderRepository"

Private Const TABLE_ORD_ORDER As String = "ord_order"
Private Const TABLE_ORD_ORDER_LINE As String = "ord_order_line"
Private Const TABLE_TEN_NUMBERRANGE As String = "ten_numberrange"

Private Const FIELD_ORDER_ID As String = "order_id"
Private Const FIELD_ORDER_NO As String = "order_no"
Private Const FIELD_ORDER_TYPE_CODE As String = "order_type_code"
Private Const FIELD_ORDER_STATUS_CODE As String = "order_status_code"
Private Const FIELD_CUSTOMER_ADDRESS_ID As String = "customer_address_id"
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
Private Const FIELD_NOTES_TEXT As String = "notes_text"
Private Const FIELD_INTERNAL_NOTES_TEXT As String = "internal_notes_text"
Private Const FIELD_CREATED_AT As String = "created_at"
Private Const FIELD_CREATED_BY As String = "created_by"
Private Const FIELD_UPDATED_AT As String = "updated_at"
Private Const FIELD_UPDATED_BY As String = "updated_by"

Private Const FIELD_ORDER_LINE_ID As String = "order_line_id"
Private Const FIELD_LINE_NO As String = "line_no"
Private Const FIELD_ARTICLE_ID As String = "article_id"
Private Const FIELD_ARTICLE_NO As String = "article_no"
Private Const FIELD_LINE_TYPE_CODE As String = "line_type_code"
Private Const FIELD_DESCRIPTION_TEXT As String = "description_text"
Private Const FIELD_QUANTITY As String = "quantity"
Private Const FIELD_UNIT_CODE As String = "unit_code"
Private Const FIELD_UNIT_PRICE As String = "unit_price"
Private Const FIELD_VAT_CODE As String = "vat_code"
Private Const FIELD_VAT_RATE As String = "vat_rate"

Private Const FIELD_NR_DOCUMENT_TYPE_CODE As String = "document_type_code"
Private Const FIELD_NR_FISCAL_YEAR As String = "fiscal_year"
Private Const FIELD_NR_PREFIX As String = "prefix"
Private Const FIELD_NR_CURRENT_VALUE As String = "current_value"
Private Const FIELD_NR_FORMAT_MASK As String = "format_mask"
Private Const FIELD_NR_IS_ACTIVE As String = "is_active"

Public Const ORDER_TYPE_SALES_ORDER As String = "SO"
Public Const ORDER_STATUS_DRAFT As String = "DRAFT"
Public Const ORDER_STATUS_OPEN As String = "OPEN"
Public Const ORDER_STATUS_CONFIRMED As String = "CONFIRMED"
Public Const ORDER_STATUS_CONVERTED As String = "CONVERTED"
Public Const ORDER_STATUS_CANCELLED As String = "CANCELLED"
Public Const ORDER_STATUS_CLOSED As String = "CLOSED"

Private Const DEFAULT_ORDER_FORMAT_MASK As String = "{PREFIX}-{YEAR}-{NUMBER:0000}"

Public Function EnsureOrderRepositoryReady() As Boolean
    On Error GoTo ErrorHandler

    EnsureOrderRepositoryReady = False

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".EnsureOrderRepositoryReady", _
            "Order repository initialization skipped because backend configuration is not valid."
        Exit Function
    End If

    If Not modBasicModuleSchema.EnsureOrderPhase1Schema() Then
        Exit Function
    End If

    If Not EnsureSalesOrderNumberRange(Year(Date)) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".EnsureOrderRepositoryReady", _
            "Sales-order number range could not be prepared. Order schema is still available."
    End If

    EnsureOrderRepositoryReady = True
    Exit Function

ErrorHandler:
    EnsureOrderRepositoryReady = False
    modErrorHandler.HandleError MODULE_NAME, "EnsureOrderRepositoryReady", Err
End Function

Public Function EnsureSalesOrderNumberRange(Optional ByVal FiscalYear As Long = 0) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim effectiveYear As Long

    EnsureSalesOrderNumberRange = False

    If FiscalYear <= 0 Then
        effectiveYear = Year(Date)
    Else
        effectiveYear = FiscalYear
    End If

    If Not modDb.ValidateBackendConfiguration() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Exit Function
    End If

    If Not TableExists(db, TABLE_TEN_NUMBERRANGE) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".EnsureSalesOrderNumberRange", _
            "Table '" & TABLE_TEN_NUMBERRANGE & "' is not available."
        Exit Function
    End If

    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_TEN_NUMBERRANGE & "];", dbOpenDynaset)

    If FindNumberRangeRow(rs, ORDER_TYPE_SALES_ORDER, effectiveYear) Then
        If modDaoHelper.RecordsetHasField(rs, FIELD_NR_PREFIX) Then
            If LenB(Trim$(modDaoHelper.NzString(rs.Fields(FIELD_NR_PREFIX).Value))) = 0 Then
                rs.Edit
                rs.Fields(FIELD_NR_PREFIX).Value = ORDER_TYPE_SALES_ORDER
                SetUpdatedAuditFields rs
                rs.Update
            End If
        End If

        If modDaoHelper.RecordsetHasField(rs, FIELD_NR_FORMAT_MASK) Then
            If LenB(Trim$(modDaoHelper.NzString(rs.Fields(FIELD_NR_FORMAT_MASK).Value))) = 0 Then
                rs.Edit
                rs.Fields(FIELD_NR_FORMAT_MASK).Value = DEFAULT_ORDER_FORMAT_MASK
                SetUpdatedAuditFields rs
                rs.Update
            End If
        End If

        If modDaoHelper.RecordsetHasField(rs, FIELD_NR_IS_ACTIVE) Then
            If Not modDaoHelper.NzBoolean(rs.Fields(FIELD_NR_IS_ACTIVE).Value, False) Then
                rs.Edit
                rs.Fields(FIELD_NR_IS_ACTIVE).Value = True
                SetUpdatedAuditFields rs
                rs.Update
            End If
        End If

        EnsureSalesOrderNumberRange = True
        GoTo CleanExit
    End If

    rs.AddNew
    SetRecordsetValue rs, FIELD_NR_DOCUMENT_TYPE_CODE, ORDER_TYPE_SALES_ORDER
    SetRecordsetValue rs, FIELD_NR_FISCAL_YEAR, effectiveYear
    SetRecordsetValue rs, FIELD_NR_PREFIX, ORDER_TYPE_SALES_ORDER
    SetRecordsetValue rs, FIELD_NR_CURRENT_VALUE, 0
    SetRecordsetValue rs, FIELD_NR_FORMAT_MASK, DEFAULT_ORDER_FORMAT_MASK
    SetRecordsetValue rs, FIELD_NR_IS_ACTIVE, True
    SetCreatedAuditFields rs
    SetUpdatedAuditFields rs
    rs.Update

    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureSalesOrderNumberRange", _
        "Prepared sales-order number range for FiscalYear=" & CStr(effectiveYear) & "."

    EnsureSalesOrderNumberRange = True

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    EnsureSalesOrderNumberRange = False
    modErrorHandler.HandleError MODULE_NAME, "EnsureSalesOrderNumberRange", Err
    Resume CleanExit
End Function

Public Function GetNextSalesOrderNumber(Optional ByVal OrderDate As Date = 0) As String
    On Error GoTo ErrorHandler

    Dim effectiveDate As Date

    If OrderDate = 0 Then
        effectiveDate = Date
    Else
        effectiveDate = OrderDate
    End If

    If Not EnsureSalesOrderNumberRange(Year(effectiveDate)) Then
        Exit Function
    End If

    GetNextSalesOrderNumber = modNumberingHandler.GetNextDocumentNumber(ORDER_TYPE_SALES_ORDER, effectiveDate)
    Exit Function

ErrorHandler:
    GetNextSalesOrderNumber = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "GetNextSalesOrderNumber", Err
End Function

Public Function CreateSalesOrderHeader( _
    Optional ByVal CustomerAddressId As Long = 0, _
    Optional ByVal OrderDate As Date = 0, _
    Optional ByVal CustomerName As String = "", _
    Optional ByVal DeliveryDate As Date = 0, _
    Optional ByVal ValidUntil As Date = 0, _
    Optional ByVal ReferenceText As String = "", _
    Optional ByVal ExternalReference As String = "", _
    Optional ByVal languageCode As String = "", _
    Optional ByVal CurrencyCode As String = "", _
    Optional ByVal PaymentTermCode As String = "", _
    Optional ByVal VatMode As String = "", _
    Optional ByVal NotesText As String = "", _
    Optional ByVal InternalNotesText As String = "" _
) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim effectiveDate As Date
    Dim effectiveCustomerName As String
    Dim orderNo As String

    CreateSalesOrderHeader = 0

    If Not EnsureOrderRepositoryReady() Then
        Exit Function
    End If

    effectiveDate = IIf(OrderDate = 0, Date, OrderDate)
    orderNo = GetNextSalesOrderNumber(effectiveDate)
    effectiveCustomerName = Trim$(CustomerName)

    If CustomerAddressId > 0 Then
        If modAddressRepository.AddressExists(CustomerAddressId) Then
            effectiveCustomerName = modAddressRepository.GetAddressDisplayName(CustomerAddressId, effectiveCustomerName)
        End If
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset(TABLE_ORD_ORDER, dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    SetRecordsetValue rs, FIELD_ORDER_NO, orderNo
    SetRecordsetValue rs, FIELD_ORDER_TYPE_CODE, ORDER_TYPE_SALES_ORDER
    SetRecordsetValue rs, FIELD_ORDER_STATUS_CODE, ORDER_STATUS_DRAFT
    SetRecordsetValue rs, FIELD_CUSTOMER_ADDRESS_ID, CustomerAddressId
    SetRecordsetValue rs, FIELD_CUSTOMER_NAME, effectiveCustomerName
    SetRecordsetValue rs, FIELD_ORDER_DATE, effectiveDate
    If DeliveryDate <> 0 Then SetRecordsetValue rs, FIELD_DELIVERY_DATE, DeliveryDate
    If ValidUntil <> 0 Then SetRecordsetValue rs, FIELD_VALID_UNTIL, ValidUntil
    SetRecordsetValue rs, FIELD_REFERENCE_TEXT, Trim$(ReferenceText)
    SetRecordsetValue rs, FIELD_EXTERNAL_REFERENCE, Trim$(ExternalReference)
    SetRecordsetValue rs, FIELD_LANGUAGE_CODE, ResolveLanguageCode(languageCode)
    SetRecordsetValue rs, FIELD_CURRENCY_CODE, ResolveCurrencyCode(CurrencyCode)
    SetRecordsetValue rs, FIELD_PAYMENT_TERM_CODE, Trim$(PaymentTermCode)
    SetRecordsetValue rs, FIELD_VAT_MODE, ResolveVatMode(VatMode)
    SetRecordsetValue rs, FIELD_NOTES_TEXT, NotesText
    SetRecordsetValue rs, FIELD_INTERNAL_NOTES_TEXT, InternalNotesText
    SetCreatedAuditFields rs
    SetUpdatedAuditFields rs
    rs.Update

    rs.Bookmark = rs.LastModified
    CreateSalesOrderHeader = modDaoHelper.NzLong(rs.Fields(FIELD_ORDER_ID).Value, 0)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    CreateSalesOrderHeader = 0
    modErrorHandler.HandleError MODULE_NAME, "CreateSalesOrderHeader", Err
    Resume CleanExit
End Function

Public Function DeleteOrderLines(ByVal OrderId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database

    DeleteOrderLines = False

    If OrderId <= 0 Then
        Exit Function
    End If

    If Not EnsureOrderRepositoryReady() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    db.Execute "DELETE FROM [" & TABLE_ORD_ORDER_LINE & "] WHERE [" & FIELD_ORDER_ID & "]=" & CStr(OrderId) & ";", dbFailOnError

    DeleteOrderLines = True

CleanExit:
    Set db = Nothing
    Exit Function

ErrorHandler:
    DeleteOrderLines = False
    modErrorHandler.HandleError MODULE_NAME, "DeleteOrderLines", Err
    Resume CleanExit
End Function

Public Function CreateOrderLine( _
    ByVal OrderId As Long, _
    ByVal LineNo As Long, _
    ByVal DescriptionText As String, _
    ByVal quantity As Double, _
    ByVal UnitPrice As Currency, _
    Optional ByVal ArticleId As Long = 0, _
    Optional ByVal ArticleNo As String = "", _
    Optional ByVal LineTypeCode As String = "", _
    Optional ByVal UnitCode As String = "", _
    Optional ByVal VatCode As String = "", _
    Optional ByVal vatRate As Double = -1 _
) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim effectiveVatRate As Double

    CreateOrderLine = 0

    If OrderId <= 0 Then
        Exit Function
    End If

    If Not EnsureOrderRepositoryReady() Then
        Exit Function
    End If

    If vatRate < 0 Then
        effectiveVatRate = modVatHandler.GetVatRate()
    Else
        effectiveVatRate = vatRate
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset(TABLE_ORD_ORDER_LINE, dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    SetRecordsetValue rs, FIELD_ORDER_ID, OrderId
    SetRecordsetValue rs, FIELD_LINE_NO, LineNo
    SetRecordsetValue rs, FIELD_ARTICLE_ID, ArticleId
    SetRecordsetValue rs, FIELD_ARTICLE_NO, Trim$(ArticleNo)
    SetRecordsetValue rs, FIELD_LINE_TYPE_CODE, Trim$(LineTypeCode)
    SetRecordsetValue rs, FIELD_DESCRIPTION_TEXT, Trim$(DescriptionText)
    SetRecordsetValue rs, FIELD_QUANTITY, quantity
    SetRecordsetValue rs, FIELD_UNIT_CODE, Trim$(UnitCode)
    SetRecordsetValue rs, FIELD_UNIT_PRICE, UnitPrice
    SetRecordsetValue rs, FIELD_VAT_CODE, Trim$(VatCode)
    SetRecordsetValue rs, FIELD_VAT_RATE, effectiveVatRate
    SetCreatedAuditFields rs
    SetUpdatedAuditFields rs
    rs.Update

    rs.Bookmark = rs.LastModified
    CreateOrderLine = modDaoHelper.NzLong(rs.Fields(FIELD_ORDER_LINE_ID).Value, 0)

    If CreateOrderLine > 0 Then
        Call modOrderCalculationService.CalculateOrderLineAmounts(CreateOrderLine)
        Call modOrderCalculationService.CalculateOrderTotals(OrderId)
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    CreateOrderLine = 0
    modErrorHandler.HandleError MODULE_NAME, "CreateOrderLine", Err
    Resume CleanExit
End Function

Public Function OrderExists(ByVal OrderId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If OrderId <= 0 Then
        Exit Function
    End If

    If Not EnsureOrderRepositoryReady() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT [" & FIELD_ORDER_ID & "] FROM [" & TABLE_ORD_ORDER & "] WHERE [" & FIELD_ORDER_ID & "]=" & CStr(OrderId) & ";", dbOpenSnapshot)
    OrderExists = Not (rs.BOF And rs.EOF)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    OrderExists = False
    modErrorHandler.HandleError MODULE_NAME, "OrderExists", Err
    Resume CleanExit
End Function

Private Function ResolveCurrencyCode(ByVal explicitValue As String) As String
    ResolveCurrencyCode = Trim$(explicitValue)
    If LenB(ResolveCurrencyCode) = 0 Then
        ResolveCurrencyCode = modTenantRepository.GetTenantParameter("CURRENCY_CODE", "CHF")
    End If
End Function

Private Function ResolveLanguageCode(ByVal explicitValue As String) As String
    ResolveLanguageCode = Trim$(explicitValue)
    If LenB(ResolveLanguageCode) = 0 Then
        ResolveLanguageCode = modFwTranslationRuntime.GetCurrentLanguageCode()
    End If
End Function

Private Function ResolveVatMode(ByVal explicitValue As String) As String
    If LenB(Trim$(explicitValue)) = 0 Then
        ResolveVatMode = modVatHandler.GetVatMode()
    Else
        ResolveVatMode = modVatHandler.NormalizeVatMode(explicitValue)
    End If
End Function

Private Function ResolveAuditUserName() As String
    ResolveAuditUserName = Trim$(currentUserId)
    If LenB(ResolveAuditUserName) = 0 Then
        ResolveAuditUserName = Trim$(CurrentUserName)
    End If
    If LenB(ResolveAuditUserName) = 0 Then
        ResolveAuditUserName = "SYSTEM"
    End If
End Function

Private Sub SetCreatedAuditFields(ByVal rs As DAO.Recordset)
    SetRecordsetValue rs, FIELD_CREATED_AT, Now()
    SetRecordsetValue rs, FIELD_CREATED_BY, ResolveAuditUserName()
End Sub

Private Sub SetUpdatedAuditFields(ByVal rs As DAO.Recordset)
    SetRecordsetValue rs, FIELD_UPDATED_AT, Now()
    SetRecordsetValue rs, FIELD_UPDATED_BY, ResolveAuditUserName()
End Sub

Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tdf As DAO.tableDef

    For Each tdf In db.TableDefs
        If StrComp(Trim$(tdf.Name), Trim$(tableName), vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next tdf
    Exit Function

ErrorHandler:
    TableExists = False
End Function

Private Function FindNumberRangeRow(ByVal rs As DAO.Recordset, ByVal DocumentTypeCode As String, ByVal FiscalYear As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim targetType As String

    targetType = UCase$(Trim$(DocumentTypeCode))
    If rs.BOF And rs.EOF Then Exit Function

    rs.MoveFirst
    Do Until rs.EOF
        If UCase$(Trim$(modDaoHelper.NzString(rs.Fields(FIELD_NR_DOCUMENT_TYPE_CODE).Value))) = targetType _
            And modDaoHelper.NzLong(rs.Fields(FIELD_NR_FISCAL_YEAR).Value, 0) = FiscalYear Then
            FindNumberRangeRow = True
            Exit Function
        End If
        rs.MoveNext
    Loop
    Exit Function

ErrorHandler:
    FindNumberRangeRow = False
End Function

Private Sub SetRecordsetValue(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal fieldValue As Variant)
    If modDaoHelper.RecordsetHasField(rs, fieldName) Then
        rs.Fields(fieldName).Value = fieldValue
    End If
End Sub