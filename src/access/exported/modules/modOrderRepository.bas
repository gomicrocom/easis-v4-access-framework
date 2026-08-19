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
Private Const TABLE_TMP_ORDER As String = "tmp_order"
Private Const TABLE_TMP_ORDER_LINE As String = "tmp_order_line"
Private Const TABLE_TEN_NUMBERRANGE As String = "ten_numberrange"

Private Const FIELD_ORDER_ID As String = "order_id"
Private Const FIELD_TMP_ORDER_ID As String = "tmp_order_id"
Private Const FIELD_SESSION_ID As String = "session_id"
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
Private Const FIELD_HEADER_DISCOUNT_TYPE As String = "header_discount_type"
Private Const FIELD_HEADER_DISCOUNT_VALUE As String = "header_discount_value"
Private Const FIELD_HEADER_DISCOUNT_AMOUNT As String = "header_discount_amount"
Private Const FIELD_HEADER_SURCHARGE_TYPE As String = "header_surcharge_type"
Private Const FIELD_HEADER_SURCHARGE_VALUE As String = "header_surcharge_value"
Private Const FIELD_HEADER_SURCHARGE_AMOUNT As String = "header_surcharge_amount"
Private Const FIELD_SUBTOTAL_NET_AMOUNT As String = "subtotal_net_amount"
Private Const FIELD_NET_AMOUNT As String = "net_amount"
Private Const FIELD_VAT_AMOUNT As String = "vat_amount"
Private Const FIELD_GROSS_AMOUNT As String = "gross_amount"
Private Const FIELD_NOTES_TEXT As String = "notes_text"
Private Const FIELD_INTERNAL_NOTES_TEXT As String = "internal_notes_text"
Private Const FIELD_CREATED_AT As String = "created_at"
Private Const FIELD_CREATED_BY As String = "created_by"
Private Const FIELD_UPDATED_AT As String = "updated_at"
Private Const FIELD_UPDATED_BY As String = "updated_by"

Private Const FIELD_ORDER_LINE_ID As String = "order_line_id"
Private Const FIELD_TMP_ORDER_LINE_ID As String = "tmp_order_line_id"
Private Const FIELD_LINE_NO As String = "line_no"
Private Const FIELD_ARTICLE_ID As String = "article_id"
Private Const FIELD_ARTICLE_NO As String = "article_no"
Private Const FIELD_LINE_TYPE_CODE As String = "line_type_code"
Private Const FIELD_DESCRIPTION_TEXT As String = "description_text"
Private Const FIELD_QUANTITY As String = "quantity"
Private Const FIELD_UNIT_CODE As String = "unit_code"
Private Const FIELD_UNIT_PRICE As String = "unit_price"
Private Const FIELD_DISCOUNT_TYPE As String = "discount_type"
Private Const FIELD_DISCOUNT_VALUE As String = "discount_value"
Private Const FIELD_LINE_DISCOUNT_AMOUNT As String = "line_discount_amount"
Private Const FIELD_SURCHARGE_TYPE As String = "surcharge_type"
Private Const FIELD_SURCHARGE_VALUE As String = "surcharge_value"
Private Const FIELD_LINE_SURCHARGE_AMOUNT As String = "line_surcharge_amount"
Private Const FIELD_LINE_BASE_AMOUNT As String = "line_base_amount"
Private Const FIELD_LINE_NET_AMOUNT As String = "line_net_amount"
Private Const FIELD_LINE_VAT_AMOUNT As String = "line_vat_amount"
Private Const FIELD_LINE_GROSS_AMOUNT As String = "line_gross_amount"
Private Const FIELD_SORT_ORDER As String = "sort_order"
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
Private Const DEFAULT_TENANT_COUNTRY_CODE As String = "CH"
Private Const DEFAULT_ZERO_VAT_CODE As String = "CH_ZERO"
Private Const DEFAULT_STANDARD_VAT_CODE As String = "CH_STANDARD"

Public Function EnsureOrderRepositoryReady() As Boolean
    On Error GoTo ErrorHandler

    EnsureOrderRepositoryReady = False

    If Not EnsureOrderRuntimeContext() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".EnsureOrderRepositoryReady", _
            "Order repository initialization skipped because runtime context could not be initialized."
        Exit Function
    End If

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

Private Function EnsureOrderRuntimeContext() As Boolean
    On Error GoTo ErrorHandler

    If IsBootstrapped Then
        EnsureOrderRuntimeContext = True
        Exit Function
    End If

    If LenB(Trim$(ConfigFilePath)) > 0 And modTenantContext.IsTenantInitialized Then
        EnsureOrderRuntimeContext = True
        Exit Function
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureOrderRuntimeContext", _
        "Runtime context not initialized. Triggering bootstrap for repository access."
    EnsureOrderRuntimeContext = modBootstrap.EnsureBootstrapped()
    Exit Function

ErrorHandler:
    EnsureOrderRuntimeContext = False
    modErrorHandler.HandleError MODULE_NAME, "EnsureOrderRuntimeContext", Err
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

    Set db = modDb.GetCurrentTenantDatabase()
    If db Is Nothing Then
        Exit Function
    End If

    If Not modDbSchema.TableExists(db, TABLE_TEN_NUMBERRANGE) Then
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
    Optional ByVal InvoiceAddressId As Long = 0, _
    Optional ByVal DeliveryAddressId As Long = 0, _
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
    Optional ByVal VatCode As String = "", _
    Optional ByVal vatRate As Double = -1, _
    Optional ByVal NotesText As String = "", _
    Optional ByVal InternalNotesText As String = "" _
) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim effectiveDate As Date
    Dim effectiveCustomerName As String
    Dim orderNo As String
    Dim effectiveInvoiceAddressId As Long
    Dim effectiveDeliveryAddressId As Long
    Dim vatContext As Object
    Dim effectiveVatCode As String
    Dim effectiveVatMode As String
    Dim effectiveVatRate As Double

    CreateSalesOrderHeader = 0

    If Not EnsureOrderRepositoryReady() Then
        Exit Function
    End If

    effectiveDate = IIf(OrderDate = 0, Date, OrderDate)
    orderNo = GetNextSalesOrderNumber(effectiveDate)
    effectiveCustomerName = Trim$(CustomerName)
    effectiveInvoiceAddressId = InvoiceAddressId
    effectiveDeliveryAddressId = DeliveryAddressId

    If CustomerAddressId > 0 Then
        If modAddressRepository.AddressExists(CustomerAddressId) Then
            effectiveCustomerName = modAddressRepository.GetAddressDisplayName(CustomerAddressId, effectiveCustomerName)
        End If
    End If

    If effectiveInvoiceAddressId <= 0 Then
        effectiveInvoiceAddressId = CustomerAddressId
    End If
    If effectiveDeliveryAddressId <= 0 Then
        effectiveDeliveryAddressId = CustomerAddressId
    End If

    Set vatContext = GetDefaultVatContextForOrder(CustomerAddressId, effectiveDeliveryAddressId)
    effectiveVatMode = ResolveVatMode(VatMode)
    effectiveVatCode = Trim$(VatCode)
    effectiveVatRate = vatRate

    If Not vatContext Is Nothing Then
        If LenB(effectiveVatMode) = 0 Then effectiveVatMode = GetDictionaryString(vatContext, FIELD_VAT_MODE)
        If LenB(effectiveVatCode) = 0 Then effectiveVatCode = GetDictionaryString(vatContext, FIELD_VAT_CODE)
        If effectiveVatRate < 0 Then effectiveVatRate = GetDictionaryDouble(vatContext, FIELD_VAT_RATE, 0)
    End If

    If effectiveVatRate < 0 Then
        effectiveVatRate = ResolveVatRateByCode(effectiveVatCode, modVatHandler.GetVatRate())
    End If

    Set db = modDb.GetCurrentTenantDatabase()
    Set rs = db.OpenRecordset(TABLE_ORD_ORDER, dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    SetRecordsetValue rs, FIELD_ORDER_NO, orderNo
    SetRecordsetValue rs, FIELD_ORDER_TYPE_CODE, ORDER_TYPE_SALES_ORDER
    SetRecordsetValue rs, FIELD_ORDER_STATUS_CODE, ORDER_STATUS_DRAFT
    SetRecordsetValue rs, FIELD_CUSTOMER_ADDRESS_ID, CustomerAddressId
    SetRecordsetValue rs, FIELD_INVOICE_ADDRESS_ID, effectiveInvoiceAddressId
    SetRecordsetValue rs, FIELD_DELIVERY_ADDRESS_ID, effectiveDeliveryAddressId
    SetRecordsetValue rs, FIELD_CUSTOMER_NAME, effectiveCustomerName
    SetRecordsetValue rs, FIELD_ORDER_DATE, effectiveDate
    If DeliveryDate <> 0 Then SetRecordsetValue rs, FIELD_DELIVERY_DATE, DeliveryDate
    If ValidUntil <> 0 Then SetRecordsetValue rs, FIELD_VALID_UNTIL, ValidUntil
    SetRecordsetValue rs, FIELD_REFERENCE_TEXT, Trim$(ReferenceText)
    SetRecordsetValue rs, FIELD_EXTERNAL_REFERENCE, Trim$(ExternalReference)
    SetRecordsetValue rs, FIELD_LANGUAGE_CODE, ResolveLanguageCode(languageCode)
    SetRecordsetValue rs, FIELD_CURRENCY_CODE, ResolveCurrencyCode(CurrencyCode)
    SetRecordsetValue rs, FIELD_PAYMENT_TERM_CODE, Trim$(PaymentTermCode)
    SetRecordsetValue rs, FIELD_VAT_MODE, effectiveVatMode
    SetRecordsetValue rs, FIELD_VAT_CODE, effectiveVatCode
    SetRecordsetValue rs, FIELD_VAT_RATE, effectiveVatRate
    SetRecordsetValue rs, FIELD_NOTES_TEXT, NotesText
    SetRecordsetValue rs, FIELD_INTERNAL_NOTES_TEXT, InternalNotesText
    SetCreatedAuditFields rs
    SetUpdatedAuditFields rs
    rs.Update

    rs.Bookmark = rs.LastModified
    CreateSalesOrderHeader = modDaoHelper.NzLong(rs.Fields(FIELD_ORDER_ID).Value, 0)

    modLoggingHandler.LogInfo MODULE_NAME & ".CreateSalesOrderHeader", _
        "Sales order header created. order_id=" & CStr(CreateSalesOrderHeader) & _
        "; customer_address_id=" & CStr(CustomerAddressId) & _
        "; invoice_address_id=" & CStr(effectiveInvoiceAddressId) & _
        "; delivery_address_id=" & CStr(effectiveDeliveryAddressId) & _
        "; customer_name='" & Replace(effectiveCustomerName, "'", "''") & "'" & _
        "; order_type_code='" & ORDER_TYPE_SALES_ORDER & "'" & _
        "; order_status_code='" & ORDER_STATUS_DRAFT & "'" & _
        "; vat_mode='" & effectiveVatMode & "'" & _
        "; vat_code='" & Replace(effectiveVatCode, "'", "''") & "'" & _
        "; vat_rate=" & Replace(CStr(effectiveVatRate), ",", ".") & "."

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

    Set db = modDb.GetCurrentTenantDatabase()
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
    Dim headerVatContext As Object
    Dim effectiveVatCode As String

    CreateOrderLine = 0

    If OrderId <= 0 Then
        Exit Function
    End If

    If Not EnsureOrderRepositoryReady() Then
        Exit Function
    End If

    Set headerVatContext = GetOrderHeaderVatContext(OrderId)
    effectiveVatCode = Trim$(VatCode)

    If LenB(effectiveVatCode) = 0 And Not headerVatContext Is Nothing Then
        effectiveVatCode = GetDictionaryString(headerVatContext, FIELD_VAT_CODE)
    End If

    If vatRate < 0 Then
        If Not headerVatContext Is Nothing Then
            effectiveVatRate = GetDictionaryDouble(headerVatContext, FIELD_VAT_RATE, modVatHandler.GetVatRate())
        Else
            effectiveVatRate = modVatHandler.GetVatRate()
        End If
        If LenB(effectiveVatCode) > 0 Then
            effectiveVatRate = ResolveVatRateByCode(effectiveVatCode, effectiveVatRate)
        End If
    Else
        effectiveVatRate = vatRate
    End If

    Set db = modDb.GetCurrentTenantDatabase()
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
    SetRecordsetValue rs, FIELD_VAT_CODE, effectiveVatCode
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

Public Function GetNewSalesOrderDefaults(ByVal CustomerAddressId As Long) As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Dim vatContext As Object
    Dim languageCode As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    result(FIELD_CUSTOMER_ADDRESS_ID) = CustomerAddressId
    result(FIELD_INVOICE_ADDRESS_ID) = CustomerAddressId
    result(FIELD_DELIVERY_ADDRESS_ID) = CustomerAddressId
    result(FIELD_CUSTOMER_NAME) = modAddressRepository.GetAddressDisplayName(CustomerAddressId, vbNullString)
    result(FIELD_ORDER_TYPE_CODE) = ORDER_TYPE_SALES_ORDER
    result(FIELD_ORDER_STATUS_CODE) = ORDER_STATUS_DRAFT
    result(FIELD_ORDER_DATE) = Date
    result(FIELD_CURRENCY_CODE) = ResolveCurrencyCode(vbNullString)
    result(FIELD_PAYMENT_TERM_CODE) = Trim$(modTenantRepository.GetTenantParameter("PAYMENT_TERM_CODE", vbNullString))

    languageCode = ResolveAddressLanguageCode(CustomerAddressId, vbNullString)
    If LenB(languageCode) = 0 Then
        languageCode = ResolveLanguageCode(vbNullString)
    End If
    result(FIELD_LANGUAGE_CODE) = languageCode

    Set vatContext = GetDefaultVatContextForOrder(CustomerAddressId, CustomerAddressId)
    If Not vatContext Is Nothing Then
        result(FIELD_VAT_MODE) = GetDictionaryString(vatContext, FIELD_VAT_MODE)
        result(FIELD_VAT_CODE) = GetDictionaryString(vatContext, FIELD_VAT_CODE)
        result(FIELD_VAT_RATE) = GetDictionaryDouble(vatContext, FIELD_VAT_RATE, 0)
    End If

    Set GetNewSalesOrderDefaults = result
    Exit Function

ErrorHandler:
    Set GetNewSalesOrderDefaults = Nothing
    modErrorHandler.HandleError MODULE_NAME, "GetNewSalesOrderDefaults", Err
End Function

Public Function CreateTemporarySalesOrderForAddress(ByVal addressId As Long) As Long
    On Error GoTo ErrorHandler

    Dim frontendDb As DAO.Database
    Dim rs As DAO.Recordset
    Dim defaults As Object
    Dim tenantBackendPath As String

    CreateTemporarySalesOrderForAddress = 0

    If addressId <= 0 Then
        Exit Function
    End If

    If Not EnsureOrderRepositoryReady() Then
        Exit Function
    End If

    Set defaults = GetNewSalesOrderDefaults(addressId)
    If defaults Is Nothing Then
        Exit Function
    End If

    Set frontendDb = modDb.GetFrontendDatabase()
    tenantBackendPath = modDb.GetCurrentTenantBackendPath()
    modLoggingHandler.LogInfo MODULE_NAME & ".CreateTemporarySalesOrderForAddress", _
        "address_id=" & CStr(addressId) & "; tenant_backend_path=" & tenantBackendPath & "; frontend_db_name=" & SafeDatabaseName(frontendDb) & "; insert_started=True."
    Set rs = frontendDb.OpenRecordset(TABLE_TMP_ORDER, dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    SetRecordsetValue rs, FIELD_SESSION_ID, ResolveTemporarySessionId()
    SetRecordsetValue rs, FIELD_ORDER_ID, 0
    SetRecordsetValue rs, FIELD_SOURCE_ORDER_ID, Null
    SetRecordsetValue rs, FIELD_ORDER_NO, vbNullString
    SetRecordsetValue rs, FIELD_CUSTOMER_ADDRESS_ID, GetDictionaryLong(defaults, FIELD_CUSTOMER_ADDRESS_ID, 0)
    SetRecordsetValue rs, FIELD_INVOICE_ADDRESS_ID, GetDictionaryLong(defaults, FIELD_INVOICE_ADDRESS_ID, 0)
    SetRecordsetValue rs, FIELD_DELIVERY_ADDRESS_ID, GetDictionaryLong(defaults, FIELD_DELIVERY_ADDRESS_ID, 0)
    SetRecordsetValue rs, FIELD_CUSTOMER_NAME, GetDictionaryString(defaults, FIELD_CUSTOMER_NAME)
    SetRecordsetValue rs, FIELD_ORDER_TYPE_CODE, GetDictionaryString(defaults, FIELD_ORDER_TYPE_CODE)
    SetRecordsetValue rs, FIELD_ORDER_STATUS_CODE, GetDictionaryString(defaults, FIELD_ORDER_STATUS_CODE)
    SetRecordsetValue rs, FIELD_ORDER_DATE, defaults(FIELD_ORDER_DATE)
    SetRecordsetValue rs, FIELD_LANGUAGE_CODE, GetDictionaryString(defaults, FIELD_LANGUAGE_CODE)
    SetRecordsetValue rs, FIELD_CURRENCY_CODE, GetDictionaryString(defaults, FIELD_CURRENCY_CODE)
    SetRecordsetValue rs, FIELD_PAYMENT_TERM_CODE, GetDictionaryString(defaults, FIELD_PAYMENT_TERM_CODE)
    SetRecordsetValue rs, FIELD_VAT_MODE, GetDictionaryString(defaults, FIELD_VAT_MODE)
    SetRecordsetValue rs, FIELD_VAT_CODE, GetDictionaryString(defaults, FIELD_VAT_CODE)
    SetRecordsetValue rs, FIELD_VAT_RATE, GetDictionaryDouble(defaults, FIELD_VAT_RATE, 0)
    SetRecordsetValue rs, FIELD_HEADER_DISCOUNT_TYPE, "NONE"
    SetRecordsetValue rs, FIELD_HEADER_SURCHARGE_TYPE, "NONE"
    SetCreatedAuditFields rs
    SetUpdatedAuditFields rs
    rs.Update

    rs.Bookmark = rs.LastModified
    CreateTemporarySalesOrderForAddress = modDaoHelper.NzLong(rs.Fields(FIELD_TMP_ORDER_ID).Value, 0)
    If CreateTemporarySalesOrderForAddress <= 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CreateTemporarySalesOrderForAddress", _
            "Insert completed without positive tmp_order_id. address_id=" & CStr(addressId) & "; frontend_db_name=" & SafeDatabaseName(frontendDb) & "."
        Exit Function
    End If

    If Not TemporaryOrderExistsInDatabase(frontendDb, CreateTemporarySalesOrderForAddress) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CreateTemporarySalesOrderForAddress", _
            "Inserted tmp_order could not be verified in the same frontend database. tmp_order_id=" & CStr(CreateTemporarySalesOrderForAddress) & "; frontend_db_name=" & SafeDatabaseName(frontendDb) & "."
        CreateTemporarySalesOrderForAddress = 0
        Exit Function
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".CreateTemporarySalesOrderForAddress", _
        "Insert verified. tmp_order_id=" & CStr(CreateTemporarySalesOrderForAddress) & "; customer_address_id=" & CStr(GetDictionaryLong(defaults, FIELD_CUSTOMER_ADDRESS_ID, 0)) & "; source_order_id=<null>; frontend_db_name=" & SafeDatabaseName(frontendDb) & "."

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set frontendDb = Nothing
    Exit Function

ErrorHandler:
    CreateTemporarySalesOrderForAddress = 0
    modErrorHandler.HandleError MODULE_NAME, "CreateTemporarySalesOrderForAddress", Err
    Resume CleanExit
End Function

Public Function CreateTemporarySalesOrderForExistingOrder(ByVal OrderId As Long) As Long
    On Error GoTo ErrorHandler

    Dim tenantDb As DAO.Database
    Dim frontendDb As DAO.Database
    Dim rsSourceOrder As DAO.Recordset
    Dim rsSourceLines As DAO.Recordset
    Dim rsTmpOrder As DAO.Recordset
    Dim rsTmpLines As DAO.Recordset
    Dim tmpOrderId As Long

    CreateTemporarySalesOrderForExistingOrder = 0

    If OrderId <= 0 Then
        Exit Function
    End If

    If Not EnsureOrderRepositoryReady() Then
        Exit Function
    End If

    If Not OrderExists(OrderId) Then
        Exit Function
    End If

    Set tenantDb = modDb.GetCurrentTenantDatabase()
    Set frontendDb = modDb.GetFrontendDatabase()
    Set rsSourceOrder = tenantDb.OpenRecordset( _
        "SELECT * FROM [" & TABLE_ORD_ORDER & "] WHERE [" & FIELD_ORDER_ID & "]=" & CStr(OrderId) & ";", _
        dbOpenSnapshot)

    If rsSourceOrder.BOF And rsSourceOrder.EOF Then
        GoTo CleanExit
    End If

    Set rsTmpOrder = frontendDb.OpenRecordset(TABLE_TMP_ORDER, dbOpenDynaset, dbAppendOnly)
    rsTmpOrder.AddNew
    SetRecordsetValue rsTmpOrder, FIELD_SESSION_ID, ResolveTemporarySessionId()
    SetRecordsetValue rsTmpOrder, FIELD_ORDER_ID, OrderId
    SetRecordsetValue rsTmpOrder, FIELD_SOURCE_ORDER_ID, OrderId
    SetRecordsetValue rsTmpOrder, FIELD_ORDER_NO, GetRecordsetStringValue(rsSourceOrder, FIELD_ORDER_NO, vbNullString)
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_CUSTOMER_ADDRESS_ID
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_INVOICE_ADDRESS_ID
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_DELIVERY_ADDRESS_ID
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_CUSTOMER_NAME
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_ORDER_TYPE_CODE
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_ORDER_STATUS_CODE
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_ORDER_DATE
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_DELIVERY_DATE
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_VALID_UNTIL
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_REFERENCE_TEXT
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_EXTERNAL_REFERENCE
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_LANGUAGE_CODE
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_CURRENCY_CODE
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_PAYMENT_TERM_CODE
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_VAT_MODE
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_VAT_CODE
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_VAT_RATE
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_HEADER_DISCOUNT_TYPE
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_HEADER_DISCOUNT_VALUE
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_HEADER_DISCOUNT_AMOUNT
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_HEADER_SURCHARGE_TYPE
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_HEADER_SURCHARGE_VALUE
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_HEADER_SURCHARGE_AMOUNT
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_SUBTOTAL_NET_AMOUNT
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_NET_AMOUNT
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_VAT_AMOUNT
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_GROSS_AMOUNT
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_NOTES_TEXT
    CopyHeaderField rsSourceOrder, rsTmpOrder, FIELD_INTERNAL_NOTES_TEXT
    SetCreatedAuditFields rsTmpOrder
    SetUpdatedAuditFields rsTmpOrder
    rsTmpOrder.Update

    rsTmpOrder.Bookmark = rsTmpOrder.LastModified
    tmpOrderId = modDaoHelper.NzLong(rsTmpOrder.Fields(FIELD_TMP_ORDER_ID).Value, 0)
    If tmpOrderId <= 0 Then
        GoTo CleanExit
    End If

    Set rsSourceLines = tenantDb.OpenRecordset( _
        "SELECT * FROM [" & TABLE_ORD_ORDER_LINE & "] WHERE [" & FIELD_ORDER_ID & "]=" & CStr(OrderId) & _
        " ORDER BY [" & FIELD_LINE_NO & "], [" & FIELD_SORT_ORDER & "], [" & FIELD_ORDER_LINE_ID & "];", _
        dbOpenSnapshot)
    Set rsTmpLines = frontendDb.OpenRecordset(TABLE_TMP_ORDER_LINE, dbOpenDynaset, dbAppendOnly)

    If Not (rsSourceLines.BOF And rsSourceLines.EOF) Then
        rsSourceLines.MoveFirst
        Do Until rsSourceLines.EOF
            rsTmpLines.AddNew
            SetRecordsetValue rsTmpLines, FIELD_ORDER_LINE_ID, GetRecordsetLongValue(rsSourceLines, FIELD_ORDER_LINE_ID, 0)
            SetRecordsetValue rsTmpLines, FIELD_ORDER_ID, OrderId
            SetRecordsetValue rsTmpLines, FIELD_TMP_ORDER_ID, tmpOrderId
            CopyLineField rsSourceLines, rsTmpLines, FIELD_LINE_NO
            CopyLineField rsSourceLines, rsTmpLines, FIELD_ARTICLE_ID
            CopyLineField rsSourceLines, rsTmpLines, FIELD_ARTICLE_NO
            CopyLineField rsSourceLines, rsTmpLines, FIELD_LINE_TYPE_CODE
            CopyLineField rsSourceLines, rsTmpLines, FIELD_DESCRIPTION_TEXT
            CopyLineField rsSourceLines, rsTmpLines, FIELD_QUANTITY
            CopyLineField rsSourceLines, rsTmpLines, FIELD_UNIT_CODE
            CopyLineField rsSourceLines, rsTmpLines, FIELD_UNIT_PRICE
            CopyLineField rsSourceLines, rsTmpLines, FIELD_DISCOUNT_TYPE
            CopyLineField rsSourceLines, rsTmpLines, FIELD_DISCOUNT_VALUE
            CopyLineField rsSourceLines, rsTmpLines, FIELD_LINE_DISCOUNT_AMOUNT
            CopyLineField rsSourceLines, rsTmpLines, FIELD_SURCHARGE_TYPE
            CopyLineField rsSourceLines, rsTmpLines, FIELD_SURCHARGE_VALUE
            CopyLineField rsSourceLines, rsTmpLines, FIELD_LINE_SURCHARGE_AMOUNT
            CopyLineField rsSourceLines, rsTmpLines, FIELD_VAT_CODE
            CopyLineField rsSourceLines, rsTmpLines, FIELD_VAT_RATE
            CopyLineField rsSourceLines, rsTmpLines, FIELD_LINE_BASE_AMOUNT
            CopyLineField rsSourceLines, rsTmpLines, FIELD_LINE_NET_AMOUNT
            CopyLineField rsSourceLines, rsTmpLines, FIELD_LINE_VAT_AMOUNT
            CopyLineField rsSourceLines, rsTmpLines, FIELD_LINE_GROSS_AMOUNT
            CopyLineField rsSourceLines, rsTmpLines, FIELD_SORT_ORDER
            SetCreatedAuditFields rsTmpLines
            SetUpdatedAuditFields rsTmpLines
            rsTmpLines.Update
            rsSourceLines.MoveNext
        Loop
    End If

    CreateTemporarySalesOrderForExistingOrder = tmpOrderId

CleanExit:
    On Error Resume Next
    If Not rsTmpLines Is Nothing Then rsTmpLines.Close
    If Not rsTmpOrder Is Nothing Then rsTmpOrder.Close
    If Not rsSourceLines Is Nothing Then rsSourceLines.Close
    If Not rsSourceOrder Is Nothing Then rsSourceOrder.Close
    Set rsTmpLines = Nothing
    Set rsTmpOrder = Nothing
    Set rsSourceLines = Nothing
    Set rsSourceOrder = Nothing
    Set frontendDb = Nothing
    Set tenantDb = Nothing
    Exit Function

ErrorHandler:
    CreateTemporarySalesOrderForExistingOrder = 0
    modErrorHandler.HandleError MODULE_NAME, "CreateTemporarySalesOrderForExistingOrder", Err
    Resume CleanExit
End Function

Public Function TemporaryOrderExists(ByVal tmpOrderId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim frontendDb As DAO.Database

    If tmpOrderId <= 0 Then
        Exit Function
    End If

    Set frontendDb = modDb.GetFrontendDatabase()
    TemporaryOrderExists = TemporaryOrderExistsInDatabase(frontendDb, tmpOrderId)

CleanExit:
    Set frontendDb = Nothing
    Exit Function

ErrorHandler:
    TemporaryOrderExists = False
    modErrorHandler.HandleError MODULE_NAME, "TemporaryOrderExists", Err
    Resume CleanExit
End Function

Public Function DeleteTemporaryOrder(ByVal tmpOrderId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim frontendDb As DAO.Database

    DeleteTemporaryOrder = False

    If tmpOrderId <= 0 Then
        DeleteTemporaryOrder = True
        Exit Function
    End If

    Set frontendDb = modDb.GetFrontendDatabase()
    frontendDb.Execute "DELETE FROM [" & TABLE_TMP_ORDER_LINE & "] WHERE [" & FIELD_TMP_ORDER_ID & "]=" & CStr(tmpOrderId) & ";", dbFailOnError
    frontendDb.Execute "DELETE FROM [" & TABLE_TMP_ORDER & "] WHERE [" & FIELD_TMP_ORDER_ID & "]=" & CStr(tmpOrderId) & ";", dbFailOnError
    DeleteTemporaryOrder = True
    Exit Function

ErrorHandler:
    DeleteTemporaryOrder = False
    modErrorHandler.HandleError MODULE_NAME, "DeleteTemporaryOrder", Err
End Function

Public Function CommitTemporaryOrderHeader(ByVal tmpOrderId As Long) As Long
    On Error GoTo ErrorHandler

    Dim frontendDb As DAO.Database
    Dim tenantDb As DAO.Database
    Dim rsTmpOrder As DAO.Recordset
    Dim rsOrder As DAO.Recordset
    Dim OrderId As Long
    Dim sourceOrderId As Long
    Dim orderNo As String
    Dim paymentTermCode As String

    CommitTemporaryOrderHeader = 0

    If tmpOrderId <= 0 Then
        Exit Function
    End If

    If Not TemporaryOrderExists(tmpOrderId) Then
        Exit Function
    End If

    Set frontendDb = modDb.GetFrontendDatabase()
    Set tenantDb = modDb.GetCurrentTenantDatabase()
    Set rsTmpOrder = frontendDb.OpenRecordset( _
        "SELECT * FROM [" & TABLE_TMP_ORDER & "] WHERE [" & FIELD_TMP_ORDER_ID & "]=" & CStr(tmpOrderId) & ";", _
        dbOpenDynaset)

    If rsTmpOrder.BOF And rsTmpOrder.EOF Then
        GoTo CleanExit
    End If

    sourceOrderId = GetRecordsetLongValue(rsTmpOrder, FIELD_SOURCE_ORDER_ID, 0)
    paymentTermCode = ResolveEffectivePaymentTermCode(rsTmpOrder)

    If sourceOrderId > 0 Then
        Set rsOrder = tenantDb.OpenRecordset( _
            "SELECT * FROM [" & TABLE_ORD_ORDER & "] WHERE [" & FIELD_ORDER_ID & "]=" & CStr(sourceOrderId) & ";", _
            dbOpenDynaset)
        If rsOrder.BOF And rsOrder.EOF Then
            GoTo CleanExit
        End If

        orderNo = GetRecordsetStringValue(rsOrder, FIELD_ORDER_NO, vbNullString)
        rsOrder.Edit
        ApplyOrderHeaderValues rsTmpOrder, rsOrder, orderNo, paymentTermCode, False
        rsOrder.Update
        OrderId = sourceOrderId
    Else
        orderNo = GetNextSalesOrderNumber(GetRecordsetDateValue(rsTmpOrder, FIELD_ORDER_DATE, Date))
        Set rsOrder = tenantDb.OpenRecordset(TABLE_ORD_ORDER, dbOpenDynaset, dbAppendOnly)
        rsOrder.AddNew
        ApplyOrderHeaderValues rsTmpOrder, rsOrder, orderNo, paymentTermCode, True
        rsOrder.Update
        rsOrder.Bookmark = rsOrder.LastModified
        OrderId = modDaoHelper.NzLong(rsOrder.Fields(FIELD_ORDER_ID).Value, 0)
        If OrderId <= 0 Then
            GoTo CleanExit
        End If
    End If

    rsTmpOrder.Edit
    SetRecordsetValue rsTmpOrder, FIELD_ORDER_ID, OrderId
    SetRecordsetValue rsTmpOrder, FIELD_ORDER_NO, orderNo
    SetUpdatedAuditFields rsTmpOrder
    rsTmpOrder.Update

    CommitTemporaryOrderHeader = OrderId

CleanExit:
    On Error Resume Next
    If Not rsOrder Is Nothing Then rsOrder.Close
    If Not rsTmpOrder Is Nothing Then rsTmpOrder.Close
    Set rsOrder = Nothing
    Set rsTmpOrder = Nothing
    Set tenantDb = Nothing
    Set frontendDb = Nothing
    Exit Function

ErrorHandler:
    CommitTemporaryOrderHeader = 0
    modErrorHandler.HandleError MODULE_NAME, "CommitTemporaryOrderHeader", Err
    Resume CleanExit
End Function

Public Function PersistTemporaryOrder(ByVal tmpOrderId As Long) As Long
    On Error GoTo ErrorHandler

    Dim frontendDb As DAO.Database
    Dim tenantDb As DAO.Database
    Dim rsTmpOrder As DAO.Recordset
    Dim rsTmpLines As DAO.Recordset
    Dim rsOrder As DAO.Recordset
    Dim rsOrderLines As DAO.Recordset
    Dim orderNo As String
    Dim OrderId As Long
    Dim paymentTermCode As String
    Dim sortOrder As Long

    PersistTemporaryOrder = 0

    If tmpOrderId <= 0 Then
        Exit Function
    End If

    If Not TemporaryOrderExists(tmpOrderId) Then
        Exit Function
    End If

    Set frontendDb = modDb.GetFrontendDatabase()
    Set tenantDb = modDb.GetCurrentTenantDatabase()
    Set rsTmpOrder = frontendDb.OpenRecordset( _
        "SELECT * FROM [" & TABLE_TMP_ORDER & "] WHERE [" & FIELD_TMP_ORDER_ID & "]=" & CStr(tmpOrderId) & ";", _
        dbOpenDynaset)

    If rsTmpOrder.BOF And rsTmpOrder.EOF Then
        GoTo CleanExit
    End If

    orderNo = GetNextSalesOrderNumber(GetRecordsetDateValue(rsTmpOrder, FIELD_ORDER_DATE, Date))
    paymentTermCode = ResolveEffectivePaymentTermCode(rsTmpOrder)

    Set rsOrder = tenantDb.OpenRecordset(TABLE_ORD_ORDER, dbOpenDynaset, dbAppendOnly)

    rsOrder.AddNew
    ApplyOrderHeaderValues rsTmpOrder, rsOrder, orderNo, paymentTermCode, True
    rsOrder.Update

    rsOrder.Bookmark = rsOrder.LastModified
    OrderId = modDaoHelper.NzLong(rsOrder.Fields(FIELD_ORDER_ID).Value, 0)
    If OrderId <= 0 Then
        GoTo CleanExit
    End If

    Set rsTmpLines = frontendDb.OpenRecordset( _
        "SELECT * FROM [" & TABLE_TMP_ORDER_LINE & "] WHERE [" & FIELD_TMP_ORDER_ID & "]=" & CStr(tmpOrderId) & _
        " ORDER BY [" & FIELD_LINE_NO & "], [" & FIELD_SORT_ORDER & "], [" & FIELD_ORDER_LINE_ID & "];", _
        dbOpenSnapshot)
    Set rsOrderLines = tenantDb.OpenRecordset(TABLE_ORD_ORDER_LINE, dbOpenDynaset, dbAppendOnly)

    If Not (rsTmpLines.BOF And rsTmpLines.EOF) Then
        rsTmpLines.MoveFirst
        Do Until rsTmpLines.EOF
            rsOrderLines.AddNew
            sortOrder = GetRecordsetLongValue(rsTmpLines, FIELD_SORT_ORDER, 0)
            If sortOrder <= 0 Then
                sortOrder = GetRecordsetLongValue(rsTmpLines, FIELD_LINE_NO, 0)
            End If

            SetRecordsetValue rsOrderLines, FIELD_ORDER_ID, OrderId
            CopyLineField rsTmpLines, rsOrderLines, FIELD_LINE_NO
            CopyLineField rsTmpLines, rsOrderLines, FIELD_ARTICLE_ID
            CopyLineField rsTmpLines, rsOrderLines, FIELD_ARTICLE_NO
            CopyLineField rsTmpLines, rsOrderLines, FIELD_LINE_TYPE_CODE
            CopyLineField rsTmpLines, rsOrderLines, FIELD_DESCRIPTION_TEXT
            CopyLineField rsTmpLines, rsOrderLines, FIELD_QUANTITY
            CopyLineField rsTmpLines, rsOrderLines, FIELD_UNIT_CODE
            CopyLineField rsTmpLines, rsOrderLines, FIELD_UNIT_PRICE
            CopyLineField rsTmpLines, rsOrderLines, FIELD_DISCOUNT_TYPE
            CopyLineField rsTmpLines, rsOrderLines, FIELD_DISCOUNT_VALUE
            CopyLineField rsTmpLines, rsOrderLines, FIELD_LINE_DISCOUNT_AMOUNT
            CopyLineField rsTmpLines, rsOrderLines, FIELD_SURCHARGE_TYPE
            CopyLineField rsTmpLines, rsOrderLines, FIELD_SURCHARGE_VALUE
            CopyLineField rsTmpLines, rsOrderLines, FIELD_LINE_SURCHARGE_AMOUNT
            CopyLineField rsTmpLines, rsOrderLines, FIELD_VAT_CODE
            CopyLineField rsTmpLines, rsOrderLines, FIELD_VAT_RATE
            CopyLineField rsTmpLines, rsOrderLines, FIELD_LINE_BASE_AMOUNT
            CopyLineField rsTmpLines, rsOrderLines, FIELD_LINE_NET_AMOUNT
            CopyLineField rsTmpLines, rsOrderLines, FIELD_LINE_VAT_AMOUNT
            CopyLineField rsTmpLines, rsOrderLines, FIELD_LINE_GROSS_AMOUNT
            SetRecordsetValue rsOrderLines, FIELD_SORT_ORDER, sortOrder
            SetCreatedAuditFields rsOrderLines
            SetUpdatedAuditFields rsOrderLines
            rsOrderLines.Update
            rsTmpLines.MoveNext
        Loop
    End If

    Call modOrderCalculationService.RecalculateOrder(OrderId)
    Call DeleteTemporaryOrder(tmpOrderId)
    PersistTemporaryOrder = OrderId

CleanExit:
    On Error Resume Next
    If Not rsOrderLines Is Nothing Then rsOrderLines.Close
    If Not rsTmpLines Is Nothing Then rsTmpLines.Close
    If Not rsOrder Is Nothing Then rsOrder.Close
    If Not rsTmpOrder Is Nothing Then rsTmpOrder.Close
    Set rsOrderLines = Nothing
    Set rsTmpLines = Nothing
    Set rsOrder = Nothing
    Set rsTmpOrder = Nothing
    Set tenantDb = Nothing
    Set frontendDb = Nothing
    Exit Function

ErrorHandler:
    PersistTemporaryOrder = 0
    modErrorHandler.HandleError MODULE_NAME, "PersistTemporaryOrder", Err
    Resume CleanExit
End Function

Public Function GetDefaultVatContextForOrder( _
    Optional ByVal CustomerAddressId As Long = 0, _
    Optional ByVal DeliveryAddressId As Long = 0) As Object
    On Error GoTo ErrorHandler

    Dim result As Object
    Dim tenantCountryCode As String
    Dim deliveryCountryCode As String
    Dim resolvedVatMode As String
    Dim resolvedVatCode As String
    Dim resolvedVatRate As Double
    Dim effectiveAddressId As Long

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    tenantCountryCode = ResolveTenantCountryCode()
    effectiveAddressId = DeliveryAddressId
    If effectiveAddressId <= 0 Then
        effectiveAddressId = CustomerAddressId
    End If

    deliveryCountryCode = modAddressRepository.GetAddressCountryCode(effectiveAddressId, tenantCountryCode)
    If LenB(deliveryCountryCode) = 0 Then
        deliveryCountryCode = tenantCountryCode
    End If

    If StrComp(deliveryCountryCode, tenantCountryCode, vbTextCompare) <> 0 Then
        resolvedVatMode = "NONE"
        resolvedVatCode = ResolveDefaultZeroVatCode(tenantCountryCode)
        resolvedVatRate = 0
    Else
        resolvedVatMode = ResolveVatMode(modTenantRepository.GetTenantParameter("VAT_MODE", modVatHandler.GetVatMode()))
        resolvedVatCode = ResolveDefaultStandardVatCode(tenantCountryCode)
        resolvedVatRate = ResolveVatRateByCode(resolvedVatCode, modVatHandler.GetVatRate())
    End If

    result("tenant_country_code") = tenantCountryCode
    result("country_code") = deliveryCountryCode
    result(FIELD_VAT_MODE) = resolvedVatMode
    result(FIELD_VAT_CODE) = resolvedVatCode
    result(FIELD_VAT_RATE) = resolvedVatRate

    Set GetDefaultVatContextForOrder = result
    Exit Function

ErrorHandler:
    Set GetDefaultVatContextForOrder = Nothing
    modErrorHandler.HandleError MODULE_NAME, "GetDefaultVatContextForOrder", Err
End Function

Public Function GetOrderHeaderVatContext(ByVal OrderId As Long) As Object
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim result As Object
    Dim deliveryAddressId As Long
    Dim customerAddressId As Long

    If OrderId <= 0 Then
        Exit Function
    End If

    Set db = modDb.GetCurrentTenantDatabase()
    Set rs = db.OpenRecordset( _
        "SELECT [" & FIELD_CUSTOMER_ADDRESS_ID & "], [" & FIELD_DELIVERY_ADDRESS_ID & "], " & _
        "[" & FIELD_VAT_MODE & "], [" & FIELD_VAT_CODE & "], [" & FIELD_VAT_RATE & "] " & _
        "FROM [" & TABLE_ORD_ORDER & "] WHERE [" & FIELD_ORDER_ID & "]=" & CStr(OrderId) & ";", _
        dbOpenSnapshot)

    If rs.BOF And rs.EOF Then
        GoTo CleanExit
    End If

    customerAddressId = GetRecordsetLongValue(rs, FIELD_CUSTOMER_ADDRESS_ID, 0)
    deliveryAddressId = GetRecordsetLongValue(rs, FIELD_DELIVERY_ADDRESS_ID, 0)

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    result("country_code") = ResolveOrderCountryCode(customerAddressId, deliveryAddressId)
    result(FIELD_VAT_MODE) = GetRecordsetStringValue(rs, FIELD_VAT_MODE, modVatHandler.GetVatMode())
    result(FIELD_VAT_CODE) = GetRecordsetStringValue(rs, FIELD_VAT_CODE, vbNullString)
    result(FIELD_VAT_RATE) = GetRecordsetDoubleValue(rs, FIELD_VAT_RATE, 0)

    Set GetOrderHeaderVatContext = result

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    Set GetOrderHeaderVatContext = Nothing
    modErrorHandler.HandleError MODULE_NAME, "GetOrderHeaderVatContext", Err
    Resume CleanExit
End Function

Public Function ApplyDefaultVatContextToOrder(ByVal OrderId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim customerAddressId As Long
    Dim deliveryAddressId As Long
    Dim vatContext As Object

    ApplyDefaultVatContextToOrder = False

    If OrderId <= 0 Then
        Exit Function
    End If

    Set db = modDb.GetCurrentTenantDatabase()
    Set rs = db.OpenRecordset( _
        "SELECT * FROM [" & TABLE_ORD_ORDER & "] WHERE [" & FIELD_ORDER_ID & "]=" & CStr(OrderId) & ";", _
        dbOpenDynaset)

    If rs.BOF And rs.EOF Then
        GoTo CleanExit
    End If

    customerAddressId = GetRecordsetLongValue(rs, FIELD_CUSTOMER_ADDRESS_ID, 0)
    deliveryAddressId = GetRecordsetLongValue(rs, FIELD_DELIVERY_ADDRESS_ID, 0)
    Set vatContext = GetDefaultVatContextForOrder(customerAddressId, deliveryAddressId)
    If vatContext Is Nothing Then
        GoTo CleanExit
    End If

    rs.Edit
    SetRecordsetValue rs, FIELD_VAT_MODE, GetDictionaryString(vatContext, FIELD_VAT_MODE)
    SetRecordsetValue rs, FIELD_VAT_CODE, GetDictionaryString(vatContext, FIELD_VAT_CODE)
    SetRecordsetValue rs, FIELD_VAT_RATE, GetDictionaryDouble(vatContext, FIELD_VAT_RATE, 0)
    SetUpdatedAuditFields rs
    rs.Update

    ApplyDefaultVatContextToOrder = True

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    ApplyDefaultVatContextToOrder = False
    modErrorHandler.HandleError MODULE_NAME, "ApplyDefaultVatContextToOrder", Err
    Resume CleanExit
End Function

Public Function GetAddressRowSourceSql() As String
    GetAddressRowSourceSql = _
        "SELECT a.address_id, " & _
        "IIf(Nz(a.company_name,'')<>'', Nz(a.company_name,''), Trim(Nz(a.first_name,'') & ' ' & Nz(a.last_name,''))) AS display_name, " & _
        "Nz(a.city,'') AS city_name, " & _
        "Nz(a.country_code,'') AS country_code " & _
        "FROM adr_address AS a " & _
        "WHERE Nz(a.is_active, True)=True " & _
        "ORDER BY IIf(Nz(a.company_name,'')<>'', Nz(a.company_name,''), Trim(Nz(a.first_name,'') & ' ' & Nz(a.last_name,''))), Nz(a.city,''), a.address_id;"
End Function

Public Function GetPaymentTermRowSourceSql(Optional ByVal languageCode As String = "") As String
    Dim effectiveLanguageCode As String

    effectiveLanguageCode = Trim$(languageCode)
    If LenB(effectiveLanguageCode) = 0 Then
        effectiveLanguageCode = modFwTranslationRuntime.GetCurrentLanguageCode()
    End If

    GetPaymentTermRowSourceSql = _
        "SELECT payment_term_code, " & _
        "ResolveText('PAYMENT_TERM.' & Nz([payment_term_code],'') & '.TITLE', Nz([payment_term_code],''), " & _
        SqlText(effectiveLanguageCode) & ") AS payment_term_title " & _
        "FROM ten_payment_term " & _
        "WHERE Len(Trim(Nz([payment_term_code], ''))) > 0 " & _
        "AND Nz([is_active], True)=True " & _
        "ORDER BY Nz([sort_order], 0), Nz([payment_term_code], '');"
End Function

Public Function GetCurrencyRowSourceSql() As String
    GetCurrencyRowSourceSql = _
        "SELECT currency_code " & _
        "FROM ref_currency " & _
        "WHERE Nz(is_active, True)=True " & _
        "ORDER BY currency_code;"
End Function

Public Function GetLanguageRowSourceSql() As String
    GetLanguageRowSourceSql = _
        "SELECT language_code, language_name " & _
        "FROM ref_language " & _
        "WHERE Nz(is_active, True)=True " & _
        "ORDER BY Nz(sort_order, 0), Nz(language_name, '');"
End Function

Public Function GetVatCodeRowSourceSql() As String
    GetVatCodeRowSourceSql = _
        "SELECT v.vat_code, " & _
        "Nz(v.vat_code,'') & ' (' & Format(Nz(v.vat_rate,0), '0.0') & '%)' AS vat_display_name, " & _
        "Nz(v.vat_rate,0) AS vat_rate " & _
        "FROM ref_vat_code AS v " & _
        "WHERE Nz(v.is_active, True)=True " & _
        "ORDER BY Nz(v.sort_order, 0), Nz(v.vat_code,'');"
End Function

Public Function ResolveVatRateByCode(ByVal VatCode As String, Optional ByVal defaultValue As Double = 0) As Double
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    VatCode = UCase$(Trim$(VatCode))
    ResolveVatRateByCode = defaultValue

    If LenB(VatCode) = 0 Then
        Exit Function
    End If

    Set db = modDb.GetCurrentTenantDatabase()
    Set rs = db.OpenRecordset( _
        "SELECT [vat_rate] FROM [ref_vat_code] WHERE [vat_code]=" & SqlText(VatCode) & ";", _
        dbOpenSnapshot)

    If Not (rs.BOF And rs.EOF) Then
        ResolveVatRateByCode = GetRecordsetDoubleValue(rs, FIELD_VAT_RATE, defaultValue)
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    ResolveVatRateByCode = defaultValue
    modErrorHandler.HandleError MODULE_NAME, "ResolveVatRateByCode", Err
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

    Set db = modDb.GetCurrentTenantDatabase()
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

Private Sub CopyHeaderField(ByVal rsSource As DAO.Recordset, ByVal rsTarget As DAO.Recordset, ByVal fieldName As String)
    If modDaoHelper.RecordsetHasField(rsSource, fieldName) Then
        SetRecordsetValue rsTarget, fieldName, rsSource.Fields(fieldName).Value
    End If
End Sub

Private Sub ApplyOrderHeaderValues( _
    ByVal rsTmpOrder As DAO.Recordset, _
    ByVal rsOrder As DAO.Recordset, _
    ByVal orderNo As String, _
    ByVal paymentTermCode As String, _
    ByVal isNewRecord As Boolean)

    If LenB(Trim$(orderNo)) > 0 Then
        SetRecordsetValue rsOrder, FIELD_ORDER_NO, Trim$(orderNo)
    End If

    CopyHeaderField rsTmpOrder, rsOrder, FIELD_ORDER_TYPE_CODE
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_ORDER_STATUS_CODE
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_CUSTOMER_ADDRESS_ID
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_INVOICE_ADDRESS_ID
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_DELIVERY_ADDRESS_ID
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_CUSTOMER_NAME
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_ORDER_DATE
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_DELIVERY_DATE
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_VALID_UNTIL
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_REFERENCE_TEXT
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_EXTERNAL_REFERENCE
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_LANGUAGE_CODE
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_CURRENCY_CODE
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_PAYMENT_TERM_CODE
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_VAT_MODE
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_VAT_CODE
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_VAT_RATE
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_HEADER_DISCOUNT_TYPE
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_HEADER_DISCOUNT_VALUE
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_HEADER_DISCOUNT_AMOUNT
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_HEADER_SURCHARGE_TYPE
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_HEADER_SURCHARGE_VALUE
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_HEADER_SURCHARGE_AMOUNT
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_SUBTOTAL_NET_AMOUNT
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_NET_AMOUNT
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_VAT_AMOUNT
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_GROSS_AMOUNT
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_NOTES_TEXT
    CopyHeaderField rsTmpOrder, rsOrder, FIELD_INTERNAL_NOTES_TEXT
    SetRecordsetValue rsOrder, FIELD_PAYMENT_TERM_CODE, paymentTermCode

    If isNewRecord Then
        SetCreatedAuditFields rsOrder
    End If
    SetUpdatedAuditFields rsOrder
End Sub

Private Sub CopyLineField(ByVal rsSource As DAO.Recordset, ByVal rsTarget As DAO.Recordset, ByVal fieldName As String)
    If modDaoHelper.RecordsetHasField(rsSource, fieldName) Then
        SetRecordsetValue rsTarget, fieldName, rsSource.Fields(fieldName).Value
    End If
End Sub

Private Function ResolveCurrencyCode(ByVal explicitValue As String) As String
    ResolveCurrencyCode = Trim$(explicitValue)
    If LenB(ResolveCurrencyCode) = 0 Then
        ResolveCurrencyCode = modTenantRepository.GetTenantParameter("CURRENCY_CODE", "CHF")
    End If
End Function

Private Function ResolveEffectivePaymentTermCode(ByVal rsTmpOrder As DAO.Recordset) As String
    ResolveEffectivePaymentTermCode = Trim$(GetRecordsetStringValue(rsTmpOrder, FIELD_PAYMENT_TERM_CODE, vbNullString))
    If LenB(ResolveEffectivePaymentTermCode) = 0 Then
        ResolveEffectivePaymentTermCode = Trim$(modTenantRepository.GetTenantParameter("PAYMENT_TERM_CODE", vbNullString))
    End If
End Function

Private Function ResolveLanguageCode(ByVal explicitValue As String) As String
    ResolveLanguageCode = Trim$(explicitValue)
    If LenB(ResolveLanguageCode) = 0 Then
        ResolveLanguageCode = modFwTranslationRuntime.GetCurrentLanguageCode()
    End If
End Function

Private Function ResolveTemporarySessionId() As String
    ResolveTemporarySessionId = Trim$(modSessionContext.currentUserId)
    If LenB(ResolveTemporarySessionId) = 0 Then
        ResolveTemporarySessionId = Trim$(modSessionContext.CurrentUserName)
    End If
    If LenB(ResolveTemporarySessionId) = 0 Then
        ResolveTemporarySessionId = "SYSTEM"
    End If
End Function

Private Function ResolveAddressLanguageCode(ByVal addressId As Long, ByVal defaultValue As String) As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ResolveAddressLanguageCode = Trim$(defaultValue)
    If addressId <= 0 Then
        Exit Function
    End If

    Set db = modDb.GetCurrentTenantDatabase()
    Set rs = db.OpenRecordset( _
        "SELECT [" & FIELD_LANGUAGE_CODE & "] FROM [adr_address] WHERE [address_id]=" & CStr(addressId) & ";", _
        dbOpenSnapshot)

    If Not (rs.BOF And rs.EOF) Then
        ResolveAddressLanguageCode = GetRecordsetStringValue(rs, FIELD_LANGUAGE_CODE, ResolveAddressLanguageCode)
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    ResolveAddressLanguageCode = Trim$(defaultValue)
    modErrorHandler.HandleError MODULE_NAME, "ResolveAddressLanguageCode", Err
    Resume CleanExit
End Function

Private Function ResolveVatMode(ByVal explicitValue As String) As String
    If LenB(Trim$(explicitValue)) = 0 Then
        ResolveVatMode = modVatHandler.GetVatMode()
    Else
        ResolveVatMode = modVatHandler.NormalizeVatMode(explicitValue)
    End If
End Function

Private Function ResolveTenantCountryCode() As String
    ResolveTenantCountryCode = UCase$(Trim$( _
        modTenantRepository.GetTenantParameter( _
            "TENANT_COUNTRY_CODE", _
            modTenantRepository.GetTenantParameter("SENDER_COUNTRY_CODE", DEFAULT_TENANT_COUNTRY_CODE))))
    If LenB(ResolveTenantCountryCode) = 0 Then
        ResolveTenantCountryCode = DEFAULT_TENANT_COUNTRY_CODE
    End If
End Function

Private Function ResolveDefaultZeroVatCode(ByVal tenantCountryCode As String) As String
    Dim candidateCode As String

    candidateCode = UCase$(Trim$(tenantCountryCode)) & "_ZERO"
    If VatCodeExists(candidateCode) Then
        ResolveDefaultZeroVatCode = candidateCode
    ElseIf VatCodeExists(DEFAULT_ZERO_VAT_CODE) Then
        ResolveDefaultZeroVatCode = DEFAULT_ZERO_VAT_CODE
    Else
        ResolveDefaultZeroVatCode = FindVatCodeByCountryAndRate(tenantCountryCode, True, vbNullString)
    End If
End Function

Private Function ResolveDefaultStandardVatCode(ByVal tenantCountryCode As String) As String
    Dim candidateCode As String

    candidateCode = UCase$(Trim$(tenantCountryCode)) & "_STANDARD"
    If VatCodeExists(candidateCode) Then
        ResolveDefaultStandardVatCode = candidateCode
    ElseIf VatCodeExists(DEFAULT_STANDARD_VAT_CODE) Then
        ResolveDefaultStandardVatCode = DEFAULT_STANDARD_VAT_CODE
    Else
        ResolveDefaultStandardVatCode = FindVatCodeByCountryAndRate(tenantCountryCode, False, DEFAULT_ZERO_VAT_CODE)
    End If
End Function

Private Function ResolveOrderCountryCode(ByVal customerAddressId As Long, ByVal deliveryAddressId As Long) As String
    Dim effectiveAddressId As Long

    effectiveAddressId = deliveryAddressId
    If effectiveAddressId <= 0 Then
        effectiveAddressId = customerAddressId
    End If

    ResolveOrderCountryCode = modAddressRepository.GetAddressCountryCode(effectiveAddressId, ResolveTenantCountryCode())
    If LenB(ResolveOrderCountryCode) = 0 Then
        ResolveOrderCountryCode = ResolveTenantCountryCode()
    End If
End Function

Private Function VatCodeExists(ByVal VatCode As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    VatCode = UCase$(Trim$(VatCode))
    If LenB(VatCode) = 0 Then
        Exit Function
    End If

    Set db = modDb.GetCurrentTenantDatabase()
    Set rs = db.OpenRecordset( _
        "SELECT [vat_code] FROM [ref_vat_code] WHERE [vat_code]=" & SqlText(VatCode) & ";", _
        dbOpenSnapshot)
    VatCodeExists = Not (rs.BOF And rs.EOF)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    VatCodeExists = False
    Resume CleanExit
End Function

Private Function FindVatCodeByCountryAndRate( _
    ByVal countryCode As String, _
    ByVal zeroRateOnly As Boolean, _
    Optional ByVal fallbackValue As String = "") As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlStatement As String

    countryCode = UCase$(Trim$(countryCode))
    If LenB(countryCode) = 0 Then
        FindVatCodeByCountryAndRate = fallbackValue
        Exit Function
    End If

    sqlStatement = _
        "SELECT TOP 1 [vat_code] FROM [ref_vat_code] " & _
        "WHERE Nz([is_active], True)=True " & _
        "AND UCase(Nz([country_code],''))=" & SqlText(countryCode) & " "

    If zeroRateOnly Then
        sqlStatement = sqlStatement & "AND Nz([vat_rate],0)=0 "
    Else
        sqlStatement = sqlStatement & "AND Nz([vat_rate],0)>0 "
    End If

    sqlStatement = sqlStatement & "ORDER BY Nz([sort_order],0), [vat_code];"

    Set db = modDb.GetCurrentTenantDatabase()
    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)

    If Not (rs.BOF And rs.EOF) Then
        FindVatCodeByCountryAndRate = GetRecordsetStringValue(rs, FIELD_VAT_CODE, fallbackValue)
    Else
        FindVatCodeByCountryAndRate = fallbackValue
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    FindVatCodeByCountryAndRate = fallbackValue
    Resume CleanExit
End Function


Private Function SqlText(ByVal valueText As String) As String
    SqlText = "'" & Replace(Trim$(valueText), "'", "''") & "'"
End Function

Private Function GetRecordsetStringValue(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal defaultValue As String) As String
    If modDaoHelper.RecordsetHasField(rs, fieldName) Then
        GetRecordsetStringValue = Trim$(modDaoHelper.NzString(rs.Fields(fieldName).Value, defaultValue))
    Else
        GetRecordsetStringValue = defaultValue
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

Private Function GetRecordsetDateValue(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal defaultValue As Date) As Date
    If modDaoHelper.RecordsetHasField(rs, fieldName) Then
        If IsDate(rs.Fields(fieldName).Value) Then
            GetRecordsetDateValue = CDate(rs.Fields(fieldName).Value)
        Else
            GetRecordsetDateValue = defaultValue
        End If
    Else
        GetRecordsetDateValue = defaultValue
    End If
End Function

Private Function GetRecordsetLongValue(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal defaultValue As Long) As Long
    If modDaoHelper.RecordsetHasField(rs, fieldName) Then
        GetRecordsetLongValue = modDaoHelper.NzLong(rs.Fields(fieldName).Value, defaultValue)
    Else
        GetRecordsetLongValue = defaultValue
    End If
End Function

Private Function GetDictionaryString(ByVal dictionaryObject As Object, ByVal keyName As String) As String
    On Error GoTo SafeExit

    If dictionaryObject Is Nothing Then
        Exit Function
    End If

    If dictionaryObject.Exists(keyName) Then
        GetDictionaryString = Trim$(modDaoHelper.NzString(dictionaryObject(keyName), vbNullString))
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
        rawValue = modDaoHelper.NzString(dictionaryObject(keyName), CStr(defaultValue))
        If IsNumeric(rawValue) Then
            GetDictionaryDouble = CDbl(rawValue)
        End If
    End If

SafeExit:
End Function

Private Function GetDictionaryLong(ByVal dictionaryObject As Object, ByVal keyName As String, ByVal defaultValue As Long) As Long
    On Error GoTo SafeExit

    GetDictionaryLong = defaultValue
    If dictionaryObject Is Nothing Then
        Exit Function
    End If

    If dictionaryObject.Exists(keyName) Then
        GetDictionaryLong = modDaoHelper.NzLong(dictionaryObject(keyName), defaultValue)
    End If

SafeExit:
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

Private Function TemporaryOrderExistsInDatabase(ByVal db As DAO.Database, ByVal tmpOrderId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset

    If db Is Nothing Then
        Exit Function
    End If

    If tmpOrderId <= 0 Then
        Exit Function
    End If

    Set rs = db.OpenRecordset( _
        "SELECT [" & FIELD_TMP_ORDER_ID & "] FROM [" & TABLE_TMP_ORDER & "] WHERE [" & FIELD_TMP_ORDER_ID & "]=" & CStr(tmpOrderId) & ";", _
        dbOpenSnapshot)
    TemporaryOrderExistsInDatabase = Not (rs.BOF And rs.EOF)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    TemporaryOrderExistsInDatabase = False
    modErrorHandler.HandleError MODULE_NAME, "TemporaryOrderExistsInDatabase", Err
    Resume CleanExit
End Function

Private Function SafeDatabaseName(ByVal db As DAO.Database) As String
    On Error GoTo SafeExit

    If db Is Nothing Then
        SafeDatabaseName = "<nothing>"
    Else
        SafeDatabaseName = Trim$(db.Name)
        If LenB(SafeDatabaseName) = 0 Then
            SafeDatabaseName = "<empty>"
        End If
    End If

SafeExit:
    If LenB(SafeDatabaseName) = 0 Then
        SafeDatabaseName = "<unavailable>"
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



