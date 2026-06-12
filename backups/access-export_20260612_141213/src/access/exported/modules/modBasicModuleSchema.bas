Attribute VB_Name = "modBasicModuleSchema"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modBasicModuleSchema
' Purpose   : Creates BasicModule v1 tables for articles, orders, and tenant references.
' Author    : Codex
' Version   : 0.2.0
'===============================================================================

Private Const MODULE_NAME As String = "modBasicModuleSchema"
Private Const ACCESS_CONNECT_PREFIX As String = ";DATABASE="

Public Sub CreateBasicModuleTables(Optional ByVal backendPath As String = vbNullString)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim shouldCloseDb As Boolean

    If Not OpenSchemaDatabase(backendPath, db, shouldCloseDb) Then
        Err.Raise vbObjectError + 6210, MODULE_NAME & ".CreateBasicModuleTables", "Schema target database could not be resolved."
    End If

    CreateTenPaymentTerms db
    CreateRefVatCodes db
    CreateRefUnits db
    CreateRefArticleTypeCodes db
    CreateRefAddressType db
    CreateRefSalutation db
    CreateRefAddressingMode db
    CreateRefContactType db

    CreateTblProductGroups db
    CreateTblArticles db

    CreateTblOrders db
    CreateTblOrderLines db
    Call EnsureOrderPhase1SchemaForDatabase(db)
    Call modOrderRepository.EnsureSalesOrderNumberRange(Year(Date))

    MsgBox "BasicModule-Tabellen wurden erstellt.", vbInformation, MODULE_NAME

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then
        If shouldCloseDb Then db.Close
    End If
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Erstellen der BasicModule-Tabellen: " & Err.description, vbExclamation, MODULE_NAME
    Resume CleanExit
End Sub

Public Function EnsureOrderPhase1Schema(Optional ByVal backendPath As String = vbNullString) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim shouldCloseDb As Boolean

    If Not OpenSchemaDatabase(backendPath, db, shouldCloseDb) Then
        Exit Function
    End If

    EnsureOrderPhase1Schema = EnsureOrderPhase1SchemaForDatabase(db)

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then
        If shouldCloseDb Then db.Close
    End If
    Set db = Nothing
    Exit Function

ErrorHandler:
    EnsureOrderPhase1Schema = False
    Resume CleanExit
End Function

Public Sub DiagnoseOrderSchema(Optional ByVal backendPath As String = vbNullString)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim shouldCloseDb As Boolean
    Dim resolvedPath As String

    If Not OpenSchemaDatabase(backendPath, db, shouldCloseDb) Then
        modLoggingHandler.LogError MODULE_NAME & ".DiagnoseOrderSchema", "Schema target database could not be resolved."
        GoTo CleanExit
    End If

    resolvedPath = ResolveDatabasePath(db)
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "backend_path=" & resolvedPath
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "table_exists ord_order=" & CStr(TableExists(db, "ord_order"))
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "field_exists ord_order.customer_address_id=" & CStr(FieldExists(db, "ord_order", "customer_address_id"))
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "table_exists ord_order_line=" & CStr(TableExists(db, "ord_order_line"))
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "field_exists ord_order_line.article_no=" & CStr(FieldExists(db, "ord_order_line", "article_no"))
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "field_exists ord_order_line.vat_rate=" & CStr(FieldExists(db, "ord_order_line", "vat_rate"))
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "index_exists ix_ord_order_customer_address_id=" & CStr(IndexExists(db, "ord_order", "ix_ord_order_customer_address_id"))

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then
        If shouldCloseDb Then db.Close
    End If
    Set db = Nothing
    Exit Sub

ErrorHandler:
    modLoggingHandler.LogError MODULE_NAME & ".DiagnoseOrderSchema", Err.description, Err.Number
    Resume CleanExit
End Sub

Private Function OpenSchemaDatabase( _
    ByVal explicitBackendPath As String, _
    ByRef resolvedDb As DAO.Database, _
    ByRef shouldCloseDb As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim frontendDb As DAO.Database
    Dim targetBackendPath As String

    Set frontendDb = currentDb
    targetBackendPath = Trim$(explicitBackendPath)
    modLoggingHandler.LogInfo MODULE_NAME & ".OpenSchemaDatabase", "frontend_db=" & ResolveDatabasePath(frontendDb)

    If LenB(targetBackendPath) = 0 Then
        If TableExists(frontendDb, "ord_order") Then
            If IsLinkedAccessTable(frontendDb, "ord_order") Then
                targetBackendPath = ResolveLinkedTableBackendPath(frontendDb, "ord_order")
            End If
        Else
            targetBackendPath = Trim$(modDb.GetBackendPath())
        End If
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".OpenSchemaDatabase", "requested_backend_path=" & targetBackendPath

    If LenB(targetBackendPath) > 0 Then
        If StrComp(NormalizePath(frontendDb.Name), NormalizePath(targetBackendPath), vbTextCompare) <> 0 Then
            Set resolvedDb = DBEngine.OpenDatabase(targetBackendPath)
            shouldCloseDb = True
        Else
            Set resolvedDb = frontendDb
        End If
    Else
        Set resolvedDb = frontendDb
    End If

    If Not resolvedDb Is Nothing Then
        modLoggingHandler.LogInfo MODULE_NAME & ".OpenSchemaDatabase", "resolved_schema_db=" & ResolveDatabasePath(resolvedDb)
    End If

    OpenSchemaDatabase = Not (resolvedDb Is Nothing)
    Exit Function

ErrorHandler:
    modLoggingHandler.LogError MODULE_NAME & ".OpenSchemaDatabase", Err.description, Err.Number
    Set resolvedDb = Nothing
    shouldCloseDb = False
    OpenSchemaDatabase = False
End Function

Private Sub CreateRefAddressType(ByVal db As DAO.Database)
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "CREATE TABLE ref_address_type ("
    sqlStatement = sqlStatement & "address_type_code TEXT(50) CONSTRAINT pk_ref_address_type PRIMARY KEY, "
    sqlStatement = sqlStatement & "translation_key TEXT(100), "
    sqlStatement = sqlStatement & "sort_order LONG, "
    sqlStatement = sqlStatement & "is_active YESNO, "
    sqlStatement = sqlStatement & "created_at DATETIME, "
    sqlStatement = sqlStatement & "created_by TEXT(50), "
    sqlStatement = sqlStatement & "updated_at DATETIME, "
    sqlStatement = sqlStatement & "updated_by TEXT(50)"
    sqlStatement = sqlStatement & ");"

    ExecuteDdl db, sqlStatement
End Sub

Private Sub CreateRefSalutation(ByVal db As DAO.Database)
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "CREATE TABLE ref_salutation ("
    sqlStatement = sqlStatement & "salutation_code TEXT(30) CONSTRAINT pk_ref_salutation PRIMARY KEY, "
    sqlStatement = sqlStatement & "translation_key TEXT(100), "
    sqlStatement = sqlStatement & "sort_order LONG, "
    sqlStatement = sqlStatement & "is_active YESNO, "
    sqlStatement = sqlStatement & "created_at DATETIME, "
    sqlStatement = sqlStatement & "created_by TEXT(50), "
    sqlStatement = sqlStatement & "updated_at DATETIME, "
    sqlStatement = sqlStatement & "updated_by TEXT(50)"
    sqlStatement = sqlStatement & ");"

    ExecuteDdl db, sqlStatement
End Sub

Private Sub CreateRefAddressingMode(ByVal db As DAO.Database)
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "CREATE TABLE ref_addressing_mode ("
    sqlStatement = sqlStatement & "addressing_mode_code TEXT(30) CONSTRAINT pk_ref_addressing_mode PRIMARY KEY, "
    sqlStatement = sqlStatement & "translation_key TEXT(100), "
    sqlStatement = sqlStatement & "sort_order LONG, "
    sqlStatement = sqlStatement & "is_active YESNO, "
    sqlStatement = sqlStatement & "created_at DATETIME, "
    sqlStatement = sqlStatement & "created_by TEXT(50), "
    sqlStatement = sqlStatement & "updated_at DATETIME, "
    sqlStatement = sqlStatement & "updated_by TEXT(50)"
    sqlStatement = sqlStatement & ");"

    ExecuteDdl db, sqlStatement
End Sub

Private Sub CreateRefContactType(ByVal db As DAO.Database)
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "CREATE TABLE ref_contact_type ("
    sqlStatement = sqlStatement & "contact_type_code TEXT(30) CONSTRAINT pk_ref_contact_type PRIMARY KEY, "
    sqlStatement = sqlStatement & "translation_key TEXT(100), "
    sqlStatement = sqlStatement & "sort_order LONG, "
    sqlStatement = sqlStatement & "is_active YESNO, "
    sqlStatement = sqlStatement & "created_at DATETIME, "
    sqlStatement = sqlStatement & "created_by TEXT(50), "
    sqlStatement = sqlStatement & "updated_at DATETIME, "
    sqlStatement = sqlStatement & "updated_by TEXT(50)"
    sqlStatement = sqlStatement & ");"

    ExecuteDdl db, sqlStatement
End Sub

Private Sub CreateTenPaymentTerms(ByVal db As DAO.Database)
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "CREATE TABLE ten_payment_term ("
    sqlStatement = sqlStatement & "payment_term_id AUTOINCREMENT CONSTRAINT pk_ten_payment_term PRIMARY KEY, "
    sqlStatement = sqlStatement & "payment_term_code TEXT(50) NOT NULL, "
    sqlStatement = sqlStatement & "language_code TEXT(10) NOT NULL, "
    sqlStatement = sqlStatement & "title TEXT(100), "
    sqlStatement = sqlStatement & "terms_text LONGTEXT, "
    sqlStatement = sqlStatement & "days_net LONG, "
    sqlStatement = sqlStatement & "discount_days LONG, "
    sqlStatement = sqlStatement & "discount_percent DOUBLE, "
    sqlStatement = sqlStatement & "is_default YESNO, "
    sqlStatement = sqlStatement & "is_active YESNO, "
    sqlStatement = sqlStatement & "sort_order LONG, "
    sqlStatement = sqlStatement & "created_at DATETIME, "
    sqlStatement = sqlStatement & "created_by TEXT(50), "
    sqlStatement = sqlStatement & "updated_at DATETIME, "
    sqlStatement = sqlStatement & "updated_by TEXT(50)"
    sqlStatement = sqlStatement & ");"

    ExecuteDdl db, sqlStatement
    ExecuteCreateIndexIfMissing db, "ten_payment_term", "ux_ten_payment_term_code_language", _
        "CREATE UNIQUE INDEX ux_ten_payment_term_code_language ON ten_payment_term (payment_term_code, language_code);"
    ExecuteCreateIndexIfMissing db, "ten_payment_term", "ix_ten_payment_term_is_default", _
        "CREATE INDEX ix_ten_payment_term_is_default ON ten_payment_term (is_default);"
    ExecuteCreateIndexIfMissing db, "ten_payment_term", "ix_ten_payment_term_is_active", _
        "CREATE INDEX ix_ten_payment_term_is_active ON ten_payment_term (is_active);"
End Sub

Private Sub CreateRefVatCodes(ByVal db As DAO.Database)
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "CREATE TABLE ref_vat_code ("
    sqlStatement = sqlStatement & "vat_code TEXT(30) CONSTRAINT pk_ref_vat_code PRIMARY KEY, "
    sqlStatement = sqlStatement & "translation_key TEXT(100), "
    sqlStatement = sqlStatement & "vat_rate DOUBLE, "
    sqlStatement = sqlStatement & "country_code TEXT(10), "
    sqlStatement = sqlStatement & "valid_from DATETIME, "
    sqlStatement = sqlStatement & "valid_to DATETIME, "
    sqlStatement = sqlStatement & "sort_order LONG, "
    sqlStatement = sqlStatement & "is_active YESNO, "
    sqlStatement = sqlStatement & "created_at DATETIME, "
    sqlStatement = sqlStatement & "created_by TEXT(50), "
    sqlStatement = sqlStatement & "updated_at DATETIME, "
    sqlStatement = sqlStatement & "updated_by TEXT(50)"
    sqlStatement = sqlStatement & ");"

    ExecuteDdl db, sqlStatement
End Sub

Private Sub CreateRefUnits(ByVal db As DAO.Database)
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "CREATE TABLE ref_unit ("
    sqlStatement = sqlStatement & "unit_code TEXT(30) CONSTRAINT pk_ref_unit PRIMARY KEY, "
    sqlStatement = sqlStatement & "translation_key TEXT(100), "
    sqlStatement = sqlStatement & "sort_order LONG, "
    sqlStatement = sqlStatement & "is_active YESNO, "
    sqlStatement = sqlStatement & "created_at DATETIME, "
    sqlStatement = sqlStatement & "created_by TEXT(50), "
    sqlStatement = sqlStatement & "updated_at DATETIME, "
    sqlStatement = sqlStatement & "updated_by TEXT(50)"
    sqlStatement = sqlStatement & ");"

    ExecuteDdl db, sqlStatement
End Sub

' Deprecated:
'   tblAddresses is a legacy table and must not be recreated by current setup paths.
'   adr_address is the authoritative address table.

Private Sub CreateTblProductGroups(ByVal db As DAO.Database)
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "CREATE TABLE art_product_group ("
    sqlStatement = sqlStatement & "product_group_id AUTOINCREMENT CONSTRAINT pk_art_product_group PRIMARY KEY, "
    sqlStatement = sqlStatement & "product_group_code TEXT(50) NOT NULL, "
    sqlStatement = sqlStatement & "product_group_name TEXT(150), "
    sqlStatement = sqlStatement & "description_text LONGTEXT, "
    sqlStatement = sqlStatement & "is_active YESNO, "
    sqlStatement = sqlStatement & "sort_order LONG, "
    sqlStatement = sqlStatement & "created_at DATETIME, "
    sqlStatement = sqlStatement & "created_by TEXT(100), "
    sqlStatement = sqlStatement & "updated_at DATETIME, "
    sqlStatement = sqlStatement & "updated_by TEXT(100)"
    sqlStatement = sqlStatement & ");"

    ExecuteDdl db, sqlStatement
    ExecuteCreateIndexIfMissing db, "art_product_group", "ux_art_product_group_code", _
        "CREATE UNIQUE INDEX ux_art_product_group_code ON art_product_group (product_group_code);"
    ExecuteCreateIndexIfMissing db, "art_product_group", "ix_art_product_group_sort_order", _
        "CREATE INDEX ix_art_product_group_sort_order ON art_product_group (sort_order);"
    ExecuteCreateIndexIfMissing db, "art_product_group", "ix_art_product_group_is_active", _
        "CREATE INDEX ix_art_product_group_is_active ON art_product_group (is_active);"
End Sub

Private Sub CreateRefArticleTypeCodes(ByVal db As DAO.Database)
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "CREATE TABLE ref_article_type_code ("
    sqlStatement = sqlStatement & "article_type_code TEXT(50) CONSTRAINT pk_ref_article_type_code PRIMARY KEY, "
    sqlStatement = sqlStatement & "article_type_name TEXT(100), "
    sqlStatement = sqlStatement & "translation_key TEXT(255), "
    sqlStatement = sqlStatement & "description_text LONGTEXT, "
    sqlStatement = sqlStatement & "sort_order LONG, "
    sqlStatement = sqlStatement & "is_active YESNO, "
    sqlStatement = sqlStatement & "created_at DATETIME, "
    sqlStatement = sqlStatement & "created_by TEXT(255), "
    sqlStatement = sqlStatement & "updated_at DATETIME, "
    sqlStatement = sqlStatement & "updated_by TEXT(255)"
    sqlStatement = sqlStatement & ");"

    ExecuteDdl db, sqlStatement
    ExecuteCreateIndexIfMissing db, "ref_article_type_code", "ix_ref_article_type_code_sort_order", _
        "CREATE INDEX ix_ref_article_type_code_sort_order ON ref_article_type_code (sort_order);"
    ExecuteCreateIndexIfMissing db, "ref_article_type_code", "ix_ref_article_type_code_is_active", _
        "CREATE INDEX ix_ref_article_type_code_is_active ON ref_article_type_code (is_active);"
End Sub

Private Sub CreateTblArticles(ByVal db As DAO.Database)
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "CREATE TABLE art_article ("
    sqlStatement = sqlStatement & "article_id AUTOINCREMENT CONSTRAINT pk_art_article PRIMARY KEY, "
    sqlStatement = sqlStatement & "article_no TEXT(50) NOT NULL, "
    sqlStatement = sqlStatement & "article_name TEXT(150), "
    sqlStatement = sqlStatement & "product_group_id LONG, "
    sqlStatement = sqlStatement & "article_type_code TEXT(30), "
    sqlStatement = sqlStatement & "unit_code TEXT(30), "
    sqlStatement = sqlStatement & "vat_code TEXT(30), "
    sqlStatement = sqlStatement & "purchase_price CURRENCY, "
    sqlStatement = sqlStatement & "sales_price CURRENCY, "
    sqlStatement = sqlStatement & "barcode TEXT(100), "
    sqlStatement = sqlStatement & "description_text LONGTEXT, "
    sqlStatement = sqlStatement & "is_active YESNO, "
    sqlStatement = sqlStatement & "created_at DATETIME, "
    sqlStatement = sqlStatement & "created_by TEXT(100), "
    sqlStatement = sqlStatement & "updated_at DATETIME, "
    sqlStatement = sqlStatement & "updated_by TEXT(100)"
    sqlStatement = sqlStatement & ");"

    ExecuteDdl db, sqlStatement
    ExecuteCreateIndexIfMissing db, "art_article", "ux_art_article_article_no", _
        "CREATE UNIQUE INDEX ux_art_article_article_no ON art_article (article_no);"
    ExecuteCreateIndexIfMissing db, "art_article", "ix_art_article_product_group_id", _
        "CREATE INDEX ix_art_article_product_group_id ON art_article (product_group_id);"
    ExecuteCreateIndexIfMissing db, "art_article", "ix_art_article_unit_code", _
        "CREATE INDEX ix_art_article_unit_code ON art_article (unit_code);"
    ExecuteCreateIndexIfMissing db, "art_article", "ix_art_article_vat_code", _
        "CREATE INDEX ix_art_article_vat_code ON art_article (vat_code);"
End Sub

Private Sub CreateTblOrders(ByVal db As DAO.Database)
    Dim SqlText As String

    SqlText = ""
    SqlText = SqlText & "CREATE TABLE ord_order ("
    SqlText = SqlText & "order_id AUTOINCREMENT CONSTRAINT pk_ord_order PRIMARY KEY, "
    SqlText = SqlText & "order_no TEXT(50), "
    SqlText = SqlText & "order_type_code TEXT(30), "
    SqlText = SqlText & "order_status_code TEXT(30), "
    SqlText = SqlText & "customer_address_id LONG, "
    SqlText = SqlText & "customer_name TEXT(150), "
    SqlText = SqlText & "order_date DATETIME, "
    SqlText = SqlText & "delivery_date DATETIME, "
    SqlText = SqlText & "valid_until DATETIME, "
    SqlText = SqlText & "reference_text TEXT(150), "
    SqlText = SqlText & "external_reference TEXT(150), "
    SqlText = SqlText & "language_code TEXT(10), "
    SqlText = SqlText & "currency_code TEXT(10), "
    SqlText = SqlText & "payment_term_code TEXT(50), "
    SqlText = SqlText & "vat_mode TEXT(20), "
    SqlText = SqlText & "header_discount_type TEXT(20), "
    SqlText = SqlText & "header_discount_value CURRENCY, "
    SqlText = SqlText & "header_discount_amount CURRENCY, "
    SqlText = SqlText & "header_surcharge_type TEXT(20), "
    SqlText = SqlText & "header_surcharge_value CURRENCY, "
    SqlText = SqlText & "header_surcharge_amount CURRENCY, "
    SqlText = SqlText & "subtotal_net_amount CURRENCY, "
    SqlText = SqlText & "net_amount CURRENCY, "
    SqlText = SqlText & "vat_amount CURRENCY, "
    SqlText = SqlText & "gross_amount CURRENCY, "
    SqlText = SqlText & "notes_text LONGTEXT, "
    SqlText = SqlText & "internal_notes_text LONGTEXT, "
    SqlText = SqlText & "result_document_id LONG, "
    SqlText = SqlText & "created_at DATETIME, "
    SqlText = SqlText & "created_by TEXT(50), "
    SqlText = SqlText & "updated_at DATETIME, "
    SqlText = SqlText & "updated_by TEXT(50)"
    SqlText = SqlText & ");"

    ExecuteDdl db, SqlText
End Sub

Private Sub CreateTblOrderLines(ByVal db As DAO.Database)
    Dim SqlText As String

    SqlText = ""
    SqlText = SqlText & "CREATE TABLE ord_order_line ("
    SqlText = SqlText & "order_line_id AUTOINCREMENT CONSTRAINT pk_ord_order_line PRIMARY KEY, "
    SqlText = SqlText & "order_id LONG NOT NULL, "
    SqlText = SqlText & "line_no LONG, "
    SqlText = SqlText & "article_id LONG, "
    SqlText = SqlText & "article_no TEXT(50), "
    SqlText = SqlText & "line_type_code TEXT(30), "
    SqlText = SqlText & "description_text LONGTEXT, "
    SqlText = SqlText & "quantity DOUBLE, "
    SqlText = SqlText & "unit_code TEXT(30), "
    SqlText = SqlText & "unit_price CURRENCY, "
    SqlText = SqlText & "discount_type TEXT(20), "
    SqlText = SqlText & "discount_value CURRENCY, "
    SqlText = SqlText & "line_discount_amount CURRENCY, "
    SqlText = SqlText & "surcharge_type TEXT(20), "
    SqlText = SqlText & "surcharge_value CURRENCY, "
    SqlText = SqlText & "line_surcharge_amount CURRENCY, "
    SqlText = SqlText & "vat_code TEXT(30), "
    SqlText = SqlText & "vat_rate DOUBLE, "
    SqlText = SqlText & "line_base_amount CURRENCY, "
    SqlText = SqlText & "line_net_amount CURRENCY, "
    SqlText = SqlText & "line_vat_amount CURRENCY, "
    SqlText = SqlText & "line_gross_amount CURRENCY, "
    SqlText = SqlText & "created_at DATETIME, "
    SqlText = SqlText & "created_by TEXT(50), "
    SqlText = SqlText & "updated_at DATETIME, "
    SqlText = SqlText & "updated_by TEXT(50)"
    SqlText = SqlText & ");"

    ExecuteDdl db, SqlText
End Sub

Private Function EnsureOrderPhase1SchemaForDatabase(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    EnsureOrderPhase1SchemaForDatabase = False

    If db Is Nothing Then
        Exit Function
    End If

    CreateTblOrders db
    CreateTblOrderLines db

    If Not TableExists(db, "ord_order") Then Exit Function
    If Not TableExists(db, "ord_order_line") Then Exit Function

    If Not EnsureOrderHeaderSchema(db) Then GoTo CleanExit
    If Not EnsureOrderLineSchema(db) Then GoTo CleanExit
    If Not VerifyRequiredOrderFields(db) Then GoTo CleanExit
    If Not EnsureOrderHeaderIndexes(db) Then GoTo CleanExit
    If Not EnsureOrderLineIndexes(db) Then GoTo CleanExit

    EnsureOrderPhase1SchemaForDatabase = True

CleanExit:
    Exit Function

ErrorHandler:
    EnsureOrderPhase1SchemaForDatabase = False
End Function

Private Function EnsureOrderHeaderSchema(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    If Not EnsureTextField(db, "ord_order", "order_no", 50, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "order_type_code", 30, "SO", True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "order_status_code", 30, "DRAFT", True) Then GoTo ErrorHandler
    If Not EnsureLongField(db, "ord_order", "customer_address_id", 0, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "customer_name", 150, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "ord_order", "order_date", Date, True) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "ord_order", "delivery_date", 0, False) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "ord_order", "valid_until", 0, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "reference_text", 150, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "external_reference", 150, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "language_code", 10, "DE-CH", True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "currency_code", 10, "CHF", True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "payment_term_code", 50, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "vat_mode", 20, "EXCLUSIVE", True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "header_discount_type", 20, "NONE", True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order", "header_discount_value", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order", "header_discount_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "header_surcharge_type", 20, "NONE", True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order", "header_surcharge_value", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order", "header_surcharge_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order", "subtotal_net_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order", "net_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order", "vat_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order", "gross_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureLongTextField(db, "ord_order", "notes_text") Then GoTo ErrorHandler
    If Not EnsureLongTextField(db, "ord_order", "internal_notes_text") Then GoTo ErrorHandler
    If Not EnsureLongField(db, "ord_order", "result_document_id", 0, False) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "ord_order", "created_at", Now(), True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "created_by", 50, "SYSTEM", True) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "ord_order", "updated_at", Now(), True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "updated_by", 50, "SYSTEM", True) Then GoTo ErrorHandler

    MigrateLegacyOrderHeaderFields db

    EnsureOrderHeaderSchema = True
    Exit Function

ErrorHandler:
    EnsureOrderHeaderSchema = False
End Function

Private Function EnsureOrderLineSchema(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    If Not EnsureLongField(db, "ord_order_line", "order_id", 0, False) Then GoTo ErrorHandler
    If Not EnsureLongField(db, "ord_order_line", "line_no", 0, False) Then GoTo ErrorHandler
    If Not EnsureLongField(db, "ord_order_line", "article_id", 0, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order_line", "article_no", 50, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order_line", "line_type_code", 30, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureLongTextField(db, "ord_order_line", "description_text") Then GoTo ErrorHandler
    If Not EnsureDoubleField(db, "ord_order_line", "quantity", 0, True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order_line", "unit_code", 30, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order_line", "unit_price", 0, True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order_line", "discount_type", 20, "NONE", True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order_line", "discount_value", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order_line", "line_discount_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order_line", "surcharge_type", 20, "NONE", True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order_line", "surcharge_value", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order_line", "line_surcharge_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order_line", "vat_code", 30, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureDoubleField(db, "ord_order_line", "vat_rate", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order_line", "line_base_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order_line", "line_net_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order_line", "line_vat_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "ord_order_line", "line_gross_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "ord_order_line", "created_at", Now(), True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order_line", "created_by", 50, "SYSTEM", True) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "ord_order_line", "updated_at", Now(), True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order_line", "updated_by", 50, "SYSTEM", True) Then GoTo ErrorHandler

    MigrateLegacyOrderLineFields db

    EnsureOrderLineSchema = True
    Exit Function

ErrorHandler:
    EnsureOrderLineSchema = False
End Function

Private Function VerifyRequiredOrderFields(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    VerifyRequiredOrderFields = False

    If Not EnsureRequiredFieldExists(db, "ord_order", "customer_address_id") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "ord_order_line", "article_no") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "ord_order_line", "vat_rate") Then Exit Function

    VerifyRequiredOrderFields = True
    Exit Function

ErrorHandler:
    VerifyRequiredOrderFields = False
End Function

Private Sub MigrateLegacyOrderHeaderFields(ByVal db As DAO.Database)
    On Error Resume Next

    If FieldExists(db, "ord_order", "address_id") Then
        db.Execute "UPDATE ord_order SET customer_address_id = address_id WHERE Nz(customer_address_id, 0)=0 AND address_id IS NOT NULL;", dbFailOnError
    End If

    If FieldExists(db, "ord_order", "total_discount_amount") Then
        db.Execute "UPDATE ord_order SET header_discount_amount = total_discount_amount WHERE Nz(header_discount_amount, 0)=0 AND total_discount_amount IS NOT NULL;", dbFailOnError
    End If

    If FieldExists(db, "ord_order", "total_surcharge_amount") Then
        db.Execute "UPDATE ord_order SET header_surcharge_amount = total_surcharge_amount WHERE Nz(header_surcharge_amount, 0)=0 AND total_surcharge_amount IS NOT NULL;", dbFailOnError
    End If

    If FieldExists(db, "ord_order", "total_vat_amount") Then
        db.Execute "UPDATE ord_order SET vat_amount = total_vat_amount WHERE Nz(vat_amount, 0)=0 AND total_vat_amount IS NOT NULL;", dbFailOnError
    End If

    If FieldExists(db, "ord_order", "total_gross_amount") Then
        db.Execute "UPDATE ord_order SET gross_amount = total_gross_amount WHERE Nz(gross_amount, 0)=0 AND total_gross_amount IS NOT NULL;", dbFailOnError
    End If

    db.Execute "UPDATE ord_order SET net_amount = subtotal_net_amount WHERE Nz(net_amount, 0)=0 AND subtotal_net_amount IS NOT NULL;", dbFailOnError
End Sub

Private Sub MigrateLegacyOrderLineFields(ByVal db As DAO.Database)
    On Error Resume Next

    If FieldExists(db, "ord_order_line", "discount_amount") Then
        db.Execute "UPDATE ord_order_line SET line_discount_amount = discount_amount WHERE Nz(line_discount_amount, 0)=0 AND discount_amount IS NOT NULL;", dbFailOnError
    End If

    If FieldExists(db, "ord_order_line", "surcharge_amount") Then
        db.Execute "UPDATE ord_order_line SET line_surcharge_amount = surcharge_amount WHERE Nz(line_surcharge_amount, 0)=0 AND surcharge_amount IS NOT NULL;", dbFailOnError
    End If

    If FieldExists(db, "ord_order_line", "line_total") Then
        db.Execute "UPDATE ord_order_line SET line_net_amount = line_total WHERE Nz(line_net_amount, 0)=0 AND line_total IS NOT NULL;", dbFailOnError
    End If
End Sub

Private Function EnsureOrderHeaderIndexes(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    EnsureIndexWhenFieldExists db, "ord_order", "order_no", "ux_ord_order_order_no", _
        "CREATE UNIQUE INDEX ux_ord_order_order_no ON ord_order (order_no);"
    EnsureIndexWhenFieldExists db, "ord_order", "customer_address_id", "ix_ord_order_customer_address_id", _
        "CREATE INDEX ix_ord_order_customer_address_id ON ord_order (customer_address_id);"
    EnsureIndexWhenFieldExists db, "ord_order", "order_date", "ix_ord_order_order_date", _
        "CREATE INDEX ix_ord_order_order_date ON ord_order (order_date);"
    EnsureIndexWhenFieldExists db, "ord_order", "order_status_code", "ix_ord_order_order_status_code", _
        "CREATE INDEX ix_ord_order_order_status_code ON ord_order (order_status_code);"
    EnsureIndexWhenFieldExists db, "ord_order", "payment_term_code", "ix_ord_order_payment_term_code", _
        "CREATE INDEX ix_ord_order_payment_term_code ON ord_order (payment_term_code);"
    EnsureIndexWhenFieldExists db, "ord_order", "result_document_id", "ix_ord_order_result_document_id", _
        "CREATE INDEX ix_ord_order_result_document_id ON ord_order (result_document_id);"

    EnsureOrderHeaderIndexes = True
    Exit Function

ErrorHandler:
    EnsureOrderHeaderIndexes = False
End Function

Private Function EnsureOrderLineIndexes(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    EnsureIndexWhenFieldExists db, "ord_order_line", "order_id", "ix_ord_order_line_order_id", _
        "CREATE INDEX ix_ord_order_line_order_id ON ord_order_line (order_id);"
    EnsureIndexWhenFieldExists db, "ord_order_line", "article_id", "ix_ord_order_line_article_id", _
        "CREATE INDEX ix_ord_order_line_article_id ON ord_order_line (article_id);"
    EnsureIndexWhenFieldExists db, "ord_order_line", "article_no", "ix_ord_order_line_article_no", _
        "CREATE INDEX ix_ord_order_line_article_no ON ord_order_line (article_no);"
    EnsureIndexWhenFieldExists db, "ord_order_line", "line_no", "ix_ord_order_line_line_no", _
        "CREATE INDEX ix_ord_order_line_line_no ON ord_order_line (line_no);"
    EnsureIndexWhenFieldExists db, "ord_order_line", "unit_code", "ix_ord_order_line_unit_code", _
        "CREATE INDEX ix_ord_order_line_unit_code ON ord_order_line (unit_code);"
    EnsureIndexWhenFieldExists db, "ord_order_line", "vat_code", "ix_ord_order_line_vat_code", _
        "CREATE INDEX ix_ord_order_line_vat_code ON ord_order_line (vat_code);"

    EnsureOrderLineIndexes = True
    Exit Function

ErrorHandler:
    EnsureOrderLineIndexes = False
End Function

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

Private Function FieldExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field

    RefreshTableDefinition db, tableName
    Set tdf = db.TableDefs(tableName)
    For Each fld In tdf.Fields
        If StrComp(Trim$(fld.Name), Trim$(fieldName), vbTextCompare) = 0 Then
            FieldExists = True
            Exit Function
        End If
    Next fld
    Exit Function

ErrorHandler:
    FieldExists = False
End Function

Private Function EnsureTextField(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByVal FieldSize As Long, ByVal defaultValue As String, ByVal updateNullValues As Boolean) As Boolean
    On Error GoTo ErrorHandler

    If Not FieldExists(db, tableName, fieldName) Then
        db.Execute "ALTER TABLE [" & tableName & "] ADD COLUMN [" & fieldName & "] TEXT(" & CStr(FieldSize) & ");", dbFailOnError
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureTextField", "Field ensured: " & tableName & "." & fieldName
    Else
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureTextField", "Field exists: " & tableName & "." & fieldName
    End If

    If LenB(defaultValue) > 0 Then
        ApplyFieldDefaultValue db, tableName, fieldName, """" & Replace(defaultValue, """", """""") & """"
        If updateNullValues Then
            db.Execute "UPDATE [" & tableName & "] SET [" & fieldName & "]='" & Replace(defaultValue, "'", "''") & "' WHERE [" & fieldName & "] IS NULL;", dbFailOnError
        End If
    End If

    EnsureTextField = True
    Exit Function

ErrorHandler:
    EnsureTextField = False
End Function

Private Function EnsureLongField(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByVal defaultValue As Long, ByVal updateNullValues As Boolean) As Boolean
    On Error GoTo ErrorHandler

    If Not FieldExists(db, tableName, fieldName) Then
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureLongField", _
            "AddField executing: " & tableName & "." & fieldName & "; db=" & ResolveDatabasePath(db)
        db.Execute "ALTER TABLE [" & tableName & "] ADD COLUMN [" & fieldName & "] LONG;", dbFailOnError
        RefreshTableDefinition db, tableName
        LogTableFieldNames db, tableName, MODULE_NAME & ".EnsureLongField"
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureLongField", _
            "AddField successful: " & tableName & "." & fieldName & "; exists_after_add=" & CStr(FieldExists(db, tableName, fieldName))
        If Not FieldExists(db, tableName, fieldName) Then
            modLoggingHandler.LogError MODULE_NAME & ".EnsureLongField", _
                "Required field " & tableName & "." & fieldName & " could not be ensured."
            GoTo ErrorHandler
        End If
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureLongField", "Field ensured: " & tableName & "." & fieldName
    Else
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureLongField", "Field exists: " & tableName & "." & fieldName
    End If

    ApplyFieldDefaultValue db, tableName, fieldName, CStr(defaultValue)
    If updateNullValues Then
        db.Execute "UPDATE [" & tableName & "] SET [" & fieldName & "]=" & CStr(defaultValue) & " WHERE [" & fieldName & "] IS NULL;", dbFailOnError
    End If

    EnsureLongField = True
    Exit Function

ErrorHandler:
    modLoggingHandler.LogError MODULE_NAME & ".EnsureLongField", _
        "AddField failed for " & tableName & "." & fieldName & " in db=" & ResolveDatabasePath(db) & ": " & Err.description, Err.Number
    EnsureLongField = False
End Function

Private Function EnsureDoubleField(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByVal defaultValue As Double, ByVal updateNullValues As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim numericText As String

    If Not FieldExists(db, tableName, fieldName) Then
        db.Execute "ALTER TABLE [" & tableName & "] ADD COLUMN [" & fieldName & "] DOUBLE;", dbFailOnError
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureDoubleField", "Field ensured: " & tableName & "." & fieldName
    Else
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureDoubleField", "Field exists: " & tableName & "." & fieldName
    End If

    numericText = Replace(CStr(defaultValue), ",", ".")
    ApplyFieldDefaultValue db, tableName, fieldName, numericText
    If updateNullValues Then
        db.Execute "UPDATE [" & tableName & "] SET [" & fieldName & "]=" & numericText & " WHERE [" & fieldName & "] IS NULL;", dbFailOnError
    End If

    EnsureDoubleField = True
    Exit Function

ErrorHandler:
    EnsureDoubleField = False
End Function

Private Function EnsureCurrencyField(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByVal defaultValue As Currency, ByVal updateNullValues As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim numericText As String

    If Not FieldExists(db, tableName, fieldName) Then
        db.Execute "ALTER TABLE [" & tableName & "] ADD COLUMN [" & fieldName & "] CURRENCY;", dbFailOnError
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureCurrencyField", "Field ensured: " & tableName & "." & fieldName
    Else
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureCurrencyField", "Field exists: " & tableName & "." & fieldName
    End If

    numericText = Replace(CStr(defaultValue), ",", ".")
    ApplyFieldDefaultValue db, tableName, fieldName, numericText
    If updateNullValues Then
        db.Execute "UPDATE [" & tableName & "] SET [" & fieldName & "]=" & numericText & " WHERE [" & fieldName & "] IS NULL;", dbFailOnError
    End If

    EnsureCurrencyField = True
    Exit Function

ErrorHandler:
    EnsureCurrencyField = False
End Function

Private Function EnsureDateField(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByVal defaultValue As Date, ByVal updateNullValues As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim defaultExpression As String
    Dim updateValue As String

    If Not FieldExists(db, tableName, fieldName) Then
        db.Execute "ALTER TABLE [" & tableName & "] ADD COLUMN [" & fieldName & "] DATETIME;", dbFailOnError
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureDateField", "Field ensured: " & tableName & "." & fieldName
    Else
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureDateField", "Field exists: " & tableName & "." & fieldName
    End If

    If defaultValue = 0 Then
        defaultExpression = "Null"
    Else
        updateValue = "#" & Format$(defaultValue, "yyyy-mm-dd hh:nn:ss") & "#"
        defaultExpression = updateValue
    End If

    ApplyFieldDefaultValue db, tableName, fieldName, defaultExpression
    If updateNullValues And LenB(updateValue) > 0 Then
        db.Execute "UPDATE [" & tableName & "] SET [" & fieldName & "]=" & updateValue & " WHERE [" & fieldName & "] IS NULL;", dbFailOnError
    End If

    EnsureDateField = True
    Exit Function

ErrorHandler:
    EnsureDateField = False
End Function

Private Function EnsureLongTextField(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) As Boolean
    On Error GoTo ErrorHandler

    If Not FieldExists(db, tableName, fieldName) Then
        db.Execute "ALTER TABLE [" & tableName & "] ADD COLUMN [" & fieldName & "] LONGTEXT;", dbFailOnError
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureLongTextField", "Field ensured: " & tableName & "." & fieldName
    Else
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureLongTextField", "Field exists: " & tableName & "." & fieldName
    End If

    EnsureLongTextField = True
    Exit Function

ErrorHandler:
    EnsureLongTextField = False
End Function

Private Sub ApplyFieldDefaultValue(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByVal DefaultValueExpression As String)
    On Error Resume Next

    db.TableDefs(tableName).Fields(fieldName).defaultValue = DefaultValueExpression
End Sub

Private Function EnsureRequiredFieldExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String) As Boolean
    On Error GoTo ErrorHandler

    If FieldExists(db, tableName, fieldName) Then
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureRequiredFieldExists", "Field exists: " & tableName & "." & fieldName
        EnsureRequiredFieldExists = True
        Exit Function
    End If

    modLoggingHandler.LogError MODULE_NAME & ".EnsureRequiredFieldExists", _
        "Required field " & tableName & "." & fieldName & " could not be ensured."
    Exit Function

ErrorHandler:
    EnsureRequiredFieldExists = False
End Function

Private Sub EnsureIndexWhenFieldExists( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal fieldName As String, _
    ByVal indexName As String, _
    ByVal SqlText As String)
    On Error GoTo ErrorHandler

    If Not FieldExists(db, tableName, fieldName) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".EnsureIndexWhenFieldExists", _
            "Index skipped because field missing: " & tableName & "." & fieldName & " -> " & indexName
        Exit Sub
    End If

    ExecuteCreateIndexIfMissing db, tableName, indexName, SqlText
    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureIndexWhenFieldExists", "Index ensured: " & indexName
    Exit Sub

ErrorHandler:
    modLoggingHandler.LogWarning MODULE_NAME & ".EnsureIndexWhenFieldExists", _
        "Index skipped because ensure failed: " & indexName & " (" & Err.Number & " - " & Err.description & ")"
End Sub

Private Function IsLinkedAccessTable(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim connectText As String

    If db Is Nothing Then
        Exit Function
    End If

    connectText = Trim$(Nz(db.TableDefs(tableName).Connect, vbNullString))
    IsLinkedAccessTable = (InStr(1, connectText, ACCESS_CONNECT_PREFIX, vbTextCompare) > 0)
    Exit Function

ErrorHandler:
    IsLinkedAccessTable = False
End Function

Private Function ResolveLinkedTableBackendPath(ByVal db As DAO.Database, ByVal tableName As String) As String
    On Error GoTo ErrorHandler

    Dim connectText As String
    Dim markerPosition As Long

    If db Is Nothing Then
        Exit Function
    End If

    connectText = Trim$(Nz(db.TableDefs(tableName).Connect, vbNullString))
    markerPosition = InStr(1, connectText, ACCESS_CONNECT_PREFIX, vbTextCompare)
    If markerPosition <= 0 Then
        Exit Function
    End If

    ResolveLinkedTableBackendPath = Trim$(Mid$(connectText, markerPosition + Len(ACCESS_CONNECT_PREFIX)))
    Exit Function

ErrorHandler:
    ResolveLinkedTableBackendPath = vbNullString
End Function

Private Function NormalizePath(ByVal pathText As String) As String
    NormalizePath = LCase$(Trim$(Replace(pathText, "/", "\")))
End Function

Private Function ResolveDatabasePath(ByVal db As DAO.Database) As String
    On Error GoTo ErrorHandler

    If db Is Nothing Then
        ResolveDatabasePath = "(none)"
    Else
        ResolveDatabasePath = Nz(db.Name, vbNullString)
    End If
    Exit Function

ErrorHandler:
    ResolveDatabasePath = "(unresolved)"
End Function

Private Sub RefreshTableDefinition(ByVal db As DAO.Database, ByVal tableName As String)
    On Error Resume Next

    Dim tdf As DAO.tableDef

    If db Is Nothing Then
        Exit Sub
    End If

    db.TableDefs.Refresh
    Set tdf = db.TableDefs(tableName)
    Set tdf = Nothing
End Sub

Private Sub LogTableFieldNames(ByVal db As DAO.Database, ByVal tableName As String, ByVal SourceProcedure As String)
    On Error GoTo ErrorHandler

    Dim tdf As DAO.tableDef
    Dim fld As DAO.Field
    Dim fieldList As String

    RefreshTableDefinition db, tableName
    Set tdf = db.TableDefs(tableName)

    For Each fld In tdf.Fields
        If LenB(fieldList) > 0 Then
            fieldList = fieldList & ", "
        End If
        fieldList = fieldList & fld.Name
    Next fld

    modLoggingHandler.LogInfo SourceProcedure, "Fields in " & tableName & ": " & fieldList
    Exit Sub

ErrorHandler:
    modLoggingHandler.LogWarning SourceProcedure, _
        "Could not enumerate fields for " & tableName & " (" & Err.Number & " - " & Err.description & ")"
End Sub

Private Sub ExecuteCreateIndexIfMissing( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal indexName As String, _
    ByVal SqlText As String)
    On Error GoTo ErrorHandler

    If IndexExists(db, tableName, indexName) Then
        Debug.Print "SKIP INDEX: " & tableName & "." & indexName
        Exit Sub
    End If

    ExecuteDdl db, SqlText
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & ".ExecuteCreateIndexIfMissing", Err.description
End Sub

Private Function IndexExists( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal indexName As String) As Boolean
    On Error GoTo SafeExit

    Dim tableDefinition As DAO.tableDef
    Dim indexDefinition As DAO.index

    If db Is Nothing Then
        Exit Function
    End If

    Set tableDefinition = db.TableDefs(tableName)
    For Each indexDefinition In tableDefinition.Indexes
        If StrComp(indexDefinition.Name, indexName, vbTextCompare) = 0 Then
            IndexExists = True
            Exit Function
        End If
    Next indexDefinition

SafeExit:
    Set indexDefinition = Nothing
    Set tableDefinition = Nothing
End Function

Private Sub ExecuteDdl(ByVal db As DAO.Database, ByVal SqlText As String)
    On Error GoTo ErrorHandler

    db.Execute SqlText, dbFailOnError
    Debug.Print "OK: " & SqlText
    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 3010, 3283, 3371
            Debug.Print "SKIP: " & Err.Number & " - " & Err.description
            Err.Clear

        Case Else
            Debug.Print "ERROR: " & Err.Number & " - " & Err.description
            Debug.Print SqlText
            Err.Raise Err.Number, MODULE_NAME & ".ExecuteDdl", Err.description
    End Select
End Sub