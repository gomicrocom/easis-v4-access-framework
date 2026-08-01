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
Private Const SYSTEM_BACKEND_PATH As String = "C:\easis\Data\sys_be.accdb"
Private Const TABLE_REF_LANGUAGE As String = "ref_language"

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
    CreateTmpOrders db
    CreateTmpOrderLines db
    Call EnsureOrderPhase1SchemaForDatabase(db)
    Call modOrderRepository.EnsureSalesOrderNumberRange(Year(Date))
    Call EnsureSystemLanguageReferenceSchema

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
    Dim frontendDb As DAO.Database
    Dim shouldCloseDb As Boolean

    If Not OpenSchemaDatabase(backendPath, db, shouldCloseDb) Then
        Exit Function
    End If

    EnsureOrderPhase1Schema = EnsureOrderPhase1SchemaForDatabase(db)
    If EnsureOrderPhase1Schema Then
        Set frontendDb = CurrentDb
        If Not EnsureTemporaryOrderWorkspaceSchema(frontendDb) Then
            EnsureOrderPhase1Schema = False
        End If
    End If

CleanExit:
    On Error Resume Next
    Set frontendDb = Nothing
    If Not db Is Nothing Then
        If shouldCloseDb Then db.Close
    End If
    Set db = Nothing
    Exit Function

ErrorHandler:
    EnsureOrderPhase1Schema = False
    Resume CleanExit
End Function

Private Function EnsureTemporaryOrderWorkspaceSchema(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    EnsureTemporaryOrderWorkspaceSchema = False

    If db Is Nothing Then
        Exit Function
    End If

    CreateTmpOrders db
    CreateTmpOrderLines db

    If Not modDbSchema.TableExists(db, "tmp_order") Then Exit Function
    If Not modDbSchema.TableExists(db, "tmp_order_line") Then Exit Function
    If Not EnsureTemporaryOrderHeaderSchema(db) Then Exit Function
    If Not EnsureTemporaryOrderLineSchema(db) Then Exit Function
    If Not EnsureRequiredFieldExists(db, "tmp_order", "session_id") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "tmp_order", "order_no") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "tmp_order_line", "tmp_order_id") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "tmp_order_line", "order_line_id") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "tmp_order_line", "vat_rate") Then Exit Function
    If Not EnsureTemporaryOrderHeaderIndexes(db) Then Exit Function
    If Not EnsureTemporaryOrderLineIndexes(db) Then Exit Function

    EnsureTemporaryOrderWorkspaceSchema = True
    Exit Function

ErrorHandler:
    EnsureTemporaryOrderWorkspaceSchema = False
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
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "table_exists ord_order=" & CStr(modDbSchema.TableExists(db, "ord_order"))
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "field_exists ord_order.customer_address_id=" & CStr(modDbSchema.FieldExists(db, "ord_order", "customer_address_id"))
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "field_exists ord_order.invoice_address_id=" & CStr(modDbSchema.FieldExists(db, "ord_order", "invoice_address_id"))
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "field_exists ord_order.delivery_address_id=" & CStr(modDbSchema.FieldExists(db, "ord_order", "delivery_address_id"))
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "field_exists ord_order.vat_code=" & CStr(modDbSchema.FieldExists(db, "ord_order", "vat_code"))
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "field_exists ord_order.vat_rate=" & CStr(modDbSchema.FieldExists(db, "ord_order", "vat_rate"))
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "table_exists ord_order_line=" & CStr(modDbSchema.TableExists(db, "ord_order_line"))
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "field_exists ord_order_line.article_no=" & CStr(modDbSchema.FieldExists(db, "ord_order_line", "article_no"))
    modLoggingHandler.LogInfo MODULE_NAME & ".DiagnoseOrderSchema", "field_exists ord_order_line.vat_rate=" & CStr(modDbSchema.FieldExists(db, "ord_order_line", "vat_rate"))
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
        If modDbSchema.TableExists(frontendDb, "ord_order") Then
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

Private Sub CreateRefLanguage(ByVal db As DAO.Database)
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "CREATE TABLE ref_language ("
    sqlStatement = sqlStatement & "language_code TEXT(10) CONSTRAINT pk_ref_language PRIMARY KEY, "
    sqlStatement = sqlStatement & "language_name TEXT(100), "
    sqlStatement = sqlStatement & "iso_language_code TEXT(10), "
    sqlStatement = sqlStatement & "country_code TEXT(10), "
    sqlStatement = sqlStatement & "is_default YESNO, "
    sqlStatement = sqlStatement & "is_active YESNO, "
    sqlStatement = sqlStatement & "sort_order LONG, "
    sqlStatement = sqlStatement & "created_at DATETIME, "
    sqlStatement = sqlStatement & "created_by TEXT(50), "
    sqlStatement = sqlStatement & "updated_at DATETIME, "
    sqlStatement = sqlStatement & "updated_by TEXT(50)"
    sqlStatement = sqlStatement & ");"

    ExecuteDdl db, sqlStatement
    ExecuteCreateIndexIfMissing db, TABLE_REF_LANGUAGE, "ix_ref_language_is_active", _
        "CREATE INDEX ix_ref_language_is_active ON ref_language (is_active);"
    ExecuteCreateIndexIfMissing db, TABLE_REF_LANGUAGE, "ix_ref_language_sort_order", _
        "CREATE INDEX ix_ref_language_sort_order ON ref_language (sort_order);"
End Sub

Private Sub CreateTenPaymentTerms(ByVal db As DAO.Database)
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "CREATE TABLE ten_payment_term ("
    sqlStatement = sqlStatement & "payment_term_id AUTOINCREMENT CONSTRAINT pk_ten_payment_term PRIMARY KEY, "
    sqlStatement = sqlStatement & "payment_term_code TEXT(50) NOT NULL, "
    sqlStatement = sqlStatement & "payment_term_type_code TEXT(50), "
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
    ExecuteCreateIndexIfMissing db, "ten_payment_term", "ux_ten_payment_term_code", _
        "CREATE UNIQUE INDEX ux_ten_payment_term_code ON ten_payment_term (payment_term_code);"
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

Private Function EnsureRefLanguageSchema(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    EnsureRefLanguageSchema = False

    If Not EnsureTextField(db, TABLE_REF_LANGUAGE, "language_code", 10, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, TABLE_REF_LANGUAGE, "language_name", 100, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, TABLE_REF_LANGUAGE, "iso_language_code", 10, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, TABLE_REF_LANGUAGE, "country_code", 10, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureLongField(db, TABLE_REF_LANGUAGE, "sort_order", 0, False) Then GoTo ErrorHandler
    If Not EnsureDateField(db, TABLE_REF_LANGUAGE, "created_at", Now(), False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, TABLE_REF_LANGUAGE, "created_by", 50, "SYSTEM", False) Then GoTo ErrorHandler
    If Not EnsureDateField(db, TABLE_REF_LANGUAGE, "updated_at", Now(), False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, TABLE_REF_LANGUAGE, "updated_by", 50, "SYSTEM", False) Then GoTo ErrorHandler

    If Not modDbSchema.FieldExists(db, TABLE_REF_LANGUAGE, "is_default") Then
        db.Execute "ALTER TABLE [" & TABLE_REF_LANGUAGE & "] ADD COLUMN [is_default] YESNO;", dbFailOnError
    End If
    If Not modDbSchema.FieldExists(db, TABLE_REF_LANGUAGE, "is_active") Then
        db.Execute "ALTER TABLE [" & TABLE_REF_LANGUAGE & "] ADD COLUMN [is_active] YESNO;", dbFailOnError
    End If

    ExecuteCreateIndexIfMissing db, TABLE_REF_LANGUAGE, "ix_ref_language_is_active", _
        "CREATE INDEX ix_ref_language_is_active ON ref_language (is_active);"
    ExecuteCreateIndexIfMissing db, TABLE_REF_LANGUAGE, "ix_ref_language_sort_order", _
        "CREATE INDEX ix_ref_language_sort_order ON ref_language (sort_order);"

    EnsureRefLanguageSchema = True
    Exit Function

ErrorHandler:
    EnsureRefLanguageSchema = False
End Function

Private Function SeedRefLanguageData(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    SeedRefLanguageData = False

    UpsertRefLanguage db, "de-CH", "Deutsch (Schweiz)", "de", "CH", False, True, 10
    UpsertRefLanguage db, "fr-CH", "Francais (Suisse)", "fr", "CH", False, True, 20
    UpsertRefLanguage db, "en-US", "English (United States)", "en", "US", True, True, 30
    UpsertRefLanguage db, "it-CH", "Italiano (Svizzera)", "it", "CH", False, False, 40
    UpsertRefLanguage db, "de-DE", "Deutsch (Deutschland)", "de", "DE", False, False, 50

    SeedRefLanguageData = True
    Exit Function

ErrorHandler:
    SeedRefLanguageData = False
End Function

Private Sub UpsertRefLanguage( _
    ByVal db As DAO.Database, _
    ByVal languageCode As String, _
    ByVal languageName As String, _
    ByVal isoLanguageCode As String, _
    ByVal countryCode As String, _
    ByVal isDefault As Boolean, _
    ByVal isActive As Boolean, _
    ByVal sortOrder As Long)
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlStatement As String

    sqlStatement = "SELECT * FROM [" & TABLE_REF_LANGUAGE & "] WHERE [language_code]=" & SqlText(languageCode) & ";"
    Set rs = db.OpenRecordset(sqlStatement, dbOpenDynaset)

    If rs.BOF And rs.EOF Then
        rs.AddNew
        rs.Fields("language_code").Value = languageCode
        rs.Fields("created_at").Value = Now()
        rs.Fields("created_by").Value = "SYSTEM"
        rs.Fields("is_active").Value = isActive
    Else
        rs.Edit
    End If

    rs.Fields("language_name").Value = languageName
    rs.Fields("iso_language_code").Value = isoLanguageCode
    rs.Fields("country_code").Value = countryCode
    rs.Fields("is_default").Value = isDefault
    rs.Fields("sort_order").Value = sortOrder
    rs.Fields("updated_at").Value = Now()
    rs.Fields("updated_by").Value = "SYSTEM"
    rs.Update

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & ".UpsertRefLanguage", Err.description
End Sub

Private Function EnsureLinkedAccessTable(ByVal frontendDb As DAO.Database, ByVal backendPath As String, ByVal tableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim existingTableDef As DAO.TableDef
    Dim newTableDef As DAO.TableDef
    Dim existingConnect As String

    EnsureLinkedAccessTable = False

    If frontendDb Is Nothing Then
        Exit Function
    End If

    If LenB(Trim$(backendPath)) = 0 Then
        Exit Function
    End If

    If modDbSchema.TableExists(frontendDb, tableName) Then
        Set existingTableDef = frontendDb.TableDefs(tableName)
        existingConnect = Trim$(Nz(existingTableDef.Connect, vbNullString))

        If LenB(existingConnect) = 0 Then
            modLoggingHandler.LogError MODULE_NAME & ".EnsureLinkedAccessTable", _
                "Local frontend table blocks required link: " & tableName
            Exit Function
        End If

        frontendDb.TableDefs.Delete tableName
        frontendDb.TableDefs.Refresh
    End If

    Set newTableDef = frontendDb.CreateTableDef(tableName)
    newTableDef.Connect = ACCESS_CONNECT_PREFIX & backendPath
    newTableDef.SourceTableName = tableName
    frontendDb.TableDefs.Append newTableDef
    frontendDb.TableDefs.Refresh

    EnsureLinkedAccessTable = True
    Exit Function

ErrorHandler:
    EnsureLinkedAccessTable = False
End Function

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
    SqlText = SqlText & "invoice_address_id LONG, "
    SqlText = SqlText & "delivery_address_id LONG, "
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
    SqlText = SqlText & "vat_code TEXT(30), "
    SqlText = SqlText & "vat_rate DOUBLE, "
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
    SqlText = SqlText & "sort_order LONG, "
    SqlText = SqlText & "created_at DATETIME, "
    SqlText = SqlText & "created_by TEXT(50), "
    SqlText = SqlText & "updated_at DATETIME, "
    SqlText = SqlText & "updated_by TEXT(50)"
    SqlText = SqlText & ");"

    ExecuteDdl db, SqlText
End Sub

Public Function EnsureSystemLanguageReferenceSchema(Optional ByVal sysBackendPath As String = vbNullString) As Boolean
    On Error GoTo ErrorHandler

    Dim backendDb As DAO.Database
    Dim frontendDb As DAO.Database
    Dim effectiveBackendPath As String

    effectiveBackendPath = Trim$(sysBackendPath)
    If LenB(effectiveBackendPath) = 0 Then
        effectiveBackendPath = SYSTEM_BACKEND_PATH
    End If

    Set backendDb = OpenOrCreateAccessDatabase(effectiveBackendPath)
    CreateRefLanguage backendDb
    If Not EnsureRefLanguageSchema(backendDb) Then GoTo CleanExit
    If Not SeedRefLanguageData(backendDb) Then GoTo CleanExit

    Set frontendDb = CurrentDb
    If Not EnsureLinkedAccessTable(frontendDb, effectiveBackendPath, TABLE_REF_LANGUAGE) Then GoTo CleanExit

    EnsureSystemLanguageReferenceSchema = True

CleanExit:
    On Error Resume Next
    If Not backendDb Is Nothing Then backendDb.Close
    Set frontendDb = Nothing
    Set backendDb = Nothing
    Exit Function

ErrorHandler:
    EnsureSystemLanguageReferenceSchema = False
    modErrorHandler.HandleError MODULE_NAME, "EnsureSystemLanguageReferenceSchema", Err
    Resume CleanExit
End Function

Private Sub CreateTmpOrders(ByVal db As DAO.Database)
    Dim SqlText As String

    SqlText = ""
    SqlText = SqlText & "CREATE TABLE tmp_order ("
    SqlText = SqlText & "tmp_order_id AUTOINCREMENT CONSTRAINT pk_tmp_order PRIMARY KEY, "
    SqlText = SqlText & "session_id TEXT(100), "
    SqlText = SqlText & "order_id LONG, "
    SqlText = SqlText & "order_no TEXT(50), "
    SqlText = SqlText & "customer_address_id LONG, "
    SqlText = SqlText & "invoice_address_id LONG, "
    SqlText = SqlText & "delivery_address_id LONG, "
    SqlText = SqlText & "customer_name TEXT(150), "
    SqlText = SqlText & "order_type_code TEXT(30), "
    SqlText = SqlText & "order_status_code TEXT(30), "
    SqlText = SqlText & "order_date DATETIME, "
    SqlText = SqlText & "delivery_date DATETIME, "
    SqlText = SqlText & "valid_until DATETIME, "
    SqlText = SqlText & "reference_text TEXT(150), "
    SqlText = SqlText & "external_reference TEXT(150), "
    SqlText = SqlText & "language_code TEXT(10), "
    SqlText = SqlText & "currency_code TEXT(10), "
    SqlText = SqlText & "payment_term_code TEXT(50), "
    SqlText = SqlText & "vat_mode TEXT(20), "
    SqlText = SqlText & "vat_code TEXT(30), "
    SqlText = SqlText & "vat_rate DOUBLE, "
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

Private Sub CreateTmpOrderLines(ByVal db As DAO.Database)
    Dim SqlText As String

    SqlText = ""
    SqlText = SqlText & "CREATE TABLE tmp_order_line ("
    SqlText = SqlText & "tmp_order_line_id AUTOINCREMENT CONSTRAINT pk_tmp_order_line PRIMARY KEY, "
    SqlText = SqlText & "order_line_id LONG, "
    SqlText = SqlText & "order_id LONG, "
    SqlText = SqlText & "tmp_order_id LONG NOT NULL, "
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
    SqlText = SqlText & "sort_order LONG, "
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
    CreateTmpOrders db
    CreateTmpOrderLines db

    If Not modDbSchema.TableExists(db, "ord_order") Then Exit Function
    If Not modDbSchema.TableExists(db, "ord_order_line") Then Exit Function
    If Not modDbSchema.TableExists(db, "tmp_order") Then Exit Function
    If Not modDbSchema.TableExists(db, "tmp_order_line") Then Exit Function

    If Not EnsureOrderHeaderSchema(db) Then GoTo CleanExit
    If Not EnsureOrderLineSchema(db) Then GoTo CleanExit
    If Not EnsureTemporaryOrderHeaderSchema(db) Then GoTo CleanExit
    If Not EnsureTemporaryOrderLineSchema(db) Then GoTo CleanExit
    If Not CleanupLegacyOrderSchema(db) Then GoTo CleanExit
    If Not VerifyRequiredOrderFields(db) Then GoTo CleanExit
    If Not EnsureOrderHeaderIndexes(db) Then GoTo CleanExit
    If Not EnsureOrderLineIndexes(db) Then GoTo CleanExit
    If Not EnsureTemporaryOrderHeaderIndexes(db) Then GoTo CleanExit
    If Not EnsureTemporaryOrderLineIndexes(db) Then GoTo CleanExit

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
    If Not EnsureLongField(db, "ord_order", "invoice_address_id", 0, False) Then GoTo ErrorHandler
    If Not EnsureLongField(db, "ord_order", "delivery_address_id", 0, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "customer_name", 150, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "ord_order", "order_date", Date, True) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "ord_order", "delivery_date", 0, False) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "ord_order", "valid_until", 0, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "reference_text", 150, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "external_reference", 150, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "language_code", 10, "en-US", True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "currency_code", 10, "CHF", True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "payment_term_code", 50, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "vat_mode", 20, "EXCLUSIVE", True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order", "vat_code", 30, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureDoubleField(db, "ord_order", "vat_rate", 0, True) Then GoTo ErrorHandler
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
    If Not EnsureLongField(db, "ord_order_line", "sort_order", 0, False) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "ord_order_line", "created_at", Now(), True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order_line", "created_by", 50, "SYSTEM", True) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "ord_order_line", "updated_at", Now(), True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "ord_order_line", "updated_by", 50, "SYSTEM", True) Then GoTo ErrorHandler

    EnsureOrderLineSchema = True
    Exit Function

ErrorHandler:
    EnsureOrderLineSchema = False
End Function

Private Function EnsureTemporaryOrderHeaderSchema(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    If Not EnsureTextField(db, "tmp_order", "session_id", 100, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureLongField(db, "tmp_order", "order_id", 0, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order", "order_no", 50, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureLongField(db, "tmp_order", "customer_address_id", 0, False) Then GoTo ErrorHandler
    If Not EnsureLongField(db, "tmp_order", "invoice_address_id", 0, False) Then GoTo ErrorHandler
    If Not EnsureLongField(db, "tmp_order", "delivery_address_id", 0, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order", "customer_name", 150, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order", "order_type_code", 30, "SO", True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order", "order_status_code", 30, "DRAFT", True) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "tmp_order", "order_date", Date, True) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "tmp_order", "delivery_date", 0, False) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "tmp_order", "valid_until", 0, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order", "reference_text", 150, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order", "external_reference", 150, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order", "language_code", 10, "en-US", True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order", "currency_code", 10, "CHF", True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order", "payment_term_code", 50, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order", "vat_mode", 20, "EXCLUSIVE", True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order", "vat_code", 30, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureDoubleField(db, "tmp_order", "vat_rate", 0, True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order", "header_discount_type", 20, "NONE", True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order", "header_discount_value", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order", "header_discount_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order", "header_surcharge_type", 20, "NONE", True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order", "header_surcharge_value", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order", "header_surcharge_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order", "subtotal_net_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order", "net_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order", "vat_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order", "gross_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureLongTextField(db, "tmp_order", "notes_text") Then GoTo ErrorHandler
    If Not EnsureLongTextField(db, "tmp_order", "internal_notes_text") Then GoTo ErrorHandler
    If Not EnsureLongField(db, "tmp_order", "result_document_id", 0, False) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "tmp_order", "created_at", Now(), True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order", "created_by", 50, "SYSTEM", True) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "tmp_order", "updated_at", Now(), True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order", "updated_by", 50, "SYSTEM", True) Then GoTo ErrorHandler

    EnsureTemporaryOrderHeaderSchema = True
    Exit Function

ErrorHandler:
    EnsureTemporaryOrderHeaderSchema = False
End Function

Private Function EnsureTemporaryOrderLineSchema(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    If Not EnsureLongField(db, "tmp_order_line", "order_line_id", 0, False) Then GoTo ErrorHandler
    If Not EnsureLongField(db, "tmp_order_line", "order_id", 0, False) Then GoTo ErrorHandler
    If Not EnsureLongField(db, "tmp_order_line", "tmp_order_id", 0, False) Then GoTo ErrorHandler
    If Not EnsureLongField(db, "tmp_order_line", "line_no", 0, False) Then GoTo ErrorHandler
    If Not EnsureLongField(db, "tmp_order_line", "article_id", 0, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order_line", "article_no", 50, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order_line", "line_type_code", 30, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureLongTextField(db, "tmp_order_line", "description_text") Then GoTo ErrorHandler
    If Not EnsureDoubleField(db, "tmp_order_line", "quantity", 0, True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order_line", "unit_code", 30, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order_line", "unit_price", 0, True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order_line", "discount_type", 20, "NONE", True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order_line", "discount_value", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order_line", "line_discount_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order_line", "surcharge_type", 20, "NONE", True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order_line", "surcharge_value", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order_line", "line_surcharge_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order_line", "vat_code", 30, vbNullString, False) Then GoTo ErrorHandler
    If Not EnsureDoubleField(db, "tmp_order_line", "vat_rate", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order_line", "line_base_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order_line", "line_net_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order_line", "line_vat_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureCurrencyField(db, "tmp_order_line", "line_gross_amount", 0, True) Then GoTo ErrorHandler
    If Not EnsureLongField(db, "tmp_order_line", "sort_order", 0, False) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "tmp_order_line", "created_at", Now(), True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order_line", "created_by", 50, "SYSTEM", True) Then GoTo ErrorHandler
    If Not EnsureDateField(db, "tmp_order_line", "updated_at", Now(), True) Then GoTo ErrorHandler
    If Not EnsureTextField(db, "tmp_order_line", "updated_by", 50, "SYSTEM", True) Then GoTo ErrorHandler

    EnsureTemporaryOrderLineSchema = True
    Exit Function

ErrorHandler:
    EnsureTemporaryOrderLineSchema = False
End Function

Private Function VerifyRequiredOrderFields(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    VerifyRequiredOrderFields = False

    If Not EnsureRequiredFieldExists(db, "ord_order", "customer_address_id") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "ord_order", "invoice_address_id") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "ord_order", "delivery_address_id") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "ord_order", "vat_code") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "ord_order", "vat_rate") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "ord_order_line", "article_no") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "ord_order_line", "vat_rate") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "tmp_order", "session_id") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "tmp_order", "order_no") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "tmp_order_line", "tmp_order_id") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "tmp_order_line", "order_line_id") Then Exit Function
    If Not EnsureRequiredFieldExists(db, "tmp_order_line", "vat_rate") Then Exit Function

    VerifyRequiredOrderFields = True
    Exit Function

ErrorHandler:
    VerifyRequiredOrderFields = False
End Function

Private Function CleanupLegacyOrderSchema(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    CleanupLegacyOrderSchema = False

    DropFieldIfExists db, "ord_order", "address_id"
    DropFieldIfExists db, "ord_order", "subtotal_amount"
    DropFieldIfExists db, "ord_order", "total_amount"
    DropFieldIfExists db, "ord_order", "tenant_code"

    DropFieldIfExists db, "ord_order_line", "line_description"
    DropFieldIfExists db, "ord_order_line", "line_total_net"
    DropFieldIfExists db, "ord_order_line", "line_total_vat"
    DropFieldIfExists db, "ord_order_line", "line_total_gross"

    CleanupLegacyOrderSchema = True
    Exit Function

ErrorHandler:
    CleanupLegacyOrderSchema = False
End Function

Private Function EnsureOrderHeaderIndexes(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    EnsureIndexWhenFieldExists db, "ord_order", "order_no", "ux_ord_order_order_no", _
        "CREATE UNIQUE INDEX ux_ord_order_order_no ON ord_order (order_no);"
    EnsureIndexWhenFieldExists db, "ord_order", "customer_address_id", "ix_ord_order_customer_address_id", _
        "CREATE INDEX ix_ord_order_customer_address_id ON ord_order (customer_address_id);"
    EnsureIndexWhenFieldExists db, "ord_order", "invoice_address_id", "ix_ord_order_invoice_address_id", _
        "CREATE INDEX ix_ord_order_invoice_address_id ON ord_order (invoice_address_id);"
    EnsureIndexWhenFieldExists db, "ord_order", "delivery_address_id", "ix_ord_order_delivery_address_id", _
        "CREATE INDEX ix_ord_order_delivery_address_id ON ord_order (delivery_address_id);"
    EnsureIndexWhenFieldExists db, "ord_order", "order_date", "ix_ord_order_order_date", _
        "CREATE INDEX ix_ord_order_order_date ON ord_order (order_date);"
    EnsureIndexWhenFieldExists db, "ord_order", "order_status_code", "ix_ord_order_order_status_code", _
        "CREATE INDEX ix_ord_order_order_status_code ON ord_order (order_status_code);"
    EnsureIndexWhenFieldExists db, "ord_order", "payment_term_code", "ix_ord_order_payment_term_code", _
        "CREATE INDEX ix_ord_order_payment_term_code ON ord_order (payment_term_code);"
    EnsureIndexWhenFieldExists db, "ord_order", "vat_code", "ix_ord_order_vat_code", _
        "CREATE INDEX ix_ord_order_vat_code ON ord_order (vat_code);"
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

Private Function EnsureTemporaryOrderHeaderIndexes(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    EnsureIndexWhenFieldExists db, "tmp_order", "session_id", "ix_tmp_order_session_id", _
        "CREATE INDEX ix_tmp_order_session_id ON tmp_order (session_id);"
    EnsureIndexWhenFieldExists db, "tmp_order", "customer_address_id", "ix_tmp_order_customer_address_id", _
        "CREATE INDEX ix_tmp_order_customer_address_id ON tmp_order (customer_address_id);"

    EnsureTemporaryOrderHeaderIndexes = True
    Exit Function

ErrorHandler:
    EnsureTemporaryOrderHeaderIndexes = False
End Function

Private Function EnsureTemporaryOrderLineIndexes(ByVal db As DAO.Database) As Boolean
    On Error GoTo ErrorHandler

    EnsureIndexWhenFieldExists db, "tmp_order_line", "tmp_order_id", "ix_tmp_order_line_tmp_order_id", _
        "CREATE INDEX ix_tmp_order_line_tmp_order_id ON tmp_order_line (tmp_order_id);"
    EnsureIndexWhenFieldExists db, "tmp_order_line", "line_no", "ix_tmp_order_line_line_no", _
        "CREATE INDEX ix_tmp_order_line_line_no ON tmp_order_line (line_no);"
    EnsureIndexWhenFieldExists db, "tmp_order_line", "vat_code", "ix_tmp_order_line_vat_code", _
        "CREATE INDEX ix_tmp_order_line_vat_code ON tmp_order_line (vat_code);"

    EnsureTemporaryOrderLineIndexes = True
    Exit Function

ErrorHandler:
    EnsureTemporaryOrderLineIndexes = False
End Function



Private Function EnsureTextField(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String, ByVal FieldSize As Long, ByVal defaultValue As String, ByVal updateNullValues As Boolean) As Boolean
    On Error GoTo ErrorHandler

    If Not modDbSchema.FieldExists(db, tableName, fieldName) Then
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

    If Not modDbSchema.FieldExists(db, tableName, fieldName) Then
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureLongField", _
            "AddField executing: " & tableName & "." & fieldName & "; db=" & ResolveDatabasePath(db)
        db.Execute "ALTER TABLE [" & tableName & "] ADD COLUMN [" & fieldName & "] LONG;", dbFailOnError
        RefreshTableDefinition db, tableName
        LogTableFieldNames db, tableName, MODULE_NAME & ".EnsureLongField"
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureLongField", _
            "AddField successful: " & tableName & "." & fieldName & "; exists_after_add=" & CStr(modDbSchema.FieldExists(db, tableName, fieldName))
        If Not modDbSchema.FieldExists(db, tableName, fieldName) Then
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

    If Not modDbSchema.FieldExists(db, tableName, fieldName) Then
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

    If Not modDbSchema.FieldExists(db, tableName, fieldName) Then
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

    If Not modDbSchema.FieldExists(db, tableName, fieldName) Then
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

    If Not modDbSchema.FieldExists(db, tableName, fieldName) Then
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

    If modDbSchema.FieldExists(db, tableName, fieldName) Then
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

Public Sub DropFieldIfExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String)
    On Error GoTo ErrorHandler

    If db Is Nothing Then
        Exit Sub
    End If

    If Not modDbSchema.FieldExists(db, tableName, fieldName) Then
        Exit Sub
    End If

    db.Execute "ALTER TABLE [" & tableName & "] DROP COLUMN [" & fieldName & "];", dbFailOnError
    modLoggingHandler.LogInfo MODULE_NAME & ".DropFieldIfExists", "Dropped field: " & tableName & "." & fieldName
    Exit Sub

ErrorHandler:
    modLoggingHandler.LogWarning MODULE_NAME & ".DropFieldIfExists", _
        "Could not drop field " & tableName & "." & fieldName & " (" & Err.Number & " - " & Err.description & ")"
End Sub

Private Sub EnsureIndexWhenFieldExists( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal fieldName As String, _
    ByVal indexName As String, _
    ByVal SqlText As String)
    On Error GoTo ErrorHandler

    If Not modDbSchema.FieldExists(db, tableName, fieldName) Then
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

Private Function OpenOrCreateAccessDatabase(ByVal databasePath As String) As DAO.Database
    On Error GoTo ErrorHandler

    If LenB(Dir$(databasePath, vbNormal)) = 0 Then
        Set OpenOrCreateAccessDatabase = DBEngine.CreateDatabase(databasePath, dbLangGeneral)
    Else
        Set OpenOrCreateAccessDatabase = DBEngine.OpenDatabase(databasePath)
    End If
    Exit Function

ErrorHandler:
    Set OpenOrCreateAccessDatabase = Nothing
    Err.Raise Err.Number, MODULE_NAME & ".OpenOrCreateAccessDatabase", Err.description
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

Private Function SqlText(ByVal valueText As String) As String
    SqlText = "'" & Replace(Trim$(valueText), "'", "''") & "'"
End Function



