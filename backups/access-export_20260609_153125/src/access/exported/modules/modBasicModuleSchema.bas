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

Public Sub CreateBasicModuleTables(Optional ByVal backendPath As String = vbNullString)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database

    If LenB(Trim$(backendPath)) > 0 Then
        Set db = DBEngine.OpenDatabase(backendPath)
    Else
        Set db = currentDb
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

    MsgBox "BasicModule-Tabellen wurden erstellt.", vbInformation, MODULE_NAME

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then
        If LenB(Trim$(backendPath)) > 0 Then db.Close
    End If
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Erstellen der BasicModule-Tabellen: " & Err.description, vbExclamation, MODULE_NAME
    Resume CleanExit
End Sub

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
    SqlText = SqlText & "address_id LONG, "
    SqlText = SqlText & "order_date DATETIME, "
    SqlText = SqlText & "delivery_date DATETIME, "
    SqlText = SqlText & "valid_until DATETIME, "
    SqlText = SqlText & "reference_text TEXT(150), "
    SqlText = SqlText & "language_code TEXT(10), "
    SqlText = SqlText & "currency_code TEXT(3), "
    SqlText = SqlText & "payment_term_code TEXT(50), "
    SqlText = SqlText & "subtotal_net_amount CURRENCY, "
    SqlText = SqlText & "total_discount_amount CURRENCY, "
    SqlText = SqlText & "total_surcharge_amount CURRENCY, "
    SqlText = SqlText & "total_vat_amount CURRENCY, "
    SqlText = SqlText & "total_gross_amount CURRENCY, "
    SqlText = SqlText & "notes_text LONGTEXT, "
    SqlText = SqlText & "internal_notes_text LONGTEXT, "
    SqlText = SqlText & "created_at DATETIME, "
    SqlText = SqlText & "created_by TEXT(50), "
    SqlText = SqlText & "updated_at DATETIME, "
    SqlText = SqlText & "updated_by TEXT(50)"
    SqlText = SqlText & ");"

    ExecuteDdl db, SqlText
    ExecuteCreateIndexIfMissing db, "ord_order", "ux_ord_order_order_no", _
        "CREATE UNIQUE INDEX ux_ord_order_order_no ON ord_order (order_no);"
    ExecuteCreateIndexIfMissing db, "ord_order", "ix_ord_order_address_id", _
        "CREATE INDEX ix_ord_order_address_id ON ord_order (address_id);"
    ExecuteCreateIndexIfMissing db, "ord_order", "ix_ord_order_order_date", _
        "CREATE INDEX ix_ord_order_order_date ON ord_order (order_date);"
    ExecuteCreateIndexIfMissing db, "ord_order", "ix_ord_order_order_status_code", _
        "CREATE INDEX ix_ord_order_order_status_code ON ord_order (order_status_code);"
    ExecuteCreateIndexIfMissing db, "ord_order", "ix_ord_order_payment_term_code", _
        "CREATE INDEX ix_ord_order_payment_term_code ON ord_order (payment_term_code);"
End Sub

Private Sub CreateTblOrderLines(ByVal db As DAO.Database)
    Dim SqlText As String

    SqlText = ""
    SqlText = SqlText & "CREATE TABLE ord_order_line ("
    SqlText = SqlText & "order_line_id AUTOINCREMENT CONSTRAINT pk_ord_order_line PRIMARY KEY, "
    SqlText = SqlText & "order_id LONG NOT NULL, "
    SqlText = SqlText & "line_no LONG, "
    SqlText = SqlText & "article_id LONG, "
    SqlText = SqlText & "line_type_code TEXT(30), "
    SqlText = SqlText & "description_text LONGTEXT, "
    SqlText = SqlText & "quantity DOUBLE, "
    SqlText = SqlText & "unit_code TEXT(30), "
    SqlText = SqlText & "unit_price CURRENCY, "
    SqlText = SqlText & "discount_percent DOUBLE, "
    SqlText = SqlText & "discount_amount CURRENCY, "
    SqlText = SqlText & "surcharge_percent DOUBLE, "
    SqlText = SqlText & "surcharge_amount CURRENCY, "
    SqlText = SqlText & "vat_code TEXT(30), "
    SqlText = SqlText & "vat_rate DOUBLE, "
    SqlText = SqlText & "line_total CURRENCY, "
    SqlText = SqlText & "created_at DATETIME, "
    SqlText = SqlText & "created_by TEXT(50), "
    SqlText = SqlText & "updated_at DATETIME, "
    SqlText = SqlText & "updated_by TEXT(50)"
    SqlText = SqlText & ");"

    ExecuteDdl db, SqlText
    ExecuteCreateIndexIfMissing db, "ord_order_line", "ix_ord_order_line_order_id", _
        "CREATE INDEX ix_ord_order_line_order_id ON ord_order_line (order_id);"
    ExecuteCreateIndexIfMissing db, "ord_order_line", "ix_ord_order_line_article_id", _
        "CREATE INDEX ix_ord_order_line_article_id ON ord_order_line (article_id);"
    ExecuteCreateIndexIfMissing db, "ord_order_line", "ix_ord_order_line_line_no", _
        "CREATE INDEX ix_ord_order_line_line_no ON ord_order_line (line_no);"
    ExecuteCreateIndexIfMissing db, "ord_order_line", "ix_ord_order_line_unit_code", _
        "CREATE INDEX ix_ord_order_line_unit_code ON ord_order_line (unit_code);"
    ExecuteCreateIndexIfMissing db, "ord_order_line", "ix_ord_order_line_vat_code", _
        "CREATE INDEX ix_ord_order_line_vat_code ON ord_order_line (vat_code);"
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