Option Compare Database
Option Explicit

'===============================================================================
' Module    : modBasicModuleSchema
' Purpose   : Creates BasicModule v1 tables for addresses, articles and orders.
' Author    : Codex
' Version   : 0.1.7
'===============================================================================

Private Const MODULE_NAME As String = "modBasicModuleSchema"

Public Sub CreateBasicModuleTables(Optional ByVal backendPath As String = vbNullString)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database

    If LenB(Trim$(backendPath)) > 0 Then
        Set db = DBEngine.OpenDatabase(backendPath)
    Else
        Set db = CurrentDb
    End If

    CreateTenPaymentTerms db
    CreateRefVatCodes db
    CreateRefUnits db
    CreateRefAddressType db
    CreateRefSalutation db
    CreateRefAddressingMode db
    CreateRefContactType db

    CreateTblAddresses db
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
    ExecuteDdl db, "CREATE UNIQUE INDEX ux_ten_payment_term_code_language ON ten_payment_term (payment_term_code, language_code);"
    ExecuteDdl db, "CREATE INDEX ix_ten_payment_term_is_default ON ten_payment_term (is_default);"
    ExecuteDdl db, "CREATE INDEX ix_ten_payment_term_is_active ON ten_payment_term (is_active);"
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

Private Sub CreateTblAddresses(ByVal db As DAO.Database)
    Dim SqlText As String

    SqlText = ""
    SqlText = SqlText & "CREATE TABLE tblAddresses ("
    SqlText = SqlText & "AddressId AUTOINCREMENT CONSTRAINT pk_tblAddresses PRIMARY KEY, "
    SqlText = SqlText & "AddressType TEXT(30), "
    SqlText = SqlText & "CompanyName TEXT(150), "
    SqlText = SqlText & "FirstName TEXT(80), "
    SqlText = SqlText & "LastName TEXT(80), "
    SqlText = SqlText & "Street TEXT(120), "
    SqlText = SqlText & "HouseNo TEXT(20), "
    SqlText = SqlText & "PostalCode TEXT(20), "
    SqlText = SqlText & "City TEXT(100), "
    SqlText = SqlText & "CountryCode TEXT(2), "
    SqlText = SqlText & "Email TEXT(150), "
    SqlText = SqlText & "Phone TEXT(50), "
    SqlText = SqlText & "VatNo TEXT(50), "
    SqlText = SqlText & "LanguageCode TEXT(10), "
    SqlText = SqlText & "CurrencyCode TEXT(3), "
    SqlText = SqlText & "payment_term_code TEXT(50), "
    SqlText = SqlText & "IsActive YESNO, "
    SqlText = SqlText & "CreatedAt DATETIME, "
    SqlText = SqlText & "CreatedBy TEXT(50), "
    SqlText = SqlText & "UpdatedAt DATETIME, "
    SqlText = SqlText & "UpdatedBy TEXT(50)"
    SqlText = SqlText & ");"

    ExecuteDdl db, SqlText
    ExecuteDdl db, "CREATE INDEX ix_tblAddresses_AddressType ON tblAddresses (AddressType);"
    ExecuteDdl db, "CREATE INDEX ix_tblAddresses_CompanyName ON tblAddresses (CompanyName);"
    ExecuteDdl db, "CREATE INDEX ix_tblAddresses_payment_term_code ON tblAddresses (payment_term_code);"
End Sub

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
    ExecuteDdl db, "CREATE UNIQUE INDEX ux_art_product_group_code ON art_product_group (product_group_code);"
    ExecuteDdl db, "CREATE INDEX ix_art_product_group_sort_order ON art_product_group (sort_order);"
    ExecuteDdl db, "CREATE INDEX ix_art_product_group_is_active ON art_product_group (is_active);"
End Sub

Private Sub CreateTblArticles(ByVal db As DAO.Database)
    Dim SqlText As String

    SqlText = ""
    SqlText = SqlText & "CREATE TABLE art_article ("
    SqlText = SqlText & "ArticleId AUTOINCREMENT CONSTRAINT pk_art_article PRIMARY KEY, "
    SqlText = SqlText & "ArticleNo TEXT(50) NOT NULL, "
    SqlText = SqlText & "ArticleName TEXT(150), "
    SqlText = SqlText & "Description LONGTEXT, "
    SqlText = SqlText & "ProductGroupId LONG, "
    SqlText = SqlText & "unit_code TEXT(30), "
    SqlText = SqlText & "SalesPrice CURRENCY, "
    SqlText = SqlText & "PurchasePrice CURRENCY, "
    SqlText = SqlText & "CurrencyCode TEXT(3), "
    SqlText = SqlText & "vat_code TEXT(30), "
    SqlText = SqlText & "IsStockArticle YESNO, "
    SqlText = SqlText & "IsServiceArticle YESNO, "
    SqlText = SqlText & "IsActive YESNO, "
    SqlText = SqlText & "CreatedAt DATETIME, "
    SqlText = SqlText & "CreatedBy TEXT(50), "
    SqlText = SqlText & "UpdatedAt DATETIME, "
    SqlText = SqlText & "UpdatedBy TEXT(50)"
    SqlText = SqlText & ");"

    ExecuteDdl db, SqlText
    ExecuteDdl db, "CREATE UNIQUE INDEX ux_art_article_ArticleNo ON art_article (ArticleNo);"
    ExecuteDdl db, "CREATE INDEX ix_art_article_ProductGroupId ON art_article (ProductGroupId);"
    ExecuteDdl db, "CREATE INDEX ix_art_article_unit_code ON art_article (unit_code);"
    ExecuteDdl db, "CREATE INDEX ix_art_article_vat_code ON art_article (vat_code);"
End Sub

Private Sub CreateTblOrders(ByVal db As DAO.Database)
    Dim SqlText As String

    SqlText = ""
    SqlText = SqlText & "CREATE TABLE ord_order ("
    SqlText = SqlText & "OrderId AUTOINCREMENT CONSTRAINT pk_ord_order PRIMARY KEY, "
    SqlText = SqlText & "OrderNo TEXT(50), "
    SqlText = SqlText & "OrderType TEXT(30), "
    SqlText = SqlText & "OrderStatus TEXT(30), "
    SqlText = SqlText & "CustomerAddressId LONG, "
    SqlText = SqlText & "OrderDate DATETIME, "
    SqlText = SqlText & "DeliveryDate DATETIME, "
    SqlText = SqlText & "ValidUntil DATETIME, "
    SqlText = SqlText & "ReferenceText TEXT(150), "
    SqlText = SqlText & "LanguageCode TEXT(10), "
    SqlText = SqlText & "CurrencyCode TEXT(3), "
    SqlText = SqlText & "payment_term_code TEXT(50), "
    SqlText = SqlText & "SubtotalNet CURRENCY, "
    SqlText = SqlText & "TotalDiscount CURRENCY, "
    SqlText = SqlText & "TotalSurcharge CURRENCY, "
    SqlText = SqlText & "TotalVat CURRENCY, "
    SqlText = SqlText & "TotalGross CURRENCY, "
    SqlText = SqlText & "Notes LONGTEXT, "
    SqlText = SqlText & "InternalNotes LONGTEXT, "
    SqlText = SqlText & "CreatedAt DATETIME, "
    SqlText = SqlText & "CreatedBy TEXT(50), "
    SqlText = SqlText & "UpdatedAt DATETIME, "
    SqlText = SqlText & "UpdatedBy TEXT(50)"
    SqlText = SqlText & ");"

    ExecuteDdl db, SqlText
    ExecuteDdl db, "CREATE UNIQUE INDEX ux_ord_order_OrderNo ON ord_order (OrderNo);"
    ExecuteDdl db, "CREATE INDEX ix_ord_order_CustomerAddressId ON ord_order (CustomerAddressId);"
    ExecuteDdl db, "CREATE INDEX ix_ord_order_OrderDate ON ord_order (OrderDate);"
    ExecuteDdl db, "CREATE INDEX ix_ord_order_OrderStatus ON ord_order (OrderStatus);"
    ExecuteDdl db, "CREATE INDEX ix_ord_order_payment_term_code ON ord_order (payment_term_code);"
End Sub

Private Sub CreateTblOrderLines(ByVal db As DAO.Database)
    Dim SqlText As String

    SqlText = ""
    SqlText = SqlText & "CREATE TABLE ord_order_line ("
    SqlText = SqlText & "OrderLineId AUTOINCREMENT CONSTRAINT pk_ord_order_line PRIMARY KEY, "
    SqlText = SqlText & "OrderId LONG NOT NULL, "
    SqlText = SqlText & "LineNo LONG, "
    SqlText = SqlText & "ArticleId LONG, "
    SqlText = SqlText & "LineType TEXT(30), "
    SqlText = SqlText & "Description LONGTEXT, "
    SqlText = SqlText & "Quantity DOUBLE, "
    SqlText = SqlText & "unit_code TEXT(30), "
    SqlText = SqlText & "UnitPrice CURRENCY, "
    SqlText = SqlText & "DiscountPercent DOUBLE, "
    SqlText = SqlText & "DiscountAmount CURRENCY, "
    SqlText = SqlText & "SurchargePercent DOUBLE, "
    SqlText = SqlText & "SurchargeAmount CURRENCY, "
    SqlText = SqlText & "vat_code TEXT(30), "
    SqlText = SqlText & "vat_rate DOUBLE, "
    SqlText = SqlText & "LineNetAmount CURRENCY, "
    SqlText = SqlText & "LineVatAmount CURRENCY, "
    SqlText = SqlText & "LineGrossAmount CURRENCY, "
    SqlText = SqlText & "SortOrder LONG, "
    SqlText = SqlText & "CreatedAt DATETIME, "
    SqlText = SqlText & "CreatedBy TEXT(50), "
    SqlText = SqlText & "UpdatedAt DATETIME, "
    SqlText = SqlText & "UpdatedBy TEXT(50)"
    SqlText = SqlText & ");"

    ExecuteDdl db, SqlText
    ExecuteDdl db, "CREATE INDEX ix_ord_order_line_OrderId ON ord_order_line (OrderId);"
    ExecuteDdl db, "CREATE INDEX ix_ord_order_line_ArticleId ON ord_order_line (ArticleId);"
    ExecuteDdl db, "CREATE INDEX ix_ord_order_line_LineNo ON ord_order_line (LineNo);"
    ExecuteDdl db, "CREATE INDEX ix_ord_order_line_unit_code ON ord_order_line (unit_code);"
    ExecuteDdl db, "CREATE INDEX ix_ord_order_line_vat_code ON ord_order_line (vat_code);"
End Sub

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
