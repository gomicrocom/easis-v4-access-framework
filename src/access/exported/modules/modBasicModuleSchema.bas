Attribute VB_Name = "modBasicModuleSchema"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modBasicModuleSchema
' Purpose   : Creates BasicModule v1 tables for addresses, articles and orders.
' Author    : Codex
' Version   : 0.1.1
'===============================================================================

Private Const MODULE_NAME As String = "modBasicModuleSchema"

Public Sub CreateBasicModuleTables(Optional ByVal BackendPath As String = vbNullString)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database

    If LenB(Trim$(BackendPath)) > 0 Then
        Set db = DBEngine.OpenDatabase(BackendPath)
    Else
        Set db = CurrentDb
    End If

    CreateRefPaymentTerms db
    CreateRefVatCodes db
    CreateRefUnits db

    CreateTblAddresses db
    CreateTblProductGroups db
    CreateTblArticles db

    CreateTblOrders db
    CreateTblOrderLines db

    MsgBox "BasicModule-Tabellen wurden erstellt.", vbInformation, MODULE_NAME

CleanExit:
    On Error Resume Next
    If Not db Is Nothing Then
        If LenB(Trim$(BackendPath)) > 0 Then db.Close
    End If
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Erstellen der BasicModule-Tabellen: " & Err.description, vbExclamation, MODULE_NAME
    Resume CleanExit
End Sub

Private Sub CreateRefPaymentTerms(ByVal db As DAO.Database)
    Dim SqlText As String

    SqlText = ""
    SqlText = SqlText & "CREATE TABLE ref_payment_term ("
    SqlText = SqlText & "PaymentTermId AUTOINCREMENT CONSTRAINT pk_ref_payment_term PRIMARY KEY, "
    SqlText = SqlText & "PaymentTermCode TEXT(30) NOT NULL, "
    SqlText = SqlText & "PaymentTermName TEXT(100), "
    SqlText = SqlText & "DueDays LONG, "
    SqlText = SqlText & "CashDiscountDays LONG, "
    SqlText = SqlText & "CashDiscountPercent DOUBLE, "
    SqlText = SqlText & "IsActive YESNO, "
    SqlText = SqlText & "SortOrder LONG"
    SqlText = SqlText & ");"

    ExecuteDdl db, SqlText
    ExecuteDdl db, "CREATE UNIQUE INDEX ux_ref_payment_term_Code ON ref_payment_term (PaymentTermCode);"
End Sub

Private Sub CreateRefVatCodes(ByVal db As DAO.Database)
    Dim SqlText As String

    SqlText = ""
    SqlText = SqlText & "CREATE TABLE ref_vat_code ("
    SqlText = SqlText & "VatCode TEXT(20) CONSTRAINT pk_ref_vat_code PRIMARY KEY, "
    SqlText = SqlText & "VatName TEXT(100), "
    SqlText = SqlText & "VatRate DOUBLE, "
    SqlText = SqlText & "CountryCode TEXT(2), "
    SqlText = SqlText & "IsActive YESNO, "
    SqlText = SqlText & "ValidFrom DATETIME, "
    SqlText = SqlText & "ValidTo DATETIME"
    SqlText = SqlText & ");"

    ExecuteDdl db, SqlText
End Sub

Private Sub CreateRefUnits(ByVal db As DAO.Database)
    Dim SqlText As String

    SqlText = ""
    SqlText = SqlText & "CREATE TABLE ref_unit ("
    SqlText = SqlText & "UnitCode TEXT(20) CONSTRAINT pk_ref_unit PRIMARY KEY, "
    SqlText = SqlText & "UnitName TEXT(100), "
    SqlText = SqlText & "SortOrder LONG, "
    SqlText = SqlText & "IsActive YESNO"
    SqlText = SqlText & ");"

    ExecuteDdl db, SqlText
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
    SqlText = SqlText & "PaymentTermId LONG, "
    SqlText = SqlText & "IsActive YESNO, "
    SqlText = SqlText & "CreatedAt DATETIME, "
    SqlText = SqlText & "CreatedBy TEXT(50), "
    SqlText = SqlText & "UpdatedAt DATETIME, "
    SqlText = SqlText & "UpdatedBy TEXT(50)"
    SqlText = SqlText & ");"

    ExecuteDdl db, SqlText
    ExecuteDdl db, "CREATE INDEX ix_tblAddresses_AddressType ON tblAddresses (AddressType);"
    ExecuteDdl db, "CREATE INDEX ix_tblAddresses_CompanyName ON tblAddresses (CompanyName);"
    ExecuteDdl db, "CREATE INDEX ix_tblAddresses_PaymentTermId ON tblAddresses (PaymentTermId);"
End Sub

Private Sub CreateTblProductGroups(ByVal db As DAO.Database)
    Dim SqlText As String

    SqlText = ""
    SqlText = SqlText & "CREATE TABLE art_product_group ("
    SqlText = SqlText & "ProductGroupId AUTOINCREMENT CONSTRAINT pk_art_product_group PRIMARY KEY, "
    SqlText = SqlText & "ProductGroupCode TEXT(30) NOT NULL, "
    SqlText = SqlText & "ProductGroupName TEXT(100), "
    SqlText = SqlText & "RevenueAccount TEXT(20), "
    SqlText = SqlText & "ExpenseAccount TEXT(20), "
    SqlText = SqlText & "VatCode TEXT(20), "
    SqlText = SqlText & "IsActive YESNO, "
    SqlText = SqlText & "SortOrder LONG, "
    SqlText = SqlText & "CreatedAt DATETIME, "
    SqlText = SqlText & "CreatedBy TEXT(50), "
    SqlText = SqlText & "UpdatedAt DATETIME, "
    SqlText = SqlText & "UpdatedBy TEXT(50)"
    SqlText = SqlText & ");"

    ExecuteDdl db, SqlText
    ExecuteDdl db, "CREATE UNIQUE INDEX ux_art_product_group_Code ON art_product_group (ProductGroupCode);"
    ExecuteDdl db, "CREATE INDEX ix_art_product_group_VatCode ON art_product_group (VatCode);"
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
    SqlText = SqlText & "UnitCode TEXT(20), "
    SqlText = SqlText & "SalesPrice CURRENCY, "
    SqlText = SqlText & "PurchasePrice CURRENCY, "
    SqlText = SqlText & "CurrencyCode TEXT(3), "
    SqlText = SqlText & "VatCode TEXT(20), "
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
    ExecuteDdl db, "CREATE INDEX ix_art_article_UnitCode ON art_article (UnitCode);"
    ExecuteDdl db, "CREATE INDEX ix_art_article_VatCode ON art_article (VatCode);"
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
    SqlText = SqlText & "PaymentTermId LONG, "
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
    ExecuteDdl db, "CREATE INDEX ix_ord_order_PaymentTermId ON ord_order (PaymentTermId);"
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
    SqlText = SqlText & "UnitCode TEXT(20), "
    SqlText = SqlText & "UnitPrice CURRENCY, "
    SqlText = SqlText & "DiscountPercent DOUBLE, "
    SqlText = SqlText & "DiscountAmount CURRENCY, "
    SqlText = SqlText & "SurchargePercent DOUBLE, "
    SqlText = SqlText & "SurchargeAmount CURRENCY, "
    SqlText = SqlText & "VatCode TEXT(20), "
    SqlText = SqlText & "VatRate DOUBLE, "
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
    ExecuteDdl db, "CREATE INDEX ix_ord_order_line_VatCode ON ord_order_line (VatCode);"
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
