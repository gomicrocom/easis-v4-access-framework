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
    Dim sqlText As String

    sqlText = ""
    sqlText = sqlText & "CREATE TABLE refPaymentTerms ("
    sqlText = sqlText & "PaymentTermId AUTOINCREMENT CONSTRAINT pk_refPaymentTerms PRIMARY KEY, "
    sqlText = sqlText & "PaymentTermCode TEXT(30) NOT NULL, "
    sqlText = sqlText & "PaymentTermName TEXT(100), "
    sqlText = sqlText & "DueDays LONG, "
    sqlText = sqlText & "CashDiscountDays LONG, "
    sqlText = sqlText & "CashDiscountPercent DOUBLE, "
    sqlText = sqlText & "IsActive YESNO, "
    sqlText = sqlText & "SortOrder LONG"
    sqlText = sqlText & ");"

    ExecuteDdl db, sqlText
    ExecuteDdl db, "CREATE UNIQUE INDEX ux_refPaymentTerms_Code ON refPaymentTerms (PaymentTermCode);"
End Sub

Private Sub CreateRefVatCodes(ByVal db As DAO.Database)
    Dim sqlText As String

    sqlText = ""
    sqlText = sqlText & "CREATE TABLE refVatCodes ("
    sqlText = sqlText & "VatCode TEXT(20) CONSTRAINT pk_refVatCodes PRIMARY KEY, "
    sqlText = sqlText & "VatName TEXT(100), "
    sqlText = sqlText & "VatRate DOUBLE, "
    sqlText = sqlText & "CountryCode TEXT(2), "
    sqlText = sqlText & "IsActive YESNO, "
    sqlText = sqlText & "ValidFrom DATETIME, "
    sqlText = sqlText & "ValidTo DATETIME"
    sqlText = sqlText & ");"

    ExecuteDdl db, sqlText
End Sub

Private Sub CreateRefUnits(ByVal db As DAO.Database)
    Dim sqlText As String

    sqlText = ""
    sqlText = sqlText & "CREATE TABLE refUnits ("
    sqlText = sqlText & "UnitCode TEXT(20) CONSTRAINT pk_refUnits PRIMARY KEY, "
    sqlText = sqlText & "UnitName TEXT(100), "
    sqlText = sqlText & "SortOrder LONG, "
    sqlText = sqlText & "IsActive YESNO"
    sqlText = sqlText & ");"

    ExecuteDdl db, sqlText
End Sub

Private Sub CreateTblAddresses(ByVal db As DAO.Database)
    Dim sqlText As String

    sqlText = ""
    sqlText = sqlText & "CREATE TABLE tblAddresses ("
    sqlText = sqlText & "AddressId AUTOINCREMENT CONSTRAINT pk_tblAddresses PRIMARY KEY, "
    sqlText = sqlText & "AddressType TEXT(30), "
    sqlText = sqlText & "CompanyName TEXT(150), "
    sqlText = sqlText & "FirstName TEXT(80), "
    sqlText = sqlText & "LastName TEXT(80), "
    sqlText = sqlText & "Street TEXT(120), "
    sqlText = sqlText & "HouseNo TEXT(20), "
    sqlText = sqlText & "PostalCode TEXT(20), "
    sqlText = sqlText & "City TEXT(100), "
    sqlText = sqlText & "CountryCode TEXT(2), "
    sqlText = sqlText & "Email TEXT(150), "
    sqlText = sqlText & "Phone TEXT(50), "
    sqlText = sqlText & "VatNo TEXT(50), "
    sqlText = sqlText & "LanguageCode TEXT(10), "
    sqlText = sqlText & "CurrencyCode TEXT(3), "
    sqlText = sqlText & "PaymentTermId LONG, "
    sqlText = sqlText & "IsActive YESNO, "
    sqlText = sqlText & "CreatedAt DATETIME, "
    sqlText = sqlText & "CreatedBy TEXT(50), "
    sqlText = sqlText & "UpdatedAt DATETIME, "
    sqlText = sqlText & "UpdatedBy TEXT(50)"
    sqlText = sqlText & ");"

    ExecuteDdl db, sqlText
    ExecuteDdl db, "CREATE INDEX ix_tblAddresses_AddressType ON tblAddresses (AddressType);"
    ExecuteDdl db, "CREATE INDEX ix_tblAddresses_CompanyName ON tblAddresses (CompanyName);"
    ExecuteDdl db, "CREATE INDEX ix_tblAddresses_PaymentTermId ON tblAddresses (PaymentTermId);"
End Sub

Private Sub CreateTblProductGroups(ByVal db As DAO.Database)
    Dim sqlText As String

    sqlText = ""
    sqlText = sqlText & "CREATE TABLE tblProductGroups ("
    sqlText = sqlText & "ProductGroupId AUTOINCREMENT CONSTRAINT pk_tblProductGroups PRIMARY KEY, "
    sqlText = sqlText & "ProductGroupCode TEXT(30) NOT NULL, "
    sqlText = sqlText & "ProductGroupName TEXT(100), "
    sqlText = sqlText & "RevenueAccount TEXT(20), "
    sqlText = sqlText & "ExpenseAccount TEXT(20), "
    sqlText = sqlText & "VatCode TEXT(20), "
    sqlText = sqlText & "IsActive YESNO, "
    sqlText = sqlText & "SortOrder LONG, "
    sqlText = sqlText & "CreatedAt DATETIME, "
    sqlText = sqlText & "CreatedBy TEXT(50), "
    sqlText = sqlText & "UpdatedAt DATETIME, "
    sqlText = sqlText & "UpdatedBy TEXT(50)"
    sqlText = sqlText & ");"

    ExecuteDdl db, sqlText
    ExecuteDdl db, "CREATE UNIQUE INDEX ux_tblProductGroups_Code ON tblProductGroups (ProductGroupCode);"
    ExecuteDdl db, "CREATE INDEX ix_tblProductGroups_VatCode ON tblProductGroups (VatCode);"
End Sub

Private Sub CreateTblArticles(ByVal db As DAO.Database)
    Dim sqlText As String

    sqlText = ""
    sqlText = sqlText & "CREATE TABLE tblArticles ("
    sqlText = sqlText & "ArticleId AUTOINCREMENT CONSTRAINT pk_tblArticles PRIMARY KEY, "
    sqlText = sqlText & "ArticleNo TEXT(50) NOT NULL, "
    sqlText = sqlText & "ArticleName TEXT(150), "
    sqlText = sqlText & "Description LONGTEXT, "
    sqlText = sqlText & "ProductGroupId LONG, "
    sqlText = sqlText & "UnitCode TEXT(20), "
    sqlText = sqlText & "SalesPrice CURRENCY, "
    sqlText = sqlText & "PurchasePrice CURRENCY, "
    sqlText = sqlText & "CurrencyCode TEXT(3), "
    sqlText = sqlText & "VatCode TEXT(20), "
    sqlText = sqlText & "IsStockArticle YESNO, "
    sqlText = sqlText & "IsServiceArticle YESNO, "
    sqlText = sqlText & "IsActive YESNO, "
    sqlText = sqlText & "CreatedAt DATETIME, "
    sqlText = sqlText & "CreatedBy TEXT(50), "
    sqlText = sqlText & "UpdatedAt DATETIME, "
    sqlText = sqlText & "UpdatedBy TEXT(50)"
    sqlText = sqlText & ");"

    ExecuteDdl db, sqlText
    ExecuteDdl db, "CREATE UNIQUE INDEX ux_tblArticles_ArticleNo ON tblArticles (ArticleNo);"
    ExecuteDdl db, "CREATE INDEX ix_tblArticles_ProductGroupId ON tblArticles (ProductGroupId);"
    ExecuteDdl db, "CREATE INDEX ix_tblArticles_UnitCode ON tblArticles (UnitCode);"
    ExecuteDdl db, "CREATE INDEX ix_tblArticles_VatCode ON tblArticles (VatCode);"
End Sub

Private Sub CreateTblOrders(ByVal db As DAO.Database)
    Dim sqlText As String

    sqlText = ""
    sqlText = sqlText & "CREATE TABLE tblOrders ("
    sqlText = sqlText & "OrderId AUTOINCREMENT CONSTRAINT pk_tblOrders PRIMARY KEY, "
    sqlText = sqlText & "OrderNo TEXT(50), "
    sqlText = sqlText & "OrderType TEXT(30), "
    sqlText = sqlText & "OrderStatus TEXT(30), "
    sqlText = sqlText & "CustomerAddressId LONG, "
    sqlText = sqlText & "OrderDate DATETIME, "
    sqlText = sqlText & "DeliveryDate DATETIME, "
    sqlText = sqlText & "ValidUntil DATETIME, "
    sqlText = sqlText & "ReferenceText TEXT(150), "
    sqlText = sqlText & "LanguageCode TEXT(10), "
    sqlText = sqlText & "CurrencyCode TEXT(3), "
    sqlText = sqlText & "PaymentTermId LONG, "
    sqlText = sqlText & "SubtotalNet CURRENCY, "
    sqlText = sqlText & "TotalDiscount CURRENCY, "
    sqlText = sqlText & "TotalSurcharge CURRENCY, "
    sqlText = sqlText & "TotalVat CURRENCY, "
    sqlText = sqlText & "TotalGross CURRENCY, "
    sqlText = sqlText & "Notes LONGTEXT, "
    sqlText = sqlText & "InternalNotes LONGTEXT, "
    sqlText = sqlText & "CreatedAt DATETIME, "
    sqlText = sqlText & "CreatedBy TEXT(50), "
    sqlText = sqlText & "UpdatedAt DATETIME, "
    sqlText = sqlText & "UpdatedBy TEXT(50)"
    sqlText = sqlText & ");"

    ExecuteDdl db, sqlText
    ExecuteDdl db, "CREATE UNIQUE INDEX ux_tblOrders_OrderNo ON tblOrders (OrderNo);"
    ExecuteDdl db, "CREATE INDEX ix_tblOrders_CustomerAddressId ON tblOrders (CustomerAddressId);"
    ExecuteDdl db, "CREATE INDEX ix_tblOrders_OrderDate ON tblOrders (OrderDate);"
    ExecuteDdl db, "CREATE INDEX ix_tblOrders_OrderStatus ON tblOrders (OrderStatus);"
    ExecuteDdl db, "CREATE INDEX ix_tblOrders_PaymentTermId ON tblOrders (PaymentTermId);"
End Sub

Private Sub CreateTblOrderLines(ByVal db As DAO.Database)
    Dim sqlText As String

    sqlText = ""
    sqlText = sqlText & "CREATE TABLE tblOrderLines ("
    sqlText = sqlText & "OrderLineId AUTOINCREMENT CONSTRAINT pk_tblOrderLines PRIMARY KEY, "
    sqlText = sqlText & "OrderId LONG NOT NULL, "
    sqlText = sqlText & "LineNo LONG, "
    sqlText = sqlText & "ArticleId LONG, "
    sqlText = sqlText & "LineType TEXT(30), "
    sqlText = sqlText & "Description LONGTEXT, "
    sqlText = sqlText & "Quantity DOUBLE, "
    sqlText = sqlText & "UnitCode TEXT(20), "
    sqlText = sqlText & "UnitPrice CURRENCY, "
    sqlText = sqlText & "DiscountPercent DOUBLE, "
    sqlText = sqlText & "DiscountAmount CURRENCY, "
    sqlText = sqlText & "SurchargePercent DOUBLE, "
    sqlText = sqlText & "SurchargeAmount CURRENCY, "
    sqlText = sqlText & "VatCode TEXT(20), "
    sqlText = sqlText & "VatRate DOUBLE, "
    sqlText = sqlText & "LineNetAmount CURRENCY, "
    sqlText = sqlText & "LineVatAmount CURRENCY, "
    sqlText = sqlText & "LineGrossAmount CURRENCY, "
    sqlText = sqlText & "SortOrder LONG, "
    sqlText = sqlText & "CreatedAt DATETIME, "
    sqlText = sqlText & "CreatedBy TEXT(50), "
    sqlText = sqlText & "UpdatedAt DATETIME, "
    sqlText = sqlText & "UpdatedBy TEXT(50)"
    sqlText = sqlText & ");"

    ExecuteDdl db, sqlText
    ExecuteDdl db, "CREATE INDEX ix_tblOrderLines_OrderId ON tblOrderLines (OrderId);"
    ExecuteDdl db, "CREATE INDEX ix_tblOrderLines_ArticleId ON tblOrderLines (ArticleId);"
    ExecuteDdl db, "CREATE INDEX ix_tblOrderLines_LineNo ON tblOrderLines (LineNo);"
    ExecuteDdl db, "CREATE INDEX ix_tblOrderLines_VatCode ON tblOrderLines (VatCode);"
End Sub

Private Sub ExecuteDdl(ByVal db As DAO.Database, ByVal sqlText As String)
    On Error GoTo ErrorHandler

    db.Execute sqlText, dbFailOnError
    Debug.Print "OK: " & sqlText
    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 3010, 3283, 3371
            Debug.Print "SKIP: " & Err.Number & " - " & Err.description
            Err.Clear

        Case Else
            Debug.Print "ERROR: " & Err.Number & " - " & Err.description
            Debug.Print sqlText
            Err.Raise Err.Number, MODULE_NAME & ".ExecuteDdl", Err.description
    End Select
End Sub
