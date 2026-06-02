Option Compare Database
Option Explicit

'===============================================================================
' Module    : modDemoDataSeeder
' Purpose   : Safely seeds deterministic Easis v4 demo data against the current schema.
' Author    : ChatGPT
' Version   : 1.0.0
' Notes     : - Uses DAO only.
'             - Does not modify ref_* tables.
'             - Default mode is non-destructive and only inserts missing demo rows.
'             - Reset mode removes only recognized demo rows before rebuilding them.
'             - Uses AutoNumber primary keys where present and keeps generated IDs in memory.
'===============================================================================

Private Const MODULE_NAME As String = "modDemoDataSeeder"

Private Const TBL_ADR_ADDRESS As String = "adr_address"
Private Const TBL_ADR_CONTACT As String = "adr_contact"
Private Const TBL_DOC_DOCUMENT As String = "doc_document"
Private Const TBL_DOC_POSITION As String = "doc_document_position"
Private Const TBL_TEN_PARAMETER As String = "ten_parameter"
Private Const TBL_ART_PRODUCT_GROUP As String = "art_product_group"
Private Const TBL_ART_ARTICLE As String = "art_article"

Private Const created_by As String = "DemoDataSeeder"

Private mInsertedCount As Long
Private mSkippedExistingCount As Long
Private mUpdatedRequiredFieldsCount As Long
Private mDeletedCount As Long

' Cached IDs generated during one seed run.
Private mAddressBillingCh As Long
Private mAddressShippingCh As Long
Private mAddressBillingDe As Long
Private mAddressShippingFr As Long
Private mAddressBillingUs As Long

Private mDocInvoiceCh As Long
Private mDocInvoiceEu As Long
Private mDocDeliveryNote As Long
Private mDocCreditNote As Long
Private mDocInvoiceUsd As Long
Private mDocLongInvoice As Long

Private mRefreshDocInvoiceChTotals As Boolean
Private mRefreshDocInvoiceEuTotals As Boolean
Private mRefreshDocDeliveryNoteTotals As Boolean
Private mRefreshDocCreditNoteTotals As Boolean
Private mRefreshDocInvoiceUsdTotals As Boolean
Private mRefreshDocLongInvoiceTotals As Boolean

Public Sub SeedDemoData(Optional ByVal TenantCode As String = "DEMO_CH")
    SeedDemoDataSafe TenantCode
End Sub

Public Sub SeedDemoDataSafe(Optional ByVal TenantCode As String = "DEMO_CH")
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tenantCodeEffective As String

    tenantCodeEffective = UCase$(Trim$(TenantCode))
    If LenB(tenantCodeEffective) = 0 Then tenantCodeEffective = "DEMO_CH"

    Set db = CurrentDb

    ResetSeedRunCounters
    DBEngine.Workspaces(0).BeginTrans

    ValidateRequiredSchema db
    EnsureTenantParameters db, tenantCodeEffective
    EnsureAddresses db
    EnsureContacts db
    EnsureDocuments db
    EnsurePositions db
    UpdateDocumentTotals db

    DBEngine.Workspaces(0).CommitTrans

    LogSeedSummary "SeedDemoDataSafe", tenantCodeEffective

    MsgBox "Demo-Daten wurden erfolgreich ergaenzt." & vbCrLf & _
           "TenantCode: " & tenantCodeEffective, vbInformation, "Easis v4 Demo Seeder"

CleanExit:
    Set db = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    DBEngine.Workspaces(0).Rollback
    modErrorHandler.HandleError MODULE_NAME, "SeedDemoDataSafe", Err
    MsgBox "Fehler beim Erstellen der Demo-Daten:" & vbCrLf & _
           Err.Number & " - " & Err.description, vbCritical, "Easis v4 Demo Seeder"
    Resume CleanExit
End Sub

Public Sub ResetAndSeedDemoData(Optional ByVal TenantCode As String = "DEMO_CH")
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tenantCodeEffective As String

    If MsgBox("Ausschliesslich erkannte Demo-Daten wirklich loeschen und neu erstellen?", _
              vbQuestion + vbYesNo + vbDefaultButton2, "Easis v4 Demo Seeder") <> vbYes Then
        Exit Sub
    End If

    tenantCodeEffective = UCase$(Trim$(TenantCode))
    If LenB(tenantCodeEffective) = 0 Then tenantCodeEffective = "DEMO_CH"

    Set db = CurrentDb

    ResetSeedRunCounters
    DBEngine.Workspaces(0).BeginTrans

    ValidateRequiredSchema db
    DeleteDemoDataForReset db
    EnsureTenantParameters db, tenantCodeEffective
    EnsureAddresses db
    EnsureContacts db
    EnsureDocuments db
    EnsurePositions db
    UpdateDocumentTotals db

    DBEngine.Workspaces(0).CommitTrans

    LogSeedSummary "ResetAndSeedDemoData", tenantCodeEffective

    MsgBox "Demo-Daten wurden geloescht und neu erstellt." & vbCrLf & _
           "TenantCode: " & tenantCodeEffective, vbInformation, "Easis v4 Demo Seeder"

CleanExit:
    Set db = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    DBEngine.Workspaces(0).Rollback
    modErrorHandler.HandleError MODULE_NAME, "ResetAndSeedDemoData", Err
    MsgBox "Fehler beim Zuruecksetzen der Demo-Daten:" & vbCrLf & _
           Err.Number & " - " & Err.description, vbCritical, "Easis v4 Demo Seeder"
    Resume CleanExit
End Sub

Public Sub SeedProductGroups()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database

    Set db = CurrentDb

    modArticleGroupService.EnsureArticleGroupTable
    RequireTable db, TBL_ART_PRODUCT_GROUP
    RequireField db, TBL_ART_PRODUCT_GROUP, "product_group_code"
    RequireField db, TBL_ART_PRODUCT_GROUP, "product_group_name"
    RequireField db, TBL_ART_PRODUCT_GROUP, "description_text"
    RequireField db, TBL_ART_PRODUCT_GROUP, "sort_order"
    RequireField db, TBL_ART_PRODUCT_GROUP, "is_active"
    RequireField db, TBL_ART_PRODUCT_GROUP, "created_at"
    RequireField db, TBL_ART_PRODUCT_GROUP, "created_by"
    RequireField db, TBL_ART_PRODUCT_GROUP, "updated_at"
    RequireField db, TBL_ART_PRODUCT_GROUP, "updated_by"

    DBEngine.Workspaces(0).BeginTrans

    ' Demo business values are intentionally language-neutral. UI localization
    ' belongs in fw_translation, not in the business rows themselves.
    UpsertProductGroup db, "SERVICES", "Dienstleistungen", "Dienstleistungs- und Beratungsangebote", 10, True
    UpsertProductGroup db, "FOOD", "Lebensmittel", "Allgemeine Lebensmittel und Verpflegungsartikel", 20, True
    UpsertProductGroup db, "BEVERAGES", "Getraenke", "Getraenke und zugehoerige Artikel", 30, True
    UpsertProductGroup db, "OFFICE", "Buero", "Buero- und Verwaltungsbedarf", 40, True
    UpsertProductGroup db, "SOFTWARE", "Software", "Softwareprodukte, Lizenzen und digitale Services", 50, True
    UpsertProductGroup db, "HARDWARE", "Hardware", "Hardwarekomponenten und Geraete", 60, True
    UpsertProductGroup db, "SUBSCRIPTIONS", "Abonnemente", "Wiederkehrende Leistungen und Abomodelle", 70, True

    DBEngine.Workspaces(0).CommitTrans

    modLoggingHandler.LogInfo MODULE_NAME & ".SeedProductGroups", _
        "Product group demo data seeded successfully."
    MsgBox "Produktgruppen-Demo-Daten wurden erfolgreich initialisiert.", vbInformation, "Easis v4 Demo Seeder"

CleanExit:
    Set db = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    DBEngine.Workspaces(0).Rollback
    modErrorHandler.HandleError MODULE_NAME, "SeedProductGroups", Err
    MsgBox "Fehler beim Initialisieren der Produktgruppen-Demo-Daten:" & vbCrLf & _
           Err.Number & " - " & Err.description, vbCritical, "Easis v4 Demo Seeder"
    Resume CleanExit
End Sub

Public Sub SeedArticles()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database

    Set db = CurrentDb

    If Not TableExists(db, TBL_ART_PRODUCT_GROUP) Then
        SeedProductGroups
    ElseIf DCount("*", TBL_ART_PRODUCT_GROUP) = 0 Then
        SeedProductGroups
    End If

    RequireTable db, TBL_ART_PRODUCT_GROUP
    RequireTable db, TBL_ART_ARTICLE

    RequireField db, TBL_ART_ARTICLE, "article_no"
    RequireField db, TBL_ART_ARTICLE, "article_name"
    RequireField db, TBL_ART_ARTICLE, "product_group_id"
    RequireField db, TBL_ART_ARTICLE, "article_type_code"
    RequireField db, TBL_ART_ARTICLE, "unit_code"
    RequireField db, TBL_ART_ARTICLE, "vat_code"
    RequireField db, TBL_ART_ARTICLE, "purchase_price"
    RequireField db, TBL_ART_ARTICLE, "sales_price"
    RequireField db, TBL_ART_ARTICLE, "barcode"
    RequireField db, TBL_ART_ARTICLE, "description_text"
    RequireField db, TBL_ART_ARTICLE, "is_active"
    RequireField db, TBL_ART_ARTICLE, "created_at"
    RequireField db, TBL_ART_ARTICLE, "created_by"
    RequireField db, TBL_ART_ARTICLE, "updated_at"
    RequireField db, TBL_ART_ARTICLE, "updated_by"

    DBEngine.Workspaces(0).BeginTrans

    ' Demo article business values are intentionally language-neutral.
    UpsertArticle db, "CONSULT-STD", "Standard Beratung", "SERVICES", "SERVICE", "H", "CH_STANDARD", 0, 180, "7611000000011", "Standardisierte Beratungsleistung.", True
    UpsertArticle db, "SERVICE-HOUR", "Servicestunde", "SERVICES", "SERVICE", "H", "CH_STANDARD", 0, 145, "7611000000012", "Abrechenbare Servicestunde.", True
    UpsertArticle db, "SOFTWARE-BASIC", "Software Basislizenz", "SOFTWARE", "LICENSE", "PCS", "CH_STANDARD", 120, 490, "7611000000013", "Einfache Basislizenz fuer Standardsoftware.", True
    UpsertArticle db, "HARDWARE-BOX", "Hardware Box", "HARDWARE", "GOODS", "PCS", "CH_STANDARD", 95, 245, "7611000000014", "Standardisierte Hardwareeinheit.", True
    UpsertArticle db, "OFFICE-MAT", "Bueromaterial", "OFFICE", "GOODS", "PCS", "CH_STANDARD", 8, 19, "7611000000015", "Allgemeines Bueromaterial.", True
    UpsertArticle db, "FOOD-SNACK", "Snack", "FOOD", "GOODS", "PCS", "CH_REDUCED", 1.2, 3.5, "7611000000016", "Einfacher Snackartikel fuer Demo-Zwecke.", True
    UpsertArticle db, "BEV-WATER", "Mineralwasser", "BEVERAGES", "GOODS", "PCS", "CH_REDUCED", 0.4, 1.8, "7611000000017", "Flasche Mineralwasser.", True

    DBEngine.Workspaces(0).CommitTrans

    modLoggingHandler.LogInfo MODULE_NAME & ".SeedArticles", _
        "Article demo data seeded successfully."
    MsgBox "Artikel-Demo-Daten wurden erfolgreich initialisiert.", vbInformation, "Easis v4 Demo Seeder"

CleanExit:
    Set db = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    DBEngine.Workspaces(0).Rollback
    modErrorHandler.HandleError MODULE_NAME, "SeedArticles", Err
    MsgBox "Fehler beim Initialisieren der Artikel-Demo-Daten:" & vbCrLf & _
           Err.Number & " - " & Err.description, vbCritical, "Easis v4 Demo Seeder"
    Resume CleanExit
End Sub

Private Sub ValidateRequiredSchema(ByVal db As DAO.Database)
    RequireTable db, TBL_ADR_ADDRESS
    RequireTable db, TBL_DOC_DOCUMENT
    RequireTable db, TBL_DOC_POSITION

    RequireField db, TBL_ADR_ADDRESS, "address_id"
    RequireField db, TBL_ADR_ADDRESS, "address_type_code"
    RequireField db, TBL_ADR_ADDRESS, "company_name"
    RequireField db, TBL_ADR_ADDRESS, "street"
    RequireField db, TBL_ADR_ADDRESS, "zip_code"
    RequireField db, TBL_ADR_ADDRESS, "city"
    RequireField db, TBL_ADR_ADDRESS, "country_code"
    RequireField db, TBL_ADR_ADDRESS, "language_code"

    RequireField db, TBL_DOC_DOCUMENT, "document_id"
    RequireField db, TBL_DOC_DOCUMENT, "document_type_code"
    RequireField db, TBL_DOC_DOCUMENT, "document_status_code"
    RequireField db, TBL_DOC_DOCUMENT, "document_no"
    RequireField db, TBL_DOC_DOCUMENT, "document_date"
    RequireField db, TBL_DOC_DOCUMENT, "customer_address_id"
    RequireField db, TBL_DOC_DOCUMENT, "customer_name"
    RequireField db, TBL_DOC_DOCUMENT, "currency_code"
    RequireField db, TBL_DOC_DOCUMENT, "vat_mode"
    RequireField db, TBL_DOC_DOCUMENT, "vat_rate"
    RequireField db, TBL_DOC_DOCUMENT, "total_net"
    RequireField db, TBL_DOC_DOCUMENT, "total_vat"
    RequireField db, TBL_DOC_DOCUMENT, "total_gross"

    RequireField db, TBL_DOC_POSITION, "document_position_id"
    RequireField db, TBL_DOC_POSITION, "document_id"
    RequireField db, TBL_DOC_POSITION, "line_no"
    RequireField db, TBL_DOC_POSITION, "description"
    RequireField db, TBL_DOC_POSITION, "quantity"
    RequireField db, TBL_DOC_POSITION, "unit_code"
    RequireField db, TBL_DOC_POSITION, "unit_price"
    RequireField db, TBL_DOC_POSITION, "vat_rate"
    RequireField db, TBL_DOC_POSITION, "line_total_net"
    RequireField db, TBL_DOC_POSITION, "line_total_vat"
    RequireField db, TBL_DOC_POSITION, "line_total_gross"
End Sub

Private Sub EnsureTenantParameters(ByVal db As DAO.Database, ByVal TenantCode As String)
    If Not TableExists(db, TBL_TEN_PARAMETER) Then Exit Sub

    EnsureTenantParameter db, TenantCode, "TENANT_CODE", TenantCode
    EnsureTenantParameter db, TenantCode, "TENANT_NAME", "Easis Demo Schweiz AG"
    EnsureTenantParameter db, TenantCode, "DEFAULT_LANGUAGE", "de-CH"
    EnsureTenantParameter db, TenantCode, "DEFAULT_CURRENCY", "CHF"
    EnsureTenantParameter db, TenantCode, "SENDER_NAME", "Easis Demo Schweiz AG"
    EnsureTenantParameter db, TenantCode, "SENDER_STREET", "Bahnhofstrasse"
    EnsureTenantParameter db, TenantCode, "SENDER_HOUSE_NO", "10"
    EnsureTenantParameter db, TenantCode, "SENDER_ZIP_CODE", "8001"
    EnsureTenantParameter db, TenantCode, "SENDER_CITY", DemoCityZurich()
    EnsureTenantParameter db, TenantCode, "SENDER_COUNTRY_CODE", "CH"
    EnsureTenantParameter db, TenantCode, "SENDER_PHONE", "+41 44 123 45 67"
    EnsureTenantParameter db, TenantCode, "SENDER_EMAIL", "demo@easis.ch"
    EnsureTenantParameter db, TenantCode, "SENDER_VAT_NO", "CHE-123.456.789 MWST"
End Sub

Private Sub EnsureAddresses(ByVal db As DAO.Database)
    mAddressBillingCh = EnsureAddress(db, "BILLING", "Muster Handel AG", "Anna", "Keller", "Industriestrasse", "15", "6300", "Zug", "CH", "de-CH")
    mAddressShippingCh = EnsureAddress(db, "SHIPPING", "Muster Handel AG - Lager Genf", "Marc", "Dubois", "Route de Meyrin", "88", "1203", DemoCityGeneva(), "CH", "fr-CH")
    mAddressBillingDe = EnsureAddress(db, "BILLING", "Beispiel GmbH", "Thomas", "Schneider", "Hauptstrasse", "22", "80331", DemoCityMunich(), "DE", "de-DE")
    mAddressShippingFr = EnsureAddress(db, "SHIPPING", "Beispiel GmbH - Site Paris", "Claire", "Martin", "Rue Lafayette", "12", "75009", "Paris", "FR", "fr-FR")
    mAddressBillingUs = EnsureAddress(db, "BILLING", "Global Components Inc.", "John", "Miller", "Market Street", "500", "94105", "San Francisco", "US", "en-US")
End Sub

Private Sub EnsureContacts(ByVal db As DAO.Database)
    If Not TableExists(db, TBL_ADR_CONTACT) Then Exit Sub

    EnsureContact db, mAddressBillingCh, "EMAIL", "buchhaltung@muster-handel.ch", True, "Demo billing contact"
    EnsureContact db, mAddressShippingCh, "EMAIL", "lager@muster-handel.ch", True, "Demo shipping contact"
    EnsureContact db, mAddressBillingDe, "EMAIL", "rechnung@beispiel-gmbh.de", True, "Demo billing contact"
    EnsureContact db, mAddressShippingFr, "EMAIL", "livraison@beispiel.fr", True, "Demo shipping contact"
    EnsureContact db, mAddressBillingUs, "EMAIL", "ap@global-components.com", True, "Demo billing contact"
End Sub

Private Sub EnsureDocuments(ByVal db As DAO.Database)

    mDocInvoiceCh = EnsureDocument(db, "INVOICE", "FINAL", "RE-2026-0001", DateSerial(2026, 5, 2), mAddressBillingCh, "Muster Handel AG", "CHF", "NET", 7.7, "Schweizer Rechnung / Standardfall", "", 0, "de-CH", "NET_30", "Zahlbar innert 30 Tagen netto.")

    mDocInvoiceEu = EnsureDocument(db, "INVOICE", "FINAL", "RE-2026-0002", DateSerial(2026, 5, 2), mAddressBillingDe, "Beispiel GmbH", "EUR", "NET", 19, "EU-Rechnung mit Positionsrabatt und Kopfrabatt. Separate Lieferadresse: Beispiel GmbH - Site Paris, 12 Rue Lafayette, 75009 Paris, FR", "PERCENT", 5, "de-DE", "CASH_DISCOUNT_10_2_NET_30", "2% Skonto bei Zahlung innert 10 Tagen, ansonsten zahlbar innert 30 Tagen netto.")

    mDocDeliveryNote = EnsureDocument(db, "DELIVERY_NOTE", "FINAL", "LS-2026-0001", DateSerial(2026, 5, 2), mAddressShippingCh, "Muster Handel AG - Lager Genf", "CHF", "NET", 7.7, DemoDeliveryNoteRemarks(), "", 0, "de-CH", "", "")

    mDocCreditNote = EnsureDocument(db, "CREDIT_NOTE", "FINAL", "GS-2026-0001", DateSerial(2026, 5, 2), mAddressBillingDe, "Beispiel GmbH", "EUR", "NET", 19, "Gutschrift mit negativer Position", "", 0, "de-DE", "NET_30", "Zahlbar innert 30 Tagen netto.")

    mDocInvoiceUsd = EnsureDocument(db, "INVOICE", "FINAL", "RE-2026-0003", DateSerial(2026, 5, 2), mAddressBillingUs, "Global Components Inc.", "USD", "EXPORT", 0, DemoExportInvoiceRemarks(), "", 0, "en-US", "PREPAYMENT", "Payable in advance.")

    mDocLongInvoice = EnsureDocument(db, "INVOICE", "FINAL", "RE-2026-0099", DateSerial(2026, 5, 3), mAddressBillingCh, "Muster Handel AG", "CHF", "NET", 7.7, DemoLongInvoiceRemarks(), "", 0, "de-CH", "NET_30", "Zahlbar innert 30 Tagen netto.")

End Sub

Private Sub EnsurePositions(ByVal db As DAO.Database)
    Dim i As Long
    Dim vatRate As Double
    Dim UnitPrice As Currency
    Dim quantity As Double
    Dim description As String

    EnsurePosition db, mDocInvoiceCh, 1, "Beratung Architekturreview Easis v4", 4, "h", 180, 7.7
    EnsurePosition db, mDocInvoiceCh, 2, "Einrichtung Reporting-Template", 1, "pauschal", 650, 7.7

    EnsurePosition db, mDocInvoiceEu, 1, "Access Frontend Erweiterung", 8, "h", 145, 19
    EnsurePosition db, mDocInvoiceEu, 2, "Technische Dokumentation", 1, "pauschal", 390, 19, "PERCENT", 10

    EnsurePosition db, mDocDeliveryNote, 1, "Demo Hardware Box", 3, "Stk", 245, 7.7
    EnsurePosition db, mDocDeliveryNote, 2, "USB-C Anschlusskabel", 6, "Stk", 18.5, 7.7

    EnsurePosition db, mDocCreditNote, 1, "Gutschrift Servicekorrektur", 1, "pauschal", -250, 19

    EnsurePosition db, mDocInvoiceUsd, 1, "Software license export", 5, "pcs", 499, 0
    EnsurePosition db, mDocInvoiceUsd, 2, "International remote support package", 1, "package", 1250, 0, "PERCENT", 15
    EnsurePosition db, mDocInvoiceUsd, 3, "Rounding and large amount test position", 12.5, "h", 1234.56, 0, "PERCENT", 2.5

    For i = 1 To 50
        If i Mod 10 = 0 Then
            vatRate = 2.5
        ElseIf i Mod 15 = 0 Then
            vatRate = 0
        Else
            vatRate = 7.7
        End If

        UnitPrice = CCur(45 + (i * 4.25))
        quantity = 1 + (i Mod 5)

        description = DemoLongPositionDescription(i)

        EnsurePosition _
            db, _
            mDocLongInvoice, _
            i, _
            description, _
            quantity, _
            "Stk", _
            UnitPrice, _
            vatRate
    Next i
End Sub
Private Function InsertAddress( _
    ByVal db As DAO.Database, _
    ByVal addressTypeCode As String, _
    ByVal CompanyName As String, _
    ByVal FirstName As String, _
    ByVal LastName As String, _
    ByVal Street As String, _
    ByVal HouseNo As String, _
    ByVal zipCode As String, _
    ByVal City As String, _
    ByVal countryCode As String, _
    ByVal LanguageCode As String _
) As Long
    Dim rs As DAO.Recordset

    Set rs = db.OpenRecordset(TBL_ADR_ADDRESS, dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    SetFieldIfExists rs, "address_type_code", UCase$(Trim$(addressTypeCode))
    SetFieldIfExists rs, "company_name", Trim$(CompanyName)
    SetFieldIfExists rs, "first_name", Trim$(FirstName)
    SetFieldIfExists rs, "last_name", Trim$(LastName)
    SetFieldIfExists rs, "street", Trim$(Street)
    SetFieldIfExists rs, "house_no", Trim$(HouseNo)
    SetFieldIfExists rs, "zip_code", Trim$(zipCode)
    SetFieldIfExists rs, "city", Trim$(City)
    SetFieldIfExists rs, "country_code", UCase$(Trim$(countryCode))
    SetFieldIfExists rs, "language_code", Trim$(LanguageCode)
    SetFieldIfExists rs, "is_active", True
    SetFieldIfExists rs, "created_at", Now()
    SetFieldIfExists rs, "created_by", created_by
    rs.Update

    rs.Bookmark = rs.LastModified
    InsertAddress = CLng(Nz(rs.Fields("address_id").Value, 0))

    rs.Close
    Set rs = Nothing
End Function

Private Sub InsertContact( _
    ByVal db As DAO.Database, _
    ByVal addressId As Long, _
    ByVal contactTypeCode As String, _
    ByVal ContactValue As String, _
    ByVal IsPrimary As Boolean, _
    ByVal remarks As String _
)
    Dim rs As DAO.Recordset

    Set rs = db.OpenRecordset(TBL_ADR_CONTACT, dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    SetFieldIfExists rs, "address_id", addressId
    SetFieldIfExists rs, "contact_type_code", UCase$(Trim$(contactTypeCode))
    SetFieldIfExists rs, "contact_value", Trim$(ContactValue)
    SetFieldIfExists rs, "is_primary", IsPrimary
    SetFieldIfExists rs, "remarks", Trim$(remarks)
    SetFieldIfExists rs, "created_at", Now()
    SetFieldIfExists rs, "created_by", created_by
    rs.Update

    rs.Close
    Set rs = Nothing
End Sub

Private Function InsertDocument( _
    ByVal db As DAO.Database, _
    ByVal DocumentTypeCode As String, _
    ByVal DocumentStatusCode As String, _
    ByVal DocumentNo As String, _
    ByVal DocumentDate As Date, _
    ByVal CustomerAddressId As Long, _
    ByVal CustomerName As String, _
    ByVal CurrencyCode As String, _
    ByVal VatMode As String, _
    ByVal vatRate As Double, _
    ByVal remarks As String, _
    Optional ByVal HeaderDiscountType As String = "", _
    Optional ByVal HeaderDiscountValue As Double = 0, _
    Optional ByVal LanguageCode As String = "", _
    Optional ByVal PaymentTermCode As String = "", _
    Optional ByVal PaymentTermsText As String = "" _
) As Long
    
    Dim rs As DAO.Recordset

    Set rs = db.OpenRecordset(TBL_DOC_DOCUMENT, dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    SetFieldIfExists rs, "document_type_code", UCase$(Trim$(DocumentTypeCode))
    SetFieldIfExists rs, "document_status_code", UCase$(Trim$(DocumentStatusCode))
    SetFieldIfExists rs, "document_no", Trim$(DocumentNo)
    SetFieldIfExists rs, "document_date", DocumentDate
    SetFieldIfExists rs, "customer_address_id", CustomerAddressId
    SetFieldIfExists rs, "customer_name", Trim$(CustomerName)
    SetFieldIfExists rs, "currency_code", UCase$(Trim$(CurrencyCode))
    SetFieldIfExists rs, "header_discount_type", NullIfEmpty(HeaderDiscountType)
    SetFieldIfExists rs, "header_discount_value", HeaderDiscountValue
    SetFieldIfExists rs, "header_discount_amount", CCur(0)
    SetFieldIfExists rs, "vat_mode", UCase$(Trim$(VatMode))
    SetFieldIfExists rs, "language_code", NullIfEmpty(LanguageCode)
    SetFieldIfExists rs, "payment_term_code", NullIfEmpty(PaymentTermCode)
    SetFieldIfExists rs, "payment_terms_text", NullIfEmpty(PaymentTermsText)
    SetFieldIfExists rs, "vat_rate", vatRate
    SetFieldIfExists rs, "total_net", CCur(0)
    SetFieldIfExists rs, "total_vat", CCur(0)
    SetFieldIfExists rs, "total_gross", CCur(0)
    SetFieldIfExists rs, "remarks", Trim$(remarks)
    SetFieldIfExists rs, "created_at", Now()
    SetFieldIfExists rs, "created_by", created_by
    rs.Update

    rs.Bookmark = rs.LastModified
    InsertDocument = CLng(Nz(rs.Fields("document_id").Value, 0))

    rs.Close
    Set rs = Nothing
End Function

Private Sub InsertPosition( _
    ByVal db As DAO.Database, _
    ByVal DocumentId As Long, _
    ByVal LineNo As Long, _
    ByVal description As String, _
    ByVal quantity As Double, _
    ByVal unitCode As String, _
    ByVal UnitPrice As Currency, _
    ByVal vatRate As Double, _
    Optional ByVal DiscountType As String = "", _
    Optional ByVal DiscountValue As Double = 0 _
)
    Dim rs As DAO.Recordset
    Dim lineBase As Currency
    Dim lineDiscount As Currency
    Dim lineNet As Currency
    Dim lineVat As Currency
    Dim lineGross As Currency

    lineBase = CCur(Round(CDbl(quantity) * CDbl(UnitPrice), 2))

    If UCase$(Trim$(DiscountType)) = "PERCENT" Then
        lineDiscount = CCur(Round(CDbl(lineBase) * CDbl(DiscountValue) / 100#, 2))
    Else
        lineDiscount = CCur(0)
    End If

    lineNet = CCur(Round(CDbl(lineBase) - CDbl(lineDiscount), 2))
    lineVat = CCur(Round(CDbl(lineNet) * CDbl(vatRate) / 100#, 2))
    lineGross = CCur(Round(CDbl(lineNet) + CDbl(lineVat), 2))

    Set rs = db.OpenRecordset(TBL_DOC_POSITION, dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    SetFieldIfExists rs, "document_id", DocumentId
    SetFieldIfExists rs, "line_no", LineNo
    SetFieldIfExists rs, "description", Trim$(description)
    SetFieldIfExists rs, "quantity", quantity
    SetFieldIfExists rs, "unit_code", Trim$(unitCode)
    SetFieldIfExists rs, "unit_price", UnitPrice
    SetFieldIfExists rs, "vat_rate", vatRate

    SetFieldIfExists rs, "discount_type", NullIfEmpty(DiscountType)
    SetFieldIfExists rs, "discount_value", DiscountValue
    SetFieldIfExists rs, "line_base_amount", lineBase
    SetFieldIfExists rs, "line_discount_amount", lineDiscount

    SetFieldIfExists rs, "line_total_net", lineNet
    SetFieldIfExists rs, "line_total_vat", lineVat
    SetFieldIfExists rs, "line_total_gross", lineGross
    rs.Update

    rs.Close
    Set rs = Nothing
End Sub

Private Function EnsureAddress( _
    ByVal db As DAO.Database, _
    ByVal addressTypeCode As String, _
    ByVal CompanyName As String, _
    ByVal FirstName As String, _
    ByVal LastName As String, _
    ByVal Street As String, _
    ByVal HouseNo As String, _
    ByVal zipCode As String, _
    ByVal City As String, _
    ByVal countryCode As String, _
    ByVal LanguageCode As String _
) As Long
    Dim addressId As Long
    Dim wasUpdated As Boolean

    addressId = ResolveAddressId(db, addressTypeCode, CompanyName, Street, HouseNo, zipCode, City, countryCode)
    If addressId > 0 Then
        wasUpdated = UpdateAddressRequiredFields(db, addressId, FirstName, LastName, LanguageCode)
        If Not wasUpdated Then mSkippedExistingCount = mSkippedExistingCount + 1
        EnsureAddress = addressId
        Exit Function
    End If

    EnsureAddress = InsertAddress(db, addressTypeCode, CompanyName, FirstName, LastName, Street, HouseNo, zipCode, City, countryCode, LanguageCode)
    mInsertedCount = mInsertedCount + 1
End Function

Private Sub EnsureContact( _
    ByVal db As DAO.Database, _
    ByVal addressId As Long, _
    ByVal contactTypeCode As String, _
    ByVal ContactValue As String, _
    ByVal IsPrimary As Boolean, _
    ByVal remarks As String _
)
    If addressId <= 0 Then Exit Sub

    If ContactExists(db, addressId, contactTypeCode, ContactValue) Then
        mSkippedExistingCount = mSkippedExistingCount + 1
        Exit Sub
    End If

    InsertContact db, addressId, contactTypeCode, ContactValue, IsPrimary, remarks
    mInsertedCount = mInsertedCount + 1
End Sub

Private Function EnsureDocument( _
    ByVal db As DAO.Database, _
    ByVal DocumentTypeCode As String, _
    ByVal DocumentStatusCode As String, _
    ByVal DocumentNo As String, _
    ByVal DocumentDate As Date, _
    ByVal CustomerAddressId As Long, _
    ByVal CustomerName As String, _
    ByVal CurrencyCode As String, _
    ByVal VatMode As String, _
    ByVal vatRate As Double, _
    ByVal remarks As String, _
    Optional ByVal HeaderDiscountType As String = "", _
    Optional ByVal HeaderDiscountValue As Double = 0, _
    Optional ByVal LanguageCode As String = "", _
    Optional ByVal PaymentTermCode As String = "", _
    Optional ByVal PaymentTermsText As String = "" _
) As Long
    Dim documentId As Long
    Dim wasUpdated As Boolean

    documentId = ResolveDocumentIdByNo(db, DocumentNo)
    If documentId > 0 Then
        wasUpdated = UpdateDocumentRequiredFields(db, documentId, CustomerAddressId, CustomerName, CurrencyCode, VatMode, vatRate, LanguageCode, PaymentTermCode, PaymentTermsText)
        If Not wasUpdated Then mSkippedExistingCount = mSkippedExistingCount + 1
        EnsureDocument = documentId
        Exit Function
    End If

    EnsureDocument = InsertDocument(db, DocumentTypeCode, DocumentStatusCode, DocumentNo, DocumentDate, CustomerAddressId, CustomerName, CurrencyCode, VatMode, vatRate, remarks, HeaderDiscountType, HeaderDiscountValue, LanguageCode, PaymentTermCode, PaymentTermsText)
    mInsertedCount = mInsertedCount + 1
End Function

Private Sub EnsurePosition( _
    ByVal db As DAO.Database, _
    ByVal DocumentId As Long, _
    ByVal LineNo As Long, _
    ByVal description As String, _
    ByVal quantity As Double, _
    ByVal unitCode As String, _
    ByVal UnitPrice As Currency, _
    ByVal vatRate As Double, _
    Optional ByVal DiscountType As String = "", _
    Optional ByVal DiscountValue As Double = 0 _
)
    If DocumentId <= 0 Then Exit Sub

    If PositionExists(db, DocumentId, LineNo) Then
        mSkippedExistingCount = mSkippedExistingCount + 1
        Exit Sub
    End If

    InsertPosition db, DocumentId, LineNo, description, quantity, unitCode, UnitPrice, vatRate, DiscountType, DiscountValue
    MarkDocumentTotalsDirty DocumentId
    mInsertedCount = mInsertedCount + 1
End Sub

Private Sub DeleteDemoDataForReset(ByVal db As DAO.Database)
    Dim documentId As Long

    documentId = ResolveDocumentIdByNo(db, "RE-2026-0001")
    DeleteDocumentWithPositions db, documentId
    documentId = ResolveDocumentIdByNo(db, "RE-2026-0002")
    DeleteDocumentWithPositions db, documentId
    documentId = ResolveDocumentIdByNo(db, "LS-2026-0001")
    DeleteDocumentWithPositions db, documentId
    documentId = ResolveDocumentIdByNo(db, "GS-2026-0001")
    DeleteDocumentWithPositions db, documentId
    documentId = ResolveDocumentIdByNo(db, "RE-2026-0003")
    DeleteDocumentWithPositions db, documentId
    documentId = ResolveDocumentIdByNo(db, "RE-2026-0099")
    DeleteDocumentWithPositions db, documentId

    DeleteContactByAddressAndValue db, ResolveAddressId(db, "BILLING", "Muster Handel AG", "Industriestrasse", "15", "6300", "Zug", "CH"), "EMAIL", "buchhaltung@muster-handel.ch"
    DeleteContactByAddressAndValue db, ResolveAddressId(db, "SHIPPING", "Muster Handel AG - Lager Genf", "Route de Meyrin", "88", "1203", DemoCityGeneva(), "CH"), "EMAIL", "lager@muster-handel.ch"
    DeleteContactByAddressAndValue db, ResolveAddressId(db, "BILLING", "Beispiel GmbH", "Hauptstrasse", "22", "80331", DemoCityMunich(), "DE"), "EMAIL", "rechnung@beispiel-gmbh.de"
    DeleteContactByAddressAndValue db, ResolveAddressId(db, "SHIPPING", "Beispiel GmbH - Site Paris", "Rue Lafayette", "12", "75009", "Paris", "FR"), "EMAIL", "livraison@beispiel.fr"
    DeleteContactByAddressAndValue db, ResolveAddressId(db, "BILLING", "Global Components Inc.", "Market Street", "500", "94105", "San Francisco", "US"), "EMAIL", "ap@global-components.com"

    DeleteAddressByNaturalKey db, "BILLING", "Muster Handel AG", "Industriestrasse", "15", "6300", "Zug", "CH"
    DeleteAddressByNaturalKey db, "SHIPPING", "Muster Handel AG - Lager Genf", "Route de Meyrin", "88", "1203", DemoCityGeneva(), "CH"
    DeleteAddressByNaturalKey db, "BILLING", "Beispiel GmbH", "Hauptstrasse", "22", "80331", DemoCityMunich(), "DE"
    DeleteAddressByNaturalKey db, "SHIPPING", "Beispiel GmbH - Site Paris", "Rue Lafayette", "12", "75009", "Paris", "FR"
    DeleteAddressByNaturalKey db, "BILLING", "Global Components Inc.", "Market Street", "500", "94105", "San Francisco", "US"
End Sub

Private Sub UpdateDocumentTotals(ByVal db As DAO.Database)
    If mRefreshDocInvoiceChTotals Then UpdateOneDocumentTotal db, mDocInvoiceCh
    If mRefreshDocInvoiceEuTotals Then UpdateOneDocumentTotal db, mDocInvoiceEu
    If mRefreshDocDeliveryNoteTotals Then UpdateOneDocumentTotal db, mDocDeliveryNote
    If mRefreshDocCreditNoteTotals Then UpdateOneDocumentTotal db, mDocCreditNote
    If mRefreshDocInvoiceUsdTotals Then UpdateOneDocumentTotal db, mDocInvoiceUsd
    If mRefreshDocLongInvoiceTotals Then UpdateOneDocumentTotal db, mDocLongInvoice
End Sub

Private Sub UpdateOneDocumentTotal(ByVal db As DAO.Database, ByVal DocumentId As Long)
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim totalNet As Currency
    Dim totalVat As Currency
    Dim totalGross As Currency

    sql = "SELECT " & _
              "SUM([line_total_net]) AS SumNet, " & _
              "SUM([line_total_vat]) AS SumVat, " & _
              "SUM([line_total_gross]) AS SumGross " & _
              "FROM [" & TBL_DOC_POSITION & "] " & _
              "WHERE [document_id]=" & CStr(DocumentId) & ";"

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    If Not (rs.BOF And rs.EOF) Then
        totalNet = CCur(Nz(rs.Fields("SumNet").Value, 0))
        totalVat = CCur(Nz(rs.Fields("SumVat").Value, 0))
        totalGross = CCur(Nz(rs.Fields("SumGross").Value, 0))
    End If

    rs.Close
    Set rs = Nothing

    sql = "UPDATE [" & TBL_DOC_DOCUMENT & "] SET " & _
              "[total_net]=" & SqlNumber(totalNet) & ", " & _
              "[total_vat]=" & SqlNumber(totalVat) & ", " & _
              "[total_gross]=" & SqlNumber(totalGross) & " " & _
              "WHERE [document_id]=" & CStr(DocumentId) & ";"

    ExecSql db, sql
End Sub

Private Sub EnsureTenantParameter(ByVal db As DAO.Database, ByVal TenantCode As String, ByVal ParamKey As String, ByVal ParamValue As String)
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim hasTenantCodeField As Boolean
    Dim wasUpdated As Boolean

    If Not FieldExists(db, TBL_TEN_PARAMETER, "param_key") Then Exit Sub
    If Not FieldExists(db, TBL_TEN_PARAMETER, "param_value") Then Exit Sub

    hasTenantCodeField = FieldExists(db, TBL_TEN_PARAMETER, "tenant_code")

    sql = "SELECT * FROM [" & TBL_TEN_PARAMETER & "] WHERE [param_key]=" & SqlText(UCase$(Trim$(ParamKey)))
    If hasTenantCodeField Then
        sql = sql & " AND [tenant_code]=" & SqlText(UCase$(Trim$(TenantCode)))
    End If
    sql = sql & ";"
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.BOF And rs.EOF Then
        rs.AddNew
        SetFieldIfExists rs, "param_key", UCase$(Trim$(ParamKey))
        SetFieldIfExists rs, "param_value", Trim$(ParamValue)
        If hasTenantCodeField Then SetFieldIfExists rs, "tenant_code", UCase$(Trim$(TenantCode))
        SetFieldIfExists rs, "created_at", Now()
        SetFieldIfExists rs, "created_by", created_by
        SetFieldIfExists rs, "updated_at", Now()
        SetFieldIfExists rs, "updated_by", created_by
        rs.Update
        mInsertedCount = mInsertedCount + 1
    Else
        rs.Edit
        wasUpdated = SetFieldIfMissing(rs, "param_value", Trim$(ParamValue))
        If hasTenantCodeField Then wasUpdated = SetFieldIfMissing(rs, "tenant_code", UCase$(Trim$(TenantCode))) Or wasUpdated
        If wasUpdated Then
            SetFieldIfExists rs, "updated_at", Now()
            SetFieldIfExists rs, "updated_by", created_by
            rs.Update
            mUpdatedRequiredFieldsCount = mUpdatedRequiredFieldsCount + 1
        Else
            rs.CancelUpdate
            mSkippedExistingCount = mSkippedExistingCount + 1
        End If
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub UpsertProductGroup( _
    ByVal db As DAO.Database, _
    ByVal productGroupCode As String, _
    ByVal productGroupName As String, _
    ByVal descriptionText As String, _
    ByVal sortOrder As Long, _
    ByVal isActive As Boolean)

    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT * FROM [" & TBL_ART_PRODUCT_GROUP & "] " & _
          "WHERE [product_group_code]=" & SqlText(UCase$(Trim$(productGroupCode))) & ";"

    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.BOF And rs.EOF Then
        rs.AddNew
        SetFieldIfExists rs, "created_at", Now()
        SetFieldIfExists rs, "created_by", created_by
        SetFieldIfExists rs, "product_group_code", UCase$(Trim$(productGroupCode))
        SetFieldIfExists rs, "product_group_name", Trim$(productGroupName)
        SetFieldIfExists rs, "description_text", Trim$(descriptionText)
        SetFieldIfExists rs, "sort_order", sortOrder
        SetFieldIfExists rs, "is_active", isActive
        SetFieldIfExists rs, "updated_at", Now()
        SetFieldIfExists rs, "updated_by", created_by
        rs.Update
        mInsertedCount = mInsertedCount + 1
    Else
        rs.Edit
        If SetFieldIfMissing(rs, "product_group_name", Trim$(productGroupName)) _
            Or SetFieldIfMissing(rs, "description_text", Trim$(descriptionText)) _
            Or SetFieldIfMissing(rs, "sort_order", sortOrder) _
            Or SetFieldIfMissing(rs, "is_active", isActive) Then
            SetFieldIfExists rs, "updated_at", Now()
            SetFieldIfExists rs, "updated_by", created_by
            rs.Update
            mUpdatedRequiredFieldsCount = mUpdatedRequiredFieldsCount + 1
        Else
            rs.CancelUpdate
            mSkippedExistingCount = mSkippedExistingCount + 1
        End If
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub UpsertArticle( _
    ByVal db As DAO.Database, _
    ByVal articleNo As String, _
    ByVal articleName As String, _
    ByVal productGroupCode As String, _
    ByVal articleTypeCode As String, _
    ByVal unitCode As String, _
    ByVal vatCode As String, _
    ByVal purchasePrice As Double, _
    ByVal salesPrice As Double, _
    ByVal barcode As String, _
    ByVal descriptionText As String, _
    ByVal isActive As Boolean)

    Dim rs As DAO.Recordset
    Dim sql As String
    Dim productGroupId As Long

    productGroupId = ResolveProductGroupIdByCode(db, productGroupCode)
    If productGroupId <= 0 Then
        Err.Raise vbObjectError + 703, MODULE_NAME, "Required product_group_code missing: " & productGroupCode
    End If

    sql = "SELECT * FROM [" & TBL_ART_ARTICLE & "] " & _
          "WHERE [article_no]=" & SqlText(UCase$(Trim$(articleNo))) & ";"

    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.BOF And rs.EOF Then
        rs.AddNew
        SetFieldIfExists rs, "created_at", Now()
        SetFieldIfExists rs, "created_by", created_by
        SetFieldIfExists rs, "article_no", UCase$(Trim$(articleNo))
        SetFieldIfExists rs, "article_name", Trim$(articleName)
        SetFieldIfExists rs, "product_group_id", productGroupId
        SetFieldIfExists rs, "article_type_code", UCase$(Trim$(articleTypeCode))
        SetFieldIfExists rs, "unit_code", UCase$(Trim$(unitCode))
        SetFieldIfExists rs, "vat_code", UCase$(Trim$(vatCode))
        SetFieldIfExists rs, "purchase_price", purchasePrice
        SetFieldIfExists rs, "sales_price", salesPrice
        SetFieldIfExists rs, "barcode", Trim$(barcode)
        SetFieldIfExists rs, "description_text", Trim$(descriptionText)
        SetFieldIfExists rs, "is_active", isActive
        SetFieldIfExists rs, "updated_at", Now()
        SetFieldIfExists rs, "updated_by", created_by
        rs.Update
        mInsertedCount = mInsertedCount + 1
    Else
        rs.Edit
        If SetFieldIfMissing(rs, "article_name", Trim$(articleName)) _
            Or SetFieldIfMissing(rs, "product_group_id", productGroupId) _
            Or SetFieldIfMissing(rs, "article_type_code", UCase$(Trim$(articleTypeCode))) _
            Or SetFieldIfMissing(rs, "unit_code", UCase$(Trim$(unitCode))) _
            Or SetFieldIfMissing(rs, "vat_code", UCase$(Trim$(vatCode))) _
            Or SetFieldIfMissing(rs, "purchase_price", purchasePrice) _
            Or SetFieldIfMissing(rs, "sales_price", salesPrice) _
            Or SetFieldIfMissing(rs, "barcode", Trim$(barcode)) _
            Or SetFieldIfMissing(rs, "description_text", Trim$(descriptionText)) _
            Or SetFieldIfMissing(rs, "is_active", isActive) Then
            SetFieldIfExists rs, "updated_at", Now()
            SetFieldIfExists rs, "updated_by", created_by
            rs.Update
            mUpdatedRequiredFieldsCount = mUpdatedRequiredFieldsCount + 1
        Else
            rs.CancelUpdate
            mSkippedExistingCount = mSkippedExistingCount + 1
        End If
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Function ResolveProductGroupIdByCode(ByVal db As DAO.Database, ByVal productGroupCode As String) As Long
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT TOP 1 product_group_id " & _
          "FROM [" & TBL_ART_PRODUCT_GROUP & "] " & _
          "WHERE [product_group_code]=" & SqlText(UCase$(Trim$(productGroupCode))) & ";"

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then
        ResolveProductGroupIdByCode = modDaoHelper.NzLong(rs.Fields("product_group_id").Value, 0)
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    ResolveProductGroupIdByCode = 0
    Resume CleanExit
End Function

Private Function UpdateAddressRequiredFields(ByVal db As DAO.Database, ByVal addressId As Long, ByVal FirstName As String, ByVal LastName As String, ByVal LanguageCode As String) As Boolean
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim wasUpdated As Boolean

    sql = "SELECT * FROM [" & TBL_ADR_ADDRESS & "] WHERE [address_id]=" & CStr(addressId) & ";"
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)
    If rs.BOF And rs.EOF Then GoTo CleanExit

    rs.Edit
    wasUpdated = SetFieldIfMissing(rs, "first_name", Trim$(FirstName))
    wasUpdated = SetFieldIfMissing(rs, "last_name", Trim$(LastName)) Or wasUpdated
    wasUpdated = SetFieldIfMissing(rs, "language_code", Trim$(LanguageCode)) Or wasUpdated
    wasUpdated = SetFieldIfMissing(rs, "is_active", True) Or wasUpdated

    If wasUpdated Then
        SetFieldIfExists rs, "updated_at", Now()
        SetFieldIfExists rs, "updated_by", created_by
        rs.Update
        mUpdatedRequiredFieldsCount = mUpdatedRequiredFieldsCount + 1
        UpdateAddressRequiredFields = True
    Else
        rs.CancelUpdate
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Sub

Private Function UpdateDocumentRequiredFields( _
    ByVal db As DAO.Database, _
    ByVal documentId As Long, _
    ByVal CustomerAddressId As Long, _
    ByVal CustomerName As String, _
    ByVal CurrencyCode As String, _
    ByVal VatMode As String, _
    ByVal vatRate As Double, _
    ByVal LanguageCode As String, _
    ByVal PaymentTermCode As String, _
    ByVal PaymentTermsText As String)

    Dim rs As DAO.Recordset
    Dim sql As String
    Dim wasUpdated As Boolean

    sql = "SELECT * FROM [" & TBL_DOC_DOCUMENT & "] WHERE [document_id]=" & CStr(documentId) & ";"
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)
    If rs.BOF And rs.EOF Then GoTo CleanExit

    rs.Edit
    wasUpdated = SetFieldIfMissing(rs, "customer_address_id", CustomerAddressId)
    wasUpdated = SetFieldIfMissing(rs, "customer_name", Trim$(CustomerName)) Or wasUpdated
    wasUpdated = SetFieldIfMissing(rs, "currency_code", UCase$(Trim$(CurrencyCode))) Or wasUpdated
    wasUpdated = SetFieldIfMissing(rs, "vat_mode", UCase$(Trim$(VatMode))) Or wasUpdated
    wasUpdated = SetFieldIfMissing(rs, "vat_rate", vatRate) Or wasUpdated
    wasUpdated = SetFieldIfMissing(rs, "language_code", Trim$(LanguageCode)) Or wasUpdated
    wasUpdated = SetFieldIfMissing(rs, "payment_term_code", Trim$(PaymentTermCode)) Or wasUpdated
    wasUpdated = SetFieldIfMissing(rs, "payment_terms_text", Trim$(PaymentTermsText)) Or wasUpdated

    If wasUpdated Then
        SetFieldIfExists rs, "updated_at", Now()
        SetFieldIfExists rs, "updated_by", created_by
        rs.Update
        mUpdatedRequiredFieldsCount = mUpdatedRequiredFieldsCount + 1
        UpdateDocumentRequiredFields = True
    Else
        rs.CancelUpdate
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Sub

Private Function ResolveAddressId( _
    ByVal db As DAO.Database, _
    ByVal addressTypeCode As String, _
    ByVal CompanyName As String, _
    ByVal Street As String, _
    ByVal HouseNo As String, _
    ByVal zipCode As String, _
    ByVal City As String, _
    ByVal countryCode As String) As Long

    ResolveAddressId = LookupLongValue( _
        db, _
        "SELECT TOP 1 [address_id] FROM [" & TBL_ADR_ADDRESS & "] " & _
        "WHERE [address_type_code]=" & SqlText(UCase$(Trim$(addressTypeCode))) & _
        " AND [company_name]=" & SqlText(Trim$(CompanyName)) & _
        " AND [street]=" & SqlText(Trim$(Street)) & _
        " AND [house_no]=" & SqlText(Trim$(HouseNo)) & _
        " AND [zip_code]=" & SqlText(Trim$(zipCode)) & _
        " AND [city]=" & SqlText(Trim$(City)) & _
        " AND [country_code]=" & SqlText(UCase$(Trim$(countryCode))) & ";", _
        "address_id")
End Function

Private Function ResolveDocumentIdByNo(ByVal db As DAO.Database, ByVal DocumentNo As String) As Long
    ResolveDocumentIdByNo = LookupLongValue( _
        db, _
        "SELECT TOP 1 [document_id] FROM [" & TBL_DOC_DOCUMENT & "] " & _
        "WHERE [document_no]=" & SqlText(Trim$(DocumentNo)) & ";", _
        "document_id")
End Function

Private Function ContactExists(ByVal db As DAO.Database, ByVal addressId As Long, ByVal contactTypeCode As String, ByVal ContactValue As String) As Boolean
    ContactExists = (LookupLongValue( _
        db, _
        "SELECT TOP 1 [address_id] FROM [" & TBL_ADR_CONTACT & "] " & _
        "WHERE [address_id]=" & CStr(addressId) & _
        " AND [contact_type_code]=" & SqlText(UCase$(Trim$(contactTypeCode))) & _
        " AND [contact_value]=" & SqlText(Trim$(ContactValue)) & ";", _
        "address_id") > 0)
End Function

Private Function PositionExists(ByVal db As DAO.Database, ByVal DocumentId As Long, ByVal LineNo As Long) As Boolean
    PositionExists = (LookupLongValue( _
        db, _
        "SELECT TOP 1 [document_position_id] FROM [" & TBL_DOC_POSITION & "] " & _
        "WHERE [document_id]=" & CStr(DocumentId) & _
        " AND [line_no]=" & CStr(LineNo) & ";", _
        "document_position_id") > 0)
End Function

Private Sub DeleteDocumentWithPositions(ByVal db As DAO.Database, ByVal documentId As Long)
    If documentId <= 0 Then Exit Sub

    mDeletedCount = mDeletedCount + ExecuteDelete(db, "DELETE FROM [" & TBL_DOC_POSITION & "] WHERE [document_id]=" & CStr(documentId) & ";")
    mDeletedCount = mDeletedCount + ExecuteDelete(db, "DELETE FROM [" & TBL_DOC_DOCUMENT & "] WHERE [document_id]=" & CStr(documentId) & ";")
End Sub

Private Sub DeleteContactByAddressAndValue(ByVal db As DAO.Database, ByVal addressId As Long, ByVal contactTypeCode As String, ByVal ContactValue As String)
    If Not TableExists(db, TBL_ADR_CONTACT) Then Exit Sub
    If addressId <= 0 Then Exit Sub
    mDeletedCount = mDeletedCount + ExecuteDelete( _
        db, _
        "DELETE FROM [" & TBL_ADR_CONTACT & "] " & _
        "WHERE [address_id]=" & CStr(addressId) & _
        " AND [contact_type_code]=" & SqlText(UCase$(Trim$(contactTypeCode))) & _
        " AND [contact_value]=" & SqlText(Trim$(ContactValue)) & ";")
End Sub

Private Sub DeleteAddressByNaturalKey( _
    ByVal db As DAO.Database, _
    ByVal addressTypeCode As String, _
    ByVal CompanyName As String, _
    ByVal Street As String, _
    ByVal HouseNo As String, _
    ByVal zipCode As String, _
    ByVal City As String, _
    ByVal countryCode As String)

    mDeletedCount = mDeletedCount + ExecuteDelete( _
        db, _
        "DELETE FROM [" & TBL_ADR_ADDRESS & "] " & _
        "WHERE [address_type_code]=" & SqlText(UCase$(Trim$(addressTypeCode))) & _
        " AND [company_name]=" & SqlText(Trim$(CompanyName)) & _
        " AND [street]=" & SqlText(Trim$(Street)) & _
        " AND [house_no]=" & SqlText(Trim$(HouseNo)) & _
        " AND [zip_code]=" & SqlText(Trim$(zipCode)) & _
        " AND [city]=" & SqlText(Trim$(City)) & _
        " AND [country_code]=" & SqlText(UCase$(Trim$(countryCode))) & ";")
End Sub

Private Function LookupLongValue(ByVal db As DAO.Database, ByVal sql As String, ByVal fieldName As String) As Long
    Dim rs As DAO.Recordset

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then
        LookupLongValue = modDaoHelper.NzLong(rs.Fields(fieldName).Value, 0)
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Sub SetFieldIfExists(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal Value As Variant)
    If RecordsetHasField(rs, fieldName) Then
        rs.Fields(fieldName).Value = Value
    End If
End Sub

Private Function SetFieldIfMissing(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal Value As Variant) As Boolean
    If Not RecordsetHasField(rs, fieldName) Then Exit Function
    If Not IsFieldMissingValue(rs.Fields(fieldName).Value) Then Exit Function

    rs.Fields(fieldName).Value = Value
    SetFieldIfMissing = True
End Function

Private Function IsFieldMissingValue(ByVal Value As Variant) As Boolean
    If IsNull(Value) Or IsEmpty(Value) Then
        IsFieldMissingValue = True
    ElseIf VarType(Value) = vbString Then
        IsFieldMissingValue = (LenB(Trim$(CStr(Value))) = 0)
    End If
End Function

Private Function RecordsetHasField(ByVal rs As DAO.Recordset, ByVal fieldName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tmp As Variant
    tmp = rs.Fields(fieldName).Name
    RecordsetHasField = True
    Exit Function

ErrorHandler:
    RecordsetHasField = False
End Function

Private Sub RequireTable(ByVal db As DAO.Database, ByVal tableName As String)
    If Not TableExists(db, tableName) Then
        Err.Raise vbObjectError + 701, MODULE_NAME, "Required table missing: " & tableName
    End If
End Sub

Private Sub RequireField(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String)
    If Not FieldExists(db, tableName, fieldName) Then
        Err.Raise vbObjectError + 702, MODULE_NAME, "Required field missing: " & tableName & "." & fieldName
    End If
End Sub

Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tdf As DAO.tableDef

    For Each tdf In db.TableDefs
        If UCase$(Trim$(tdf.Name)) = UCase$(Trim$(tableName)) Then
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

    Dim tmp As String
    tmp = db.TableDefs(tableName).Fields(fieldName).Name
    FieldExists = True
    Exit Function

ErrorHandler:
    FieldExists = False
End Function

Private Sub ExecSql(ByVal db As DAO.Database, ByVal sql As String)
    Debug.Print sql
    db.Execute sql, dbFailOnError
End Sub

Private Function ExecuteDelete(ByVal db As DAO.Database, ByVal sql As String) As Long
    Debug.Print sql
    db.Execute sql, dbFailOnError
    ExecuteDelete = db.RecordsAffected
End Function

Private Sub MarkDocumentTotalsDirty(ByVal DocumentId As Long)
    If DocumentId <= 0 Then Exit Sub

    If DocumentId = mDocInvoiceCh Then
        mRefreshDocInvoiceChTotals = True
    ElseIf DocumentId = mDocInvoiceEu Then
        mRefreshDocInvoiceEuTotals = True
    ElseIf DocumentId = mDocDeliveryNote Then
        mRefreshDocDeliveryNoteTotals = True
    ElseIf DocumentId = mDocCreditNote Then
        mRefreshDocCreditNoteTotals = True
    ElseIf DocumentId = mDocInvoiceUsd Then
        mRefreshDocInvoiceUsdTotals = True
    ElseIf DocumentId = mDocLongInvoice Then
        mRefreshDocLongInvoiceTotals = True
    End If
End Sub

Private Function SqlText(ByVal Value As String) As String
    SqlText = "'" & Replace(Value, "'", "''") & "'"
End Function

Private Function SqlNumber(ByVal Value As Variant) As String
    If IsNull(Value) Or LenB(Trim$(CStr(Value))) = 0 Then
        SqlNumber = "NULL"
    Else
        SqlNumber = Replace(CStr(Value), ",", ".")
    End If
End Function

Private Function SqlDate(ByVal Value As Variant) As String
    If IsNull(Value) Then
        SqlDate = "NULL"
    Else
        SqlDate = "#" & Format$(CDate(Value), "yyyy-mm-dd") & "#"
    End If
End Function

Private Function NullIfEmpty(ByVal Value As String) As Variant
    If LenB(Trim$(Value)) = 0 Then
        NullIfEmpty = Null
    Else
        NullIfEmpty = Trim$(Value)
    End If
End Function

Private Sub ResetSeedRunCounters()
    mInsertedCount = 0
    mSkippedExistingCount = 0
    mUpdatedRequiredFieldsCount = 0
    mDeletedCount = 0

    mAddressBillingCh = 0
    mAddressShippingCh = 0
    mAddressBillingDe = 0
    mAddressShippingFr = 0
    mAddressBillingUs = 0

    mDocInvoiceCh = 0
    mDocInvoiceEu = 0
    mDocDeliveryNote = 0
    mDocCreditNote = 0
    mDocInvoiceUsd = 0
    mDocLongInvoice = 0

    mRefreshDocInvoiceChTotals = False
    mRefreshDocInvoiceEuTotals = False
    mRefreshDocDeliveryNoteTotals = False
    mRefreshDocCreditNoteTotals = False
    mRefreshDocInvoiceUsdTotals = False
    mRefreshDocLongInvoiceTotals = False
End Sub

Private Sub LogSeedSummary(ByVal operationName As String, ByVal tenantCode As String)
    modLoggingHandler.LogInfo MODULE_NAME & "." & operationName, _
        "tenant_code=" & tenantCode & _
        "; inserted_count=" & CStr(mInsertedCount) & _
        "; skipped_existing_count=" & CStr(mSkippedExistingCount) & _
        "; updated_required_fields_count=" & CStr(mUpdatedRequiredFieldsCount) & _
        "; deleted_count=" & CStr(mDeletedCount)
End Sub

Private Function DemoCityZurich() As String
    DemoCityZurich = "Z" & ChrW$(252) & "rich"
End Function

Private Function DemoCityGeneva() As String
    DemoCityGeneva = "Gen" & ChrW$(232) & "ve"
End Function

Private Function DemoCityMunich() As String
    DemoCityMunich = "M" & ChrW$(252) & "nchen"
End Function

Private Function DemoDeliveryNoteRemarks() As String
    DemoDeliveryNoteRemarks = "Lieferschein-Sonderfall: Lieferadresse im Fenster, Rechnungsadresse gegen" & ChrW$(252) & "ber: Muster Handel AG, Industriestrasse 15, 6300 Zug"
End Function

Private Function DemoExportInvoiceRemarks() As String
    DemoExportInvoiceRemarks = "Exportrechnung mit 0% VAT, Positionsrabatt, gro" & ChrW$(223) & "en Zahlen und Rundungstest"
End Function

Private Function DemoLongInvoiceRemarks() As String
    DemoLongInvoiceRemarks = "Langdokument f" & ChrW$(252) & "r Report- und Seitenumbruchtests"
End Function

Private Function DemoLongPositionDescription(ByVal lineNo As Long) As String
    DemoLongPositionDescription = _
        "Testposition " & Format$(lineNo, "00") & _
        " - Automatisch generierte Langbeschreibung f" & ChrW$(252) & "r Seitenumbruch-, PDF- und VAT-Tests"
End Function



