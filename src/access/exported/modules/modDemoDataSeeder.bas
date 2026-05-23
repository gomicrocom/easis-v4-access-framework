Option Compare Database
Option Explicit

'===============================================================================
' Module    : modDemoDataSeeder
' Purpose   : Rebuilds deterministic Easis v4 demo data against the current schema.
' Author    : ChatGPT
' Version   : 1.0.0
' Notes     : - Uses DAO only.
'             - Does not modify ref_* tables.
'             - Does not delete ten_parameter; demo parameters are upserted.
'             - Uses AutoNumber primary keys where present and keeps generated IDs in memory.
'===============================================================================

Private Const MODULE_NAME As String = "modDemoDataSeeder"

Private Const TBL_ADR_ADDRESS As String = "adr_address"
Private Const TBL_ADR_CONTACT As String = "adr_contact"
Private Const TBL_DOC_DOCUMENT As String = "doc_document"
Private Const TBL_DOC_POSITION As String = "doc_document_position"
Private Const TBL_TEN_PARAMETER As String = "ten_parameter"

Private Const created_by As String = "DemoDataSeeder"

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

Public Sub SeedDemoData(Optional ByVal TenantCode As String = "DEMO_CH")
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tenantCodeEffective As String

    tenantCodeEffective = UCase$(Trim$(TenantCode))
    If LenB(tenantCodeEffective) = 0 Then tenantCodeEffective = "DEMO_CH"

    Set db = CurrentDb

    DBEngine.Workspaces(0).BeginTrans

    ValidateRequiredSchema db
    DeleteExistingDemoData db
    UpsertTenantParameters db, tenantCodeEffective
    InsertAddresses db
    InsertContacts db
    InsertDocuments db
    InsertPositions db
    UpdateDocumentTotals db

    DBEngine.Workspaces(0).CommitTrans

    MsgBox "Demo-Daten wurden erfolgreich neu erstellt." & vbCrLf & _
           "TenantCode: " & tenantCodeEffective, vbInformation, "Easis v4 Demo Seeder"

CleanExit:
    Set db = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    DBEngine.Workspaces(0).Rollback
    MsgBox "Fehler beim Erstellen der Demo-Daten:" & vbCrLf & _
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

Private Sub DeleteExistingDemoData(ByVal db As DAO.Database)
    ' Delete order follows enforced 1:n relationships.
    ExecSql db, "DELETE FROM [" & TBL_DOC_POSITION & "];"
    ExecSql db, "DELETE FROM [" & TBL_DOC_DOCUMENT & "];"

    If TableExists(db, TBL_ADR_CONTACT) Then
        ExecSql db, "DELETE FROM [" & TBL_ADR_CONTACT & "];"
    End If

    ExecSql db, "DELETE FROM [" & TBL_ADR_ADDRESS & "];"
End Sub

Private Sub UpsertTenantParameters(ByVal db As DAO.Database, ByVal TenantCode As String)
    If Not TableExists(db, TBL_TEN_PARAMETER) Then Exit Sub

    UpsertTenantParameter db, "TENANT_CODE", TenantCode
    UpsertTenantParameter db, "TENANT_NAME", "Easis Demo Schweiz AG"
    UpsertTenantParameter db, "DEFAULT_LANGUAGE", "de-CH"
    UpsertTenantParameter db, "DEFAULT_CURRENCY", "CHF"
    UpsertTenantParameter db, "SENDER_NAME", "Easis Demo Schweiz AG"
    UpsertTenantParameter db, "SENDER_STREET", "Bahnhofstrasse"
    UpsertTenantParameter db, "SENDER_HOUSE_NO", "10"
    UpsertTenantParameter db, "SENDER_ZIP_CODE", "8001"
    UpsertTenantParameter db, "SENDER_CITY", "Zürich"
    UpsertTenantParameter db, "SENDER_COUNTRY_CODE", "CH"
    UpsertTenantParameter db, "SENDER_PHONE", "+41 44 123 45 67"
    UpsertTenantParameter db, "SENDER_EMAIL", "demo@easis.ch"
    UpsertTenantParameter db, "SENDER_VAT_NO", "CHE-123.456.789 MWST"
End Sub

Private Sub InsertAddresses(ByVal db As DAO.Database)
    mAddressBillingCh = InsertAddress(db, "BILLING", "Muster Handel AG", "Anna", "Keller", "Industriestrasse", "15", "6300", "Zug", "CH", "de-CH")
    mAddressShippingCh = InsertAddress(db, "SHIPPING", "Muster Handel AG - Lager Genf", "Marc", "Dubois", "Route de Meyrin", "88", "1203", "Genève", "CH", "fr-CH")
    mAddressBillingDe = InsertAddress(db, "BILLING", "Beispiel GmbH", "Thomas", "Schneider", "Hauptstrasse", "22", "80331", "München", "DE", "de-DE")
    mAddressShippingFr = InsertAddress(db, "SHIPPING", "Beispiel GmbH - Site Paris", "Claire", "Martin", "Rue Lafayette", "12", "75009", "Paris", "FR", "fr-FR")
    mAddressBillingUs = InsertAddress(db, "BILLING", "Global Components Inc.", "John", "Miller", "Market Street", "500", "94105", "San Francisco", "US", "en-US")
End Sub

Private Sub InsertContacts(ByVal db As DAO.Database)
    If Not TableExists(db, TBL_ADR_CONTACT) Then Exit Sub

    InsertContact db, mAddressBillingCh, "EMAIL", "buchhaltung@muster-handel.ch", True, "Demo billing contact"
    InsertContact db, mAddressShippingCh, "EMAIL", "lager@muster-handel.ch", True, "Demo shipping contact"
    InsertContact db, mAddressBillingDe, "EMAIL", "rechnung@beispiel-gmbh.de", True, "Demo billing contact"
    InsertContact db, mAddressShippingFr, "EMAIL", "livraison@beispiel.fr", True, "Demo shipping contact"
    InsertContact db, mAddressBillingUs, "EMAIL", "ap@global-components.com", True, "Demo billing contact"
End Sub

Private Sub InsertDocuments(ByVal db As DAO.Database)
    
    mDocInvoiceCh = InsertDocument(db, "INVOICE", "FINAL", "RE-2026-0001", DateSerial(2026, 5, 2), mAddressBillingCh, "Muster Handel AG", "CHF", "NET", 7.7, "Schweizer Rechnung / Standardfall", "", 0, "de-CH", "NET_30", "Zahlbar innert 30 Tagen netto.")
    
    mDocInvoiceEu = InsertDocument(db, "INVOICE", "FINAL", "RE-2026-0002", DateSerial(2026, 5, 2), mAddressBillingDe, "Beispiel GmbH", "EUR", "NET", 19, "EU-Rechnung mit Positionsrabatt und Kopfrabatt. Separate Lieferadresse: Beispiel GmbH - Site Paris, 12 Rue Lafayette, 75009 Paris, FR", "PERCENT", 5, "de-DE", "CASH_DISCOUNT_10_2_NET_30", "2% Skonto bei Zahlung innert 10 Tagen, ansonsten zahlbar innert 30 Tagen netto.")
    
    mDocDeliveryNote = InsertDocument(db, "DELIVERY_NOTE", "FINAL", "LS-2026-0001", DateSerial(2026, 5, 2), mAddressShippingCh, "Muster Handel AG - Lager Genf", "CHF", "NET", 7.7, "Lieferschein-Sonderfall: Lieferadresse im Fenster, Rechnungsadresse gegenüber: Muster Handel AG, Industriestrasse 15, 6300 Zug", "", 0, "de-CH", "", "")
    
    mDocCreditNote = InsertDocument(db, "CREDIT_NOTE", "FINAL", "GS-2026-0001", DateSerial(2026, 5, 2), mAddressBillingDe, "Beispiel GmbH", "EUR", "NET", 19, "Gutschrift mit negativer Position", "", 0, "de-DE", "NET_30", "Zahlbar innert 30 Tagen netto.")
    
    mDocInvoiceUsd = InsertDocument(db, "INVOICE", "FINAL", "RE-2026-0003", DateSerial(2026, 5, 2), mAddressBillingUs, "Global Components Inc.", "USD", "EXPORT", 0, "Exportrechnung mit 0% VAT, Positionsrabatt, großen Zahlen und Rundungstest", "", 0, "en-US", "PREPAYMENT", "Payable in advance.")
    
    mDocLongInvoice = InsertDocument(db, "INVOICE", "FINAL", "RE-2026-0099", DateSerial(2026, 5, 3), mAddressBillingCh, "Muster Handel AG", "CHF", "NET", 7.7, "Langdokument für Report- und Seitenumbruchtests", "", 0, "de-CH", "NET_30", "Zahlbar innert 30 Tagen netto.")

End Sub

Private Sub InsertPositions(ByVal db As DAO.Database)
    Dim i As Long
    Dim vatRate As Double
    Dim UnitPrice As Currency
    Dim quantity As Double
    Dim description As String

    InsertPosition db, mDocInvoiceCh, 1, "Beratung Architekturreview Easis v4", 4, "h", 180, 7.7
    InsertPosition db, mDocInvoiceCh, 2, "Einrichtung Reporting-Template", 1, "pauschal", 650, 7.7

    InsertPosition db, mDocInvoiceEu, 1, "Access Frontend Erweiterung", 8, "h", 145, 19
    InsertPosition db, mDocInvoiceEu, 2, "Technische Dokumentation", 1, "pauschal", 390, 19, "PERCENT", 10

    InsertPosition db, mDocDeliveryNote, 1, "Demo Hardware Box", 3, "Stk", 245, 7.7
    InsertPosition db, mDocDeliveryNote, 2, "USB-C Anschlusskabel", 6, "Stk", 18.5, 7.7

    InsertPosition db, mDocCreditNote, 1, "Gutschrift Servicekorrektur", 1, "pauschal", -250, 19

    InsertPosition db, mDocInvoiceUsd, 1, "Software license export", 5, "pcs", 499, 0
    InsertPosition db, mDocInvoiceUsd, 2, "International remote support package", 1, "package", 1250, 0, "PERCENT", 15
    InsertPosition db, mDocInvoiceUsd, 3, "Rounding and large amount test position", 12.5, "h", 1234.56, 0, "PERCENT", 2.5

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

        description = _
            "Testposition " & Format$(i, "00") & _
            " - Automatisch generierte Langbeschreibung für Seitenumbruch-, PDF- und VAT-Tests"

        InsertPosition _
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
    ByVal AddressId As Long, _
    ByVal contactTypeCode As String, _
    ByVal ContactValue As String, _
    ByVal IsPrimary As Boolean, _
    ByVal remarks As String _
)
    Dim rs As DAO.Recordset

    Set rs = db.OpenRecordset(TBL_ADR_CONTACT, dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    SetFieldIfExists rs, "address_id", AddressId
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
Private Sub UpdateDocumentTotals(ByVal db As DAO.Database)
    UpdateOneDocumentTotal db, mDocInvoiceCh
    UpdateOneDocumentTotal db, mDocInvoiceEu
    UpdateOneDocumentTotal db, mDocDeliveryNote
    UpdateOneDocumentTotal db, mDocCreditNote
    UpdateOneDocumentTotal db, mDocInvoiceUsd
    UpdateOneDocumentTotal db, mDocLongInvoice
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

Private Sub UpsertTenantParameter(ByVal db As DAO.Database, ByVal ParamKey As String, ByVal ParamValue As String)
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim hasTenantCodeField As Boolean

    If Not FieldExists(db, TBL_TEN_PARAMETER, "param_key") Then Exit Sub
    If Not FieldExists(db, TBL_TEN_PARAMETER, "param_value") Then Exit Sub

    hasTenantCodeField = FieldExists(db, TBL_TEN_PARAMETER, "tenant_code")

    sql = "SELECT * FROM [" & TBL_TEN_PARAMETER & "] WHERE [param_key]=" & SqlText(UCase$(Trim$(ParamKey))) & ";"
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.BOF And rs.EOF Then
        rs.AddNew
        SetFieldIfExists rs, "param_key", UCase$(Trim$(ParamKey))
    Else
        rs.Edit
    End If

    SetFieldIfExists rs, "param_value", Trim$(ParamValue)
    If hasTenantCodeField Then SetFieldIfExists rs, "tenant_code", "DEMO_CH"
    SetFieldIfExists rs, "created_at", Now()
    SetFieldIfExists rs, "created_by", created_by
    SetFieldIfExists rs, "updated_at", Now()
    SetFieldIfExists rs, "updated_by", created_by
    rs.Update

    rs.Close
    Set rs = Nothing
End Sub

Private Sub SetFieldIfExists(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal Value As Variant)
    If RecordsetHasField(rs, fieldName) Then
        rs.Fields(fieldName).Value = Value
    End If
End Sub

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

Private Function SqlText(ByVal Value As Variant) As String
    If IsNull(Value) Then
        SqlText = "NULL"
    Else
        SqlText = "'" & Replace(CStr(Value), "'", "''") & "'"
    End If
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



