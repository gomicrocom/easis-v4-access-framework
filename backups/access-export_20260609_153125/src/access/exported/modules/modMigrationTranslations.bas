Attribute VB_Name = "modMigrationTranslations"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modMigrationTranslations
' Purpose   : Seeds missing standard report and document translations into
'             fw_translation without overwriting existing values.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modMigrationTranslations"

Private Const TABLE_FW_TRANSLATIONS As String = "fw_translation"
Private Const FIELD_TRANSLATION_KEY As String = "translation_key"
Private Const FIELD_LANGUAGE_CODE As String = "language_code"
Private Const FIELD_TRANSLATION_VALUE As String = "translation_value"
Private Const FIELD_IS_ACTIVE As String = "is_active"
Private Const FIELD_MODULE_CODE As String = "module_code"
Private Const FIELD_SORT_ORDER As String = "sort_order"
Private Const FIELD_UPDATED_AT As String = "updated_at"

Private mInsertedCount As Long
Private mSkippedCount As Long

Public Sub ApplyTranslationSeed()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database

    mInsertedCount = 0
    mSkippedCount = 0

    Set db = currentDb

    If Not TableExists(db, TABLE_FW_TRANSLATIONS) Then
        Err.Raise vbObjectError + 3100, MODULE_NAME, _
            "Required table not found: " & TABLE_FW_TRANSLATIONS
    End If

    If Not FieldExists(db, TABLE_FW_TRANSLATIONS, FIELD_TRANSLATION_KEY) Then
        Err.Raise vbObjectError + 3101, MODULE_NAME, _
            "Required field missing: " & TABLE_FW_TRANSLATIONS & "." & FIELD_TRANSLATION_KEY
    End If

    If Not FieldExists(db, TABLE_FW_TRANSLATIONS, FIELD_LANGUAGE_CODE) Then
        Err.Raise vbObjectError + 3102, MODULE_NAME, _
            "Required field missing: " & TABLE_FW_TRANSLATIONS & "." & FIELD_LANGUAGE_CODE
    End If

    If Not FieldExists(db, TABLE_FW_TRANSLATIONS, FIELD_TRANSLATION_VALUE) Then
        Err.Raise vbObjectError + 3103, MODULE_NAME, _
            "Required field missing: " & TABLE_FW_TRANSLATIONS & "." & FIELD_TRANSLATION_VALUE
    End If

    SeedDocumentTranslations db
    SeedReportTranslations db

    Debug.Print MODULE_NAME & ".ApplyTranslationSeed: inserted=" & CStr(mInsertedCount) & ", skipped=" & CStr(mSkippedCount)
    MsgBox "Translation seed completed. Inserted: " & CStr(mInsertedCount) & ", skipped: " & CStr(mSkippedCount), vbInformation, MODULE_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Translation seed failed:" & vbCrLf & Err.Number & " - " & Err.description, vbCritical, MODULE_NAME
End Sub

Private Sub SeedDocumentTranslations(ByVal db As DAO.Database)
    EnsureTranslation db, "DOCUMENT.INVOICE", "DE", "Rechnung", "REPORT", 1000
    EnsureTranslation db, "DOCUMENT.INVOICE", "EN", "Invoice", "REPORT", 1000
    EnsureTranslation db, "DOCUMENT.INVOICE", "FR", "Facture", "REPORT", 1000

    EnsureTranslation db, "DOCUMENT.CREDIT_NOTE", "DE", "Gutschrift", "REPORT", 1010
    EnsureTranslation db, "DOCUMENT.CREDIT_NOTE", "EN", "Credit note", "REPORT", 1010
    EnsureTranslation db, "DOCUMENT.CREDIT_NOTE", "FR", "Note de cr" & ChrW$(233) & "dit", "REPORT", 1010

    EnsureTranslation db, "DOCUMENT.DELIVERY_NOTE", "DE", "Lieferschein", "REPORT", 1020
    EnsureTranslation db, "DOCUMENT.DELIVERY_NOTE", "EN", "Delivery note", "REPORT", 1020
    EnsureTranslation db, "DOCUMENT.DELIVERY_NOTE", "FR", "Bon de livraison", "REPORT", 1020

    EnsureTranslation db, "DOCUMENT.QUOTE", "DE", "Offerte", "REPORT", 1030
    EnsureTranslation db, "DOCUMENT.QUOTE", "EN", "Quote", "REPORT", 1030
    EnsureTranslation db, "DOCUMENT.QUOTE", "FR", "Offre", "REPORT", 1030
End Sub

Private Sub SeedReportTranslations(ByVal db As DAO.Database)
    EnsureTranslation db, "REPORT.DOCUMENT_TYPE", "DE", "Dokumenttyp", "REPORT", 2000
    EnsureTranslation db, "REPORT.DOCUMENT_TYPE", "EN", "Document type", "REPORT", 2000
    EnsureTranslation db, "REPORT.DOCUMENT_TYPE", "FR", "Type de document", "REPORT", 2000

    EnsureTranslation db, "REPORT.STATUS", "DE", "Status", "REPORT", 2010
    EnsureTranslation db, "REPORT.STATUS", "EN", "Status", "REPORT", 2010
    EnsureTranslation db, "REPORT.STATUS", "FR", "Statut", "REPORT", 2010

    EnsureTranslation db, "REPORT.DATE", "DE", "Datum", "REPORT", 2020
    EnsureTranslation db, "REPORT.DATE", "EN", "Date", "REPORT", 2020
    EnsureTranslation db, "REPORT.DATE", "FR", "Date", "REPORT", 2020

    EnsureTranslation db, "REPORT.DOCUMENT_NO", "DE", "Nummer", "REPORT", 2030
    EnsureTranslation db, "REPORT.DOCUMENT_NO", "EN", "No.", "REPORT", 2030
    EnsureTranslation db, "REPORT.DOCUMENT_NO", "FR", "N" & ChrW$(176), "REPORT", 2030

    EnsureTranslation db, "REPORT.POS", "DE", "Pos.", "REPORT", 2040
    EnsureTranslation db, "REPORT.POS", "EN", "Pos.", "REPORT", 2040
    EnsureTranslation db, "REPORT.POS", "FR", "Pos.", "REPORT", 2040

    EnsureTranslation db, "REPORT.DESCRIPTION", "DE", "Beschreibung", "REPORT", 2050
    EnsureTranslation db, "REPORT.DESCRIPTION", "EN", "Description", "REPORT", 2050
    EnsureTranslation db, "REPORT.DESCRIPTION", "FR", "Description", "REPORT", 2050

    EnsureTranslation db, "REPORT.QUANTITY", "DE", "Anz.", "REPORT", 2060
    EnsureTranslation db, "REPORT.QUANTITY", "EN", "Qty", "REPORT", 2060
    EnsureTranslation db, "REPORT.QUANTITY", "FR", "Qt" & ChrW$(233), "REPORT", 2060

    EnsureTranslation db, "REPORT.UNIT", "DE", "Einh.", "REPORT", 2070
    EnsureTranslation db, "REPORT.UNIT", "EN", "Unit", "REPORT", 2070
    EnsureTranslation db, "REPORT.UNIT", "FR", "Unit" & ChrW$(233), "REPORT", 2070

    EnsureTranslation db, "REPORT.UNIT_PRICE", "DE", "Preis", "REPORT", 2080
    EnsureTranslation db, "REPORT.UNIT_PRICE", "EN", "Price", "REPORT", 2080
    EnsureTranslation db, "REPORT.UNIT_PRICE", "FR", "Prix", "REPORT", 2080

    EnsureTranslation db, "REPORT.DISCOUNT", "DE", "Rabatt", "REPORT", 2090
    EnsureTranslation db, "REPORT.DISCOUNT", "EN", "Discount", "REPORT", 2090
    EnsureTranslation db, "REPORT.DISCOUNT", "FR", "Rabais", "REPORT", 2090

    EnsureTranslation db, "REPORT.SURCHARGE", "DE", "Zuschlag", "REPORT", 2100
    EnsureTranslation db, "REPORT.SURCHARGE", "EN", "Surcharge", "REPORT", 2100
    EnsureTranslation db, "REPORT.SURCHARGE", "FR", "Suppl" & ChrW$(233) & "ment", "REPORT", 2100

    EnsureTranslation db, "REPORT.NET_AMOUNT", "DE", "Netto", "REPORT", 2110
    EnsureTranslation db, "REPORT.NET_AMOUNT", "EN", "Net", "REPORT", 2110
    EnsureTranslation db, "REPORT.NET_AMOUNT", "FR", "Net", "REPORT", 2110

    EnsureTranslation db, "REPORT.VAT", "DE", "MwSt.", "REPORT", 2120
    EnsureTranslation db, "REPORT.VAT", "EN", "VAT", "REPORT", 2120
    EnsureTranslation db, "REPORT.VAT", "FR", "TVA", "REPORT", 2120

    EnsureTranslation db, "REPORT.GROSS_AMOUNT", "DE", "Brutto", "REPORT", 2130
    EnsureTranslation db, "REPORT.GROSS_AMOUNT", "EN", "Gross", "REPORT", 2130
    EnsureTranslation db, "REPORT.GROSS_AMOUNT", "FR", "Brut", "REPORT", 2130

    EnsureTranslation db, "REPORT.SUBTOTAL", "DE", "Zwischensumme", "REPORT", 2140
    EnsureTranslation db, "REPORT.SUBTOTAL", "EN", "Subtotal", "REPORT", 2140
    EnsureTranslation db, "REPORT.SUBTOTAL", "FR", "Sous-total", "REPORT", 2140

    EnsureTranslation db, "REPORT.HEADER_DISCOUNT", "DE", "Kopfrabatt", "REPORT", 2150
    EnsureTranslation db, "REPORT.HEADER_DISCOUNT", "EN", "Header discount", "REPORT", 2150
    EnsureTranslation db, "REPORT.HEADER_DISCOUNT", "FR", "Rabais global", "REPORT", 2150

    EnsureTranslation db, "REPORT.HEADER_SURCHARGE", "DE", "Kopfzuschlag", "REPORT", 2160
    EnsureTranslation db, "REPORT.HEADER_SURCHARGE", "EN", "Header surcharge", "REPORT", 2160
    EnsureTranslation db, "REPORT.HEADER_SURCHARGE", "FR", "Suppl" & ChrW$(233) & "ment global", "REPORT", 2160

    EnsureTranslation db, "REPORT.TOTAL", "DE", "Total", "REPORT", 2170
    EnsureTranslation db, "REPORT.TOTAL", "EN", "Total", "REPORT", 2170
    EnsureTranslation db, "REPORT.TOTAL", "FR", "Total", "REPORT", 2170

    EnsureTranslation db, "REPORT.PAYMENT_TERMS", "DE", "Zahlungsbedingungen", "REPORT", 2180
    EnsureTranslation db, "REPORT.PAYMENT_TERMS", "EN", "Payment terms", "REPORT", 2180
    EnsureTranslation db, "REPORT.PAYMENT_TERMS", "FR", "Conditions de paiement", "REPORT", 2180

    EnsureTranslation db, "REPORT.VAT_SUMMARY", "DE", "MwSt.-Aufschlüsselung", "REPORT", 2190
    EnsureTranslation db, "REPORT.VAT_SUMMARY", "EN", "VAT summary", "REPORT", 2190
    EnsureTranslation db, "REPORT.VAT_SUMMARY", "FR", "R" & ChrW$(233) & "capitulatif TVA", "REPORT", 2190

    EnsureTranslation db, "REPORT.NO_VAT", "DE", "Keine MwSt.", "REPORT", 2200
    EnsureTranslation db, "REPORT.NO_VAT", "EN", "No VAT", "REPORT", 2200
    EnsureTranslation db, "REPORT.NO_VAT", "FR", "Sans TVA", "REPORT", 2200

    EnsureTranslation db, "REPORT.VAT_RATE", "DE", "Satz", "REPORT", 2210
    EnsureTranslation db, "REPORT.VAT_RATE", "EN", "Rate", "REPORT", 2210
    EnsureTranslation db, "REPORT.VAT_RATE", "FR", "Taux", "REPORT", 2210

    EnsureTranslation db, "REPORT.POSITION_COUNT", "DE", "Anz.", "REPORT", 2220
    EnsureTranslation db, "REPORT.POSITION_COUNT", "EN", "Count", "REPORT", 2220
    EnsureTranslation db, "REPORT.POSITION_COUNT", "FR", "Nbre", "REPORT", 2220

    EnsureTranslation db, "REPORT.VAT_BASE", "DE", "Basis", "REPORT", 2230
    EnsureTranslation db, "REPORT.VAT_BASE", "EN", "Base", "REPORT", 2230
    EnsureTranslation db, "REPORT.VAT_BASE", "FR", "Base", "REPORT", 2230

    EnsureTranslation db, "REPORT.VAT_AMOUNT", "DE", "MwSt.", "REPORT", 2240
    EnsureTranslation db, "REPORT.VAT_AMOUNT", "EN", "VAT", "REPORT", 2240
    EnsureTranslation db, "REPORT.VAT_AMOUNT", "FR", "TVA", "REPORT", 2240
End Sub

Private Sub EnsureTranslation( _
    ByVal db As DAO.Database, _
    ByVal translationKey As String, _
    ByVal languageCode As String, _
    ByVal TranslationValue As String, _
    Optional ByVal moduleCode As String = "REPORT", _
    Optional ByVal sortOrder As Long = 0)
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset

    If TranslationExists(db, translationKey, languageCode) Then
        mSkippedCount = mSkippedCount + 1
        Debug.Print MODULE_NAME & ".EnsureTranslation: skipped " & languageCode & "|" & translationKey
        Exit Sub
    End If

    Set rs = db.OpenRecordset(TABLE_FW_TRANSLATIONS, dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    rs.Fields(FIELD_TRANSLATION_KEY).Value = translationKey
    rs.Fields(FIELD_LANGUAGE_CODE).Value = languageCode
    rs.Fields(FIELD_TRANSLATION_VALUE).Value = TranslationValue
    SetFieldIfExists rs, FIELD_IS_ACTIVE, True
    SetFieldIfExists rs, FIELD_MODULE_CODE, moduleCode
    SetFieldIfExists rs, FIELD_SORT_ORDER, sortOrder
    SetFieldIfExists rs, FIELD_UPDATED_AT, Now()
    rs.Update

    mInsertedCount = mInsertedCount + 1
    Debug.Print MODULE_NAME & ".EnsureTranslation: inserted " & languageCode & "|" & translationKey

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Function TranslationExists( _
    ByVal db As DAO.Database, _
    ByVal translationKey As String, _
    ByVal languageCode As String _
) As Boolean
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlStatement As String

    sqlStatement = "SELECT [" & FIELD_TRANSLATION_KEY & "] " & _
                   "FROM [" & TABLE_FW_TRANSLATIONS & "] " & _
                   "WHERE [" & FIELD_TRANSLATION_KEY & "] = " & SqlText(translationKey) & _
                   " AND [" & FIELD_LANGUAGE_CODE & "] = " & SqlText(languageCode) & ";"

    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)
    TranslationExists = Not (rs.BOF And rs.EOF)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    TranslationExists = False
    Resume CleanExit
End Function

Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tdf As DAO.tableDef

    If db Is Nothing Then
        Exit Function
    End If

    For Each tdf In db.TableDefs
        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then
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

    If db Is Nothing Then
        Exit Function
    End If

    If Not TableExists(db, tableName) Then
        Exit Function
    End If

    Set tdf = db.TableDefs(tableName)

    For Each fld In tdf.Fields
        If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then
            FieldExists = True
            Exit Function
        End If
    Next fld

    Exit Function

ErrorHandler:
    FieldExists = False
End Function

Private Sub SetFieldIfExists(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal Value As Variant)
    If modDaoHelper.RecordsetHasField(rs, fieldName) Then
        rs.Fields(fieldName).Value = Value
    End If
End Sub