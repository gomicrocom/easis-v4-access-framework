Attribute VB_Name = "modTranslationService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modTranslationService
' Purpose   : Centralized translation lookup service for runtime language support.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modTranslationService"
Private Const FALLBACK_LANGUAGE As String = "EN"
Private Const TABLE_FW_TRANSLATIONS As String = "fw_translation"
Private Const FIELD_TRANSLATION_KEY As String = "translation_key"
Private Const FIELD_LANGUAGE_CODE As String = "language_code"
Private Const FIELD_TRANSLATION_VALUE As String = "translation_value"
Private Const FIELD_IS_ACTIVE As String = "is_active"

Private mTranslations As Object
Private mCurrentLanguage As String
Private mDefaultLanguage As String

Public Sub InitializeTranslations()
    On Error GoTo ErrorHandler

    Dim preservedLanguage As String

    preservedLanguage = NormalizeLanguageCode(mCurrentLanguage)

    Set mTranslations = CreateObject("Scripting.Dictionary")
    mTranslations.CompareMode = vbTextCompare

    mDefaultLanguage = ResolveDefaultLanguage()
    If LenB(preservedLanguage) > 0 Then
        mCurrentLanguage = preservedLanguage
    Else
        mCurrentLanguage = mDefaultLanguage
    End If

    LoadTranslations

    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeTranslations", _
        "Translations initialized for language '" & mCurrentLanguage & "'."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "InitializeTranslations", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Public Sub SetCurrentLanguage(ByVal languageCode As String)
    On Error GoTo ErrorHandler

    Dim normalizedLanguage As String

    If LenB(mDefaultLanguage) = 0 Then
        mDefaultLanguage = ResolveDefaultLanguage()
    End If

    normalizedLanguage = NormalizeLanguageCode(languageCode)
    If LenB(normalizedLanguage) = 0 Then
        mCurrentLanguage = mDefaultLanguage
    Else
        mCurrentLanguage = normalizedLanguage
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SetCurrentLanguage", Err
End Sub

Public Function GetCurrentLanguage() As String
    On Error GoTo ErrorHandler

    If LenB(mCurrentLanguage) = 0 Then
        mCurrentLanguage = ResolveDefaultLanguage()
    End If

    GetCurrentLanguage = mCurrentLanguage
    Exit Function

ErrorHandler:
    GetCurrentLanguage = FALLBACK_LANGUAGE
    modErrorHandler.HandleError MODULE_NAME, "GetCurrentLanguage", Err
End Function

Public Function T(ByVal TextKey As String, Optional ByVal Fallback As String = "") As String
    On Error GoTo ErrorHandler

    Dim normalizedKey As String
    Dim currentLanguageCode As String
    Dim currentBaseLanguageCode As String
    Dim defaultLanguageCode As String
    Dim defaultBaseLanguageCode As String
    Dim translatedValue As String

    normalizedKey = NormalizeTextKey(TextKey)
    If LenB(normalizedKey) = 0 Then
        T = Fallback
        Exit Function
    End If

    EnsureTranslationStore

    currentLanguageCode = GetCurrentLanguage()
    currentBaseLanguageCode = BaseLanguageCode(currentLanguageCode)
    translatedValue = LookupTranslation(currentLanguageCode, normalizedKey)
    If LenB(translatedValue) > 0 Then
        T = translatedValue
        Exit Function
    End If

    If StrComp(currentBaseLanguageCode, currentLanguageCode, vbTextCompare) <> 0 Then
        translatedValue = LookupTranslation(currentBaseLanguageCode, normalizedKey)
        If LenB(translatedValue) > 0 Then
            T = translatedValue
            Exit Function
        End If
    End If

    If LenB(mDefaultLanguage) = 0 Then
        mDefaultLanguage = ResolveDefaultLanguage()
    End If

    defaultLanguageCode = mDefaultLanguage
    defaultBaseLanguageCode = BaseLanguageCode(defaultLanguageCode)

    If StrComp(currentLanguageCode, defaultLanguageCode, vbTextCompare) <> 0 Then
        translatedValue = LookupTranslation(defaultLanguageCode, normalizedKey)
        If LenB(translatedValue) > 0 Then
            T = translatedValue
            Exit Function
        End If
    End If

    If StrComp(defaultBaseLanguageCode, currentLanguageCode, vbTextCompare) <> 0 _
        And StrComp(defaultBaseLanguageCode, currentBaseLanguageCode, vbTextCompare) <> 0 _
        And StrComp(defaultBaseLanguageCode, defaultLanguageCode, vbTextCompare) <> 0 Then
        translatedValue = LookupTranslation(defaultBaseLanguageCode, normalizedKey)
        If LenB(translatedValue) > 0 Then
            T = translatedValue
            Exit Function
        End If
    End If

    If StrComp(FALLBACK_LANGUAGE, currentLanguageCode, vbTextCompare) <> 0 _
        And StrComp(FALLBACK_LANGUAGE, currentBaseLanguageCode, vbTextCompare) <> 0 _
        And StrComp(FALLBACK_LANGUAGE, defaultLanguageCode, vbTextCompare) <> 0 _
        And StrComp(FALLBACK_LANGUAGE, defaultBaseLanguageCode, vbTextCompare) <> 0 Then
        translatedValue = LookupTranslation(FALLBACK_LANGUAGE, normalizedKey)
        If LenB(translatedValue) > 0 Then
            T = translatedValue
            Exit Function
        End If
    End If

    If LenB(Fallback) > 0 Then
        T = Fallback
    Else
        T = TextKey
    End If
    Exit Function

ErrorHandler:
    If LenB(Fallback) > 0 Then
        T = Fallback
    Else
        T = TextKey
    End If
    modErrorHandler.HandleError MODULE_NAME, "T", Err
End Function

Private Sub EnsureTranslationStore()
    On Error GoTo ErrorHandler

    If mTranslations Is Nothing Then
        InitializeTranslations
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureTranslationStore", Err
End Sub

Private Sub LoadTranslations()
    On Error GoTo ErrorHandler

    Dim loadedCount As Long

    LoadFallbackTranslations

    loadedCount = LoadTranslationsFromTable()
    If loadedCount > 0 Then
        modLoggingHandler.LogInfo MODULE_NAME & ".LoadTranslations", _
            CStr(loadedCount) & " translation(s) loaded from table '" & TABLE_FW_TRANSLATIONS & "' on top of fallback translations."
        Exit Sub
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadTranslations", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Function ResolveDefaultLanguage() As String
    On Error GoTo ErrorHandler

    Dim resolvedLanguage As String

    resolvedLanguage = NormalizeLanguageCode(modFwTranslationRuntime.ResolveCurrentLanguageCode())

    If LenB(resolvedLanguage) = 0 Then
        resolvedLanguage = FALLBACK_LANGUAGE
    End If

    ResolveDefaultLanguage = resolvedLanguage
    Exit Function

ErrorHandler:
    ResolveDefaultLanguage = FALLBACK_LANGUAGE
    modErrorHandler.HandleError MODULE_NAME, "ResolveDefaultLanguage", Err
End Function

Private Function LoadTranslationsFromTable() As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlStatement As String
    Dim hasIsActiveField As Boolean

    If Not TranslationTableExists(TABLE_FW_TRANSLATIONS) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".LoadTranslationsFromTable", _
            "Translation table '" & TABLE_FW_TRANSLATIONS & "' not found. Falling back to minimal internal translations."
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Err.Raise vbObjectError + 2800, MODULE_NAME & ".LoadTranslationsFromTable", _
            "Current database could not be resolved."
    End If

    sqlStatement = "SELECT * FROM [" & TABLE_FW_TRANSLATIONS & "];"
    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)

    If rs.EOF Then
        modLoggingHandler.LogWarning MODULE_NAME & ".LoadTranslationsFromTable", _
            "Translation table '" & TABLE_FW_TRANSLATIONS & "' contains no rows. Falling back to minimal internal translations."
        GoTo CleanExit
    End If

    If Not modDaoHelper.RecordsetHasField(rs, FIELD_TRANSLATION_KEY) _
        Or Not modDaoHelper.RecordsetHasField(rs, FIELD_LANGUAGE_CODE) _
        Or Not modDaoHelper.RecordsetHasField(rs, FIELD_TRANSLATION_VALUE) Then
        Err.Raise vbObjectError + 2801, MODULE_NAME & ".LoadTranslationsFromTable", _
            "Translation table '" & TABLE_FW_TRANSLATIONS & "' is missing one or more required fields."
    End If

    hasIsActiveField = modDaoHelper.RecordsetHasField(rs, FIELD_IS_ACTIVE)

    Do While Not rs.EOF
        If Not hasIsActiveField Or modDaoHelper.NzBoolean(rs.Fields(FIELD_IS_ACTIVE).Value, True) Then
            AddTranslation modDaoHelper.NzString(rs.Fields(FIELD_LANGUAGE_CODE).Value), _
                           modDaoHelper.NzString(rs.Fields(FIELD_TRANSLATION_KEY).Value), _
                           modDaoHelper.NzString(rs.Fields(FIELD_TRANSLATION_VALUE).Value)

            If LenB(NormalizeLanguageCode(modDaoHelper.NzString(rs.Fields(FIELD_LANGUAGE_CODE).Value))) > 0 _
                And LenB(NormalizeTextKey(modDaoHelper.NzString(rs.Fields(FIELD_TRANSLATION_KEY).Value))) > 0 Then
                LoadTranslationsFromTable = LoadTranslationsFromTable + 1
            End If
        End If
        rs.MoveNext
    Loop

    If LoadTranslationsFromTable = 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".LoadTranslationsFromTable", _
            "Translation table '" & TABLE_FW_TRANSLATIONS & "' contains no active translation rows. Falling back to minimal internal translations."
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
    Set db = Nothing
    modErrorHandler.HandleError MODULE_NAME, "LoadTranslationsFromTable", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Private Sub LoadFallbackTranslations()
    On Error GoTo ErrorHandler

    AddTranslation "EN", "APP_TITLE", "Easis Version 4"
    AddTranslation "EN", "DOCUMENT", "Document"
    AddTranslation "EN", "CUSTOMER", "Customer"
    AddTranslation "EN", "TOTAL", "Total"
    AddTranslation "EN", "DOCUMENT.INVOICE", "Invoice"
    AddTranslation "EN", "DOCUMENT.CREDIT_NOTE", "Credit note"
    AddTranslation "EN", "DOCUMENT.DELIVERY_NOTE", "Delivery note"
    AddTranslation "EN", "DOCUMENT.QUOTE", "Quote"
    AddTranslation "EN", "REPORT.DATE", "Date"
    AddTranslation "EN", "REPORT.DOCUMENT_NO", "No."
    AddTranslation "EN", "REPORT.DOCUMENT_TYPE", "Document type"
    AddTranslation "EN", "REPORT.STATUS", "Status"
    AddTranslation "EN", "REPORT.POS", "Pos."
    AddTranslation "EN", "REPORT.DESCRIPTION", "Description"
    AddTranslation "EN", "REPORT.QUANTITY", "Qty"
    AddTranslation "EN", "REPORT.UNIT", "Unit"
    AddTranslation "EN", "REPORT.UNIT_PRICE", "Price"
    AddTranslation "EN", "REPORT.DISCOUNT", "Discount"
    AddTranslation "EN", "REPORT.SURCHARGE", "Surcharge"
    AddTranslation "EN", "REPORT.NET_AMOUNT", "Net"
    AddTranslation "EN", "REPORT.VAT", "VAT"
    AddTranslation "EN", "REPORT.GROSS_AMOUNT", "Gross"
    AddTranslation "EN", "REPORT.SUBTOTAL", "Subtotal"
    AddTranslation "EN", "REPORT.HEADER_DISCOUNT", "Header discount"
    AddTranslation "EN", "REPORT.HEADER_SURCHARGE", "Header surcharge"
    AddTranslation "EN", "REPORT.TOTAL", "Total"
    AddTranslation "EN", "REPORT.PAYMENT_TERMS", "Payment terms"
    AddTranslation "EN", "REPORT.VAT_SUMMARY", "VAT summary"
    AddTranslation "EN", "REPORT.NO_VAT", "No VAT"
    AddTranslation "EN", "REPORT.VAT_RATE", "Rate"
    AddTranslation "EN", "REPORT.VAT_BASE", "Base"
    AddTranslation "EN", "REPORT.VAT_AMOUNT", "VAT"
    AddTranslation "EN", "REPORT.POSITION_COUNT", "Count"

    AddTranslation "DE", "APP_TITLE", "Easis Version 4"
    AddTranslation "DE", "DOCUMENT", "Beleg"
    AddTranslation "DE", "CUSTOMER", "Kunde"
    AddTranslation "DE", "TOTAL", "Total"
    AddTranslation "DE", "DOCUMENT.INVOICE", "Rechnung"
    AddTranslation "DE", "DOCUMENT.CREDIT_NOTE", "Gutschrift"
    AddTranslation "DE", "DOCUMENT.DELIVERY_NOTE", "Lieferschein"
    AddTranslation "DE", "DOCUMENT.QUOTE", "Offerte"
    AddTranslation "DE", "REPORT.DATE", "Datum"
    AddTranslation "DE", "REPORT.DOCUMENT_NO", "Nummer"
    AddTranslation "DE", "REPORT.DOCUMENT_TYPE", "Dokumenttyp"
    AddTranslation "DE", "REPORT.STATUS", "Status"
    AddTranslation "DE", "REPORT.POS", "Pos."
    AddTranslation "DE", "REPORT.DESCRIPTION", "Beschreibung"
    AddTranslation "DE", "REPORT.QUANTITY", "Anz."
    AddTranslation "DE", "REPORT.UNIT", "Einh."
    AddTranslation "DE", "REPORT.UNIT_PRICE", "Preis"
    AddTranslation "DE", "REPORT.DISCOUNT", "Rabatt"
    AddTranslation "DE", "REPORT.SURCHARGE", "Zuschlag"
    AddTranslation "DE", "REPORT.NET_AMOUNT", "Netto"
    AddTranslation "DE", "REPORT.VAT", "MwSt."
    AddTranslation "DE", "REPORT.GROSS_AMOUNT", "Brutto"
    AddTranslation "DE", "REPORT.SUBTOTAL", "Zwischensumme"
    AddTranslation "DE", "REPORT.HEADER_DISCOUNT", "Kopfrabatt"
    AddTranslation "DE", "REPORT.HEADER_SURCHARGE", "Kopfzuschlag"
    AddTranslation "DE", "REPORT.TOTAL", "Total"
    AddTranslation "DE", "REPORT.PAYMENT_TERMS", "Zahlungsbedingungen"
    AddTranslation "DE", "REPORT.VAT_SUMMARY", "MwSt.-Zusammenfassung"
    AddTranslation "DE", "REPORT.NO_VAT", "Keine MwSt."
    AddTranslation "DE", "REPORT.VAT_RATE", "Satz"
    AddTranslation "DE", "REPORT.VAT_BASE", "Basis"
    AddTranslation "DE", "REPORT.VAT_AMOUNT", "MwSt."
    AddTranslation "DE", "REPORT.POSITION_COUNT", "Anz."

    AddTranslation "FR", "DOCUMENT.INVOICE", "Facture"
    AddTranslation "FR", "DOCUMENT.CREDIT_NOTE", "Note de cr" & ChrW$(233) & "dit"
    AddTranslation "FR", "DOCUMENT.DELIVERY_NOTE", "Bon de livraison"
    AddTranslation "FR", "DOCUMENT.QUOTE", "Offre"
    AddTranslation "FR", "REPORT.DATE", "Date"
    AddTranslation "FR", "REPORT.DOCUMENT_NO", "N" & ChrW$(176)
    AddTranslation "FR", "REPORT.DOCUMENT_TYPE", "Type de document"
    AddTranslation "EN", "REPORT.STATUS", "Statut"
    AddTranslation "FR", "REPORT.POS", "Pos."
    AddTranslation "FR", "REPORT.DESCRIPTION", "Description"
    AddTranslation "FR", "REPORT.QUANTITY", "Qt" & ChrW$(233)
    AddTranslation "FR", "REPORT.UNIT", "Unit" & ChrW$(233)
    AddTranslation "FR", "REPORT.UNIT_PRICE", "Prix"
    AddTranslation "FR", "REPORT.DISCOUNT", "Rabais"
    AddTranslation "FR", "REPORT.SURCHARGE", "Suppl" & ChrW$(233) & "ment"
    AddTranslation "FR", "REPORT.NET_AMOUNT", "Net"
    AddTranslation "FR", "REPORT.VAT", "TVA"
    AddTranslation "FR", "REPORT.GROSS_AMOUNT", "Brut"
    AddTranslation "FR", "REPORT.SUBTOTAL", "Sous-total"
    AddTranslation "FR", "REPORT.HEADER_DISCOUNT", "Rabais global"
    AddTranslation "FR", "REPORT.HEADER_SURCHARGE", "Suppl" & ChrW$(233) & "ment global"
    AddTranslation "FR", "REPORT.TOTAL", "Total"
    AddTranslation "FR", "REPORT.PAYMENT_TERMS", "Conditions de paiement"
    AddTranslation "FR", "REPORT.VAT_SUMMARY", "R" & ChrW$(233) & "capitulatif TVA"
    AddTranslation "FR", "REPORT.NO_VAT", "Sans TVA"
    AddTranslation "FR", "REPORT.VAT_RATE", "Taux"
    AddTranslation "FR", "REPORT.VAT_BASE", "Base"
    AddTranslation "FR", "REPORT.VAT_AMOUNT", "TVA"
    AddTranslation "FR", "REPORT.POSITION_COUNT", "Nbre"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadFallbackTranslations", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub AddTranslation(ByVal languageCode As String, ByVal TextKey As String, ByVal textValue As String)
    On Error GoTo ErrorHandler

    Dim compositeKey As String

    compositeKey = BuildTranslationKey(languageCode, TextKey)
    If LenB(compositeKey) = 0 Then
        Exit Sub
    End If

    EnsureTranslationStore
    mTranslations(compositeKey) = textValue
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "AddTranslation", Err
End Sub

Private Function LookupTranslation(ByVal languageCode As String, ByVal TextKey As String) As String
    On Error GoTo ErrorHandler

    Dim compositeKey As String

    compositeKey = BuildTranslationKey(languageCode, TextKey)
    If LenB(compositeKey) = 0 Then
        Exit Function
    End If

    If mTranslations.Exists(compositeKey) Then
        LookupTranslation = CStr(mTranslations(compositeKey))
    End If
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LookupTranslation", Err
End Function

Private Function BuildTranslationKey(ByVal languageCode As String, ByVal TextKey As String) As String
    Dim normalizedLanguage As String
    Dim normalizedTextKey As String

    normalizedLanguage = NormalizeLanguageCode(languageCode)
    normalizedTextKey = NormalizeTextKey(TextKey)

    If LenB(normalizedLanguage) = 0 Or LenB(normalizedTextKey) = 0 Then
        Exit Function
    End If

    BuildTranslationKey = normalizedLanguage & "|" & normalizedTextKey
End Function

Private Function BaseLanguageCode(ByVal languageCode As String) As String
    Dim normalizedLanguage As String
    Dim separatorPosition As Long

    normalizedLanguage = NormalizeLanguageCode(languageCode)
    If LenB(normalizedLanguage) = 0 Then
        Exit Function
    End If

    separatorPosition = InStr(1, normalizedLanguage, "-", vbTextCompare)
    If separatorPosition > 1 Then
        BaseLanguageCode = Left$(normalizedLanguage, separatorPosition - 1)
    Else
        BaseLanguageCode = normalizedLanguage
    End If
End Function

Private Function NormalizeLanguageCode(ByVal languageCode As String) As String
    Dim normalizedLanguage As String
    Dim languageParts() As String
    Dim i As Long

    normalizedLanguage = Trim$(languageCode)
    If LenB(normalizedLanguage) = 0 Then
        Exit Function
    End If

    normalizedLanguage = Replace(normalizedLanguage, "_", "-")
    languageParts = Split(normalizedLanguage, "-")

    For i = LBound(languageParts) To UBound(languageParts)
        languageParts(i) = UCase$(Trim$(languageParts(i)))
    Next i

    NormalizeLanguageCode = languageParts(LBound(languageParts))
    For i = LBound(languageParts) + 1 To UBound(languageParts)
        If LenB(languageParts(i)) > 0 Then
            NormalizeLanguageCode = NormalizeLanguageCode & "-" & languageParts(i)
        End If
    Next i
End Function

Private Function NormalizeTextKey(ByVal TextKey As String) As String
    NormalizeTextKey = UCase$(Trim$(TextKey))
End Function

Private Function TranslationTableExists(ByVal tableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tableDef As DAO.tableDef
    Dim normalizedTableName As String

    normalizedTableName = UCase$(Trim$(tableName))
    If LenB(normalizedTableName) = 0 Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Err.Raise vbObjectError + 2802, MODULE_NAME & ".TranslationTableExists", _
            "Current database could not be resolved."
    End If

    For Each tableDef In db.TableDefs
        If UCase$(tableDef.Name) = normalizedTableName Then
            TranslationTableExists = True
            Exit For
        End If
    Next tableDef

    Set tableDef = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    TranslationTableExists = False
    modErrorHandler.HandleError MODULE_NAME, "TranslationTableExists", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Public Function TEx(ByVal TextKey As String, ByVal Fallback As String, ParamArray Args() As Variant) As String
    On Error GoTo ErrorHandler

    Dim resultText As String
    Dim i As Long

    resultText = T(TextKey, Fallback)

    For i = LBound(Args) To UBound(Args)
        resultText = Replace(resultText, "{" & CStr(i) & "}", NzArgumentValue(Args(i)))
    Next i

    TEx = resultText
    Exit Function

ErrorHandler:
    TEx = T(TextKey, Fallback)
End Function

Public Sub DebugPrintReportTranslations(Optional ByVal languageCode As String = "de-CH")
    On Error GoTo ErrorHandler

    SetCurrentLanguage languageCode

    Debug.Print "Language=" & GetCurrentLanguage()
    Debug.Print "DOCUMENT.INVOICE=" & T("DOCUMENT.INVOICE", "Invoice")
    Debug.Print "DOCUMENT.DELIVERY_NOTE=" & T("DOCUMENT.DELIVERY_NOTE", "Delivery note")
    Debug.Print "REPORT.PAYMENT_TERMS=" & T("REPORT.PAYMENT_TERMS", "Payment terms")
    Debug.Print "REPORT.TOTAL=" & T("REPORT.TOTAL", "Total")
    Debug.Print "REPORT.VAT_SUMMARY=" & T("REPORT.VAT_SUMMARY", "VAT summary")
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "DebugPrintReportTranslations", Err
End Sub

Private Function NzArgumentValue(ByVal Value As Variant) As String
    On Error GoTo SafeExit

    If IsNull(Value) Or IsEmpty(Value) Then
        NzArgumentValue = vbNullString
    Else
        NzArgumentValue = CStr(Value)
    End If
    Exit Function

SafeExit:
    NzArgumentValue = vbNullString
End Function
