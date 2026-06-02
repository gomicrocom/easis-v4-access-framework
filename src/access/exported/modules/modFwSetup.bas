Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFwSetup
' Purpose   : Provides initialization and seeding routines for framework data
'             such as translations, tag help definitions, and demo content.
' Author    : Codex
' Version   : 1.7.0
' Notes     : Safe to re-run. Existing data is preserved by default.
'===============================================================================

Private Const MODULE_NAME As String = "modFwSetup"

Public Sub SeedTranslations()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    ' ===== EN =====
    InsertTranslation db, "EN", "MSG_REQUIRED_FIELDS_MISSING", "Please fill in all required fields.", True
    InsertTranslation db, "EN", "MSG_INVALID_FIELD_VALUES", "Please correct invalid field values.", True
    InsertTranslation db, "EN", "MSG_MODULE_NOT_ACTIVE", "The required module is not active", True
    InsertTranslation db, "EN", "MSG_ROLE_NOT_ALLOWED", "You do not have permission to open this form", True

    InsertTranslation db, "EN", "ERR_REQUIRED", "is required", True
    InsertTranslation db, "EN", "ERR_NUMERIC", "must be a number", True
    InsertTranslation db, "EN", "ERR_INTEGER", "must be an integer", True
    InsertTranslation db, "EN", "ERR_MIN", "must be >= {0}", True
    InsertTranslation db, "EN", "ERR_MAX", "must be <= {0}", True
    InsertTranslation db, "EN", "ERR_MINLEN", "minimum length is {0}", True
    InsertTranslation db, "EN", "ERR_MAXLEN", "maximum length is {0}", True
    InsertTranslation db, "EN", "ERR_DATE", "must be a valid date", True

    ' Demo/UI texts
    InsertTranslation db, "EN", "APP_TITLE", "Easis Version 4", True
    InsertTranslation db, "EN", "DOCUMENT", "Document", True
    InsertTranslation db, "EN", "CUSTOMER", "Customer", True
    InsertTranslation db, "EN", "TOTAL", "Total", True

    ' ===== DE =====
    InsertTranslation db, "DE", "MSG_REQUIRED_FIELDS_MISSING", "Bitte fuellen Sie alle Pflichtfelder aus.", True
    InsertTranslation db, "DE", "MSG_INVALID_FIELD_VALUES", "Bitte korrigieren Sie die ungueltigen Feldwerte.", True
    InsertTranslation db, "DE", "MSG_MODULE_NOT_ACTIVE", "Das erforderliche Modul ist nicht aktiv", True
    InsertTranslation db, "DE", "MSG_ROLE_NOT_ALLOWED", "Sie sind nicht berechtigt, dieses Formular zu oeffnen", True

    InsertTranslation db, "DE", "ERR_REQUIRED", "ist erforderlich", True
    InsertTranslation db, "DE", "ERR_NUMERIC", "muss eine Zahl sein", True
    InsertTranslation db, "DE", "ERR_INTEGER", "muss eine ganze Zahl sein", True
    InsertTranslation db, "DE", "ERR_MIN", "muss >= {0} sein", True
    InsertTranslation db, "DE", "ERR_MAX", "muss <= {0} sein", True
    InsertTranslation db, "DE", "ERR_MINLEN", "Mindestlaenge ist {0}", True
    InsertTranslation db, "DE", "ERR_MAXLEN", "Maximallaenge ist {0}", True
    InsertTranslation db, "DE", "ERR_DATE", "muss ein gueltiges Datum sein", True

    ' Demo/UI texts
    InsertTranslation db, "DE", "APP_TITLE", "Easis Version 4", True
    InsertTranslation db, "DE", "DOCUMENT", "Beleg", True
    InsertTranslation db, "DE", "CUSTOMER", "Kunde", True
    InsertTranslation db, "DE", "TOTAL", "Total", True

    InsertAddressTypeTranslations db
    InsertSalutationTranslations db
    InsertAddressingModeTranslations db
    InsertContactTypeTranslations db
    InsertUnitTranslations db
    InsertVatCodeTranslations db
    InsertShellTranslations db

    MsgBox "fw_translation wurde erfolgreich ergaenzt.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Initialisieren von fw_translation: " & Err.description, vbExclamation
End Sub

Public Sub SeedShellTranslations()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    InsertShellTranslations db

    MsgBox "Shell-, Dashboard- und Navigations-Uebersetzungen wurden erfolgreich initialisiert.", vbInformation, MODULE_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Initialisieren der Shell-Uebersetzungen: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Public Sub SeedVatCodeReference()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    db.Execute "DELETE FROM ref_vat_code", dbFailOnError
    InsertVatCode db, "CH_STANDARD", "VAT.CH.STANDARD", 7.7, "CH", #1/1/2018#, Null, 10, True
    InsertVatCode db, "CH_REDUCED", "VAT.CH.REDUCED", 2.5, "CH", #1/1/2018#, Null, 20, True
    InsertVatCode db, "CH_SPECIAL", "VAT.CH.SPECIAL", 3.7, "CH", #1/1/2018#, Null, 30, True
    InsertVatCode db, "CH_ZERO", "VAT.CH.ZERO", 0, "CH", #1/1/2018#, Null, 40, True

    InsertVatCodeTranslations db

    MsgBox "ref_vat_code und VAT-Uebersetzungen wurden erfolgreich initialisiert.", vbInformation, MODULE_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Initialisieren von ref_vat_code: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Public Sub SeedUnitReference()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    db.Execute "DELETE FROM ref_unit", dbFailOnError
    InsertUnit db, "PCS", "UNIT.PCS", 10, True
    InsertUnit db, "H", "UNIT.HOUR", 20, True
    InsertUnit db, "KG", "UNIT.KG", 30, True
    InsertUnit db, "M", "UNIT.METER", 40, True
    InsertUnit db, "L", "UNIT.LITER", 50, True
    InsertUnit db, "PK", "UNIT.PACKAGE", 60, True

    InsertUnitTranslations db

    MsgBox "ref_unit und UNIT-Uebersetzungen wurden erfolgreich initialisiert.", vbInformation, MODULE_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Initialisieren von ref_unit: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Public Sub SeedAddressTypeReference()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    db.Execute "DELETE FROM ref_address_type", dbFailOnError
    InsertAddressType db, "PRIVATE", "ADDRESS_TYPE.PRIVATE", 10, True
    InsertAddressType db, "COMPANY", "ADDRESS_TYPE.COMPANY", 20, True
    InsertAddressType db, "CUSTOMER", "ADDRESS_TYPE.CUSTOMER", 30, True
    InsertAddressType db, "SUPPLIER", "ADDRESS_TYPE.SUPPLIER", 40, True
    InsertAddressType db, "PARTNER", "ADDRESS_TYPE.PARTNER", 50, True
    InsertAddressType db, "EMPLOYEE", "ADDRESS_TYPE.EMPLOYEE", 60, True
    InsertAddressType db, "OTHER", "ADDRESS_TYPE.OTHER", 90, True

    InsertAddressTypeTranslations db

    MsgBox "ref_address_type und ADDRESS_TYPE-Uebersetzungen wurden erfolgreich initialisiert.", vbInformation, MODULE_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Initialisieren von ref_address_type: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Public Sub SeedAddressPersonalizationReference()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    db.Execute "DELETE FROM ref_salutation", dbFailOnError
    db.Execute "DELETE FROM ref_addressing_mode", dbFailOnError
    InsertSalutation db, "MR", "SALUTATION.MR", 10, True
    InsertSalutation db, "MS", "SALUTATION.MS", 20, True
    InsertSalutation db, "COMPANY", "SALUTATION.COMPANY", 30, True
    InsertSalutation db, "NEUTRAL", "SALUTATION.NEUTRAL", 90, True

    InsertAddressingMode db, "FORMAL", "ADDRESSING_MODE.FORMAL", 10, True
    InsertAddressingMode db, "INFORMAL", "ADDRESSING_MODE.INFORMAL", 20, True

    InsertSalutationTranslations db
    InsertAddressingModeTranslations db

    MsgBox "ref_salutation, ref_addressing_mode und Personalisierungs-Uebersetzungen wurden erfolgreich initialisiert.", vbInformation, MODULE_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Initialisieren der Personalisierungs-Referenzen: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Public Sub SeedContactTypeReference()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    db.Execute "DELETE FROM ref_contact_type", dbFailOnError
    InsertContactType db, "EMAIL", "CONTACT_TYPE.EMAIL", 10, True
    InsertContactType db, "PHONE", "CONTACT_TYPE.PHONE", 20, True
    InsertContactType db, "MOBILE", "CONTACT_TYPE.MOBILE", 30, True
    InsertContactType db, "WEBSITE", "CONTACT_TYPE.WEBSITE", 40, True
    InsertContactType db, "FAX", "CONTACT_TYPE.FAX", 50, True
    InsertContactType db, "OTHER", "CONTACT_TYPE.OTHER", 90, True

    InsertContactTypeTranslations db

    MsgBox "ref_contact_type und CONTACT_TYPE-Uebersetzungen wurden erfolgreich initialisiert.", vbInformation, MODULE_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Initialisieren von ref_contact_type: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub InsertTranslation( _
    ByVal db As DAO.Database, _
    ByVal LanguageCode As String, _
    ByVal translationKey As String, _
    ByVal TranslationValue As String, _
    ByVal isActive As Boolean, _
    Optional ByVal moduleCode As String = "", _
    Optional ByVal sortOrder As Long = 0)

    Dim sqlStatement As String

    If TranslationSeedExists(db, LanguageCode, translationKey) Then
        Exit Sub
    End If

    sqlStatement = "INSERT INTO fw_translation " & _
                   "(language_code, translation_key, translation_value, is_active, module_code, sort_order, created_at, created_by, updated_at, updated_by) " & _
                   "VALUES (" & _
                   SqlText(LanguageCode) & ", " & _
                   SqlText(translationKey) & ", " & _
                   SqlText(TranslationValue) & ", " & _
                   IIf(isActive, "True", "False") & ", " & _
                   SqlNullableText(moduleCode) & ", " & _
                   CStr(sortOrder) & ", " & _
                   "Now(), 'SYSTEM', Now(), 'SYSTEM')"

    db.Execute sqlStatement, dbFailOnError
End Sub

Public Sub SeedArticleTypeReference()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    EnsureArticleTypeSeed db, "PRODUCT", "Product", "REF.ARTICLE_TYPE.PRODUCT", "Standard article or physical product.", 10, True
    EnsureArticleTypeSeed db, "SERVICE", "Service", "REF.ARTICLE_TYPE.SERVICE", "Service article without stock behavior.", 20, True
    EnsureArticleTypeSeed db, "SUBSCRIPTION", "Subscription", "REF.ARTICLE_TYPE.SUBSCRIPTION", "Recurring service or subscription article.", 30, True
    EnsureArticleTypeSeed db, "FEE", "Fee", "REF.ARTICLE_TYPE.FEE", "Fee article for fixed charges.", 40, True
    EnsureArticleTypeSeed db, "DISCOUNT", "Discount", "REF.ARTICLE_TYPE.DISCOUNT", "Discount article for explicit reductions.", 50, True
    EnsureArticleTypeSeed db, "WINE", "Wine", "REF.ARTICLE_TYPE.WINE", "Wine article prepared for future wine-specific extensions.", 60, True
    EnsureArticleTypeSeed db, "CUSTOM_SIZE", "Custom Size", "REF.ARTICLE_TYPE.CUSTOM_SIZE", "Article with future dimensional specialization.", 70, True
    EnsureArticleTypeSeed db, "APPAREL_SIZE", "Apparel Size", "REF.ARTICLE_TYPE.APPAREL_SIZE", "Article with future apparel size specialization.", 80, True

    EnsureTranslationSeed db, "DE-CH", "REF.ARTICLE_TYPE.PRODUCT", "Produkt", "REF", 400
    EnsureTranslationSeed db, "EN-US", "REF.ARTICLE_TYPE.PRODUCT", "Product", "REF", 400
    EnsureTranslationSeed db, "FR-FR", "REF.ARTICLE_TYPE.PRODUCT", "Produit", "REF", 400

    EnsureTranslationSeed db, "DE-CH", "REF.ARTICLE_TYPE.SERVICE", "Dienstleistung", "REF", 401
    EnsureTranslationSeed db, "EN-US", "REF.ARTICLE_TYPE.SERVICE", "Service", "REF", 401
    EnsureTranslationSeed db, "FR-FR", "REF.ARTICLE_TYPE.SERVICE", "Service", "REF", 401

    EnsureTranslationSeed db, "DE-CH", "REF.ARTICLE_TYPE.SUBSCRIPTION", "Abonnement", "REF", 402
    EnsureTranslationSeed db, "EN-US", "REF.ARTICLE_TYPE.SUBSCRIPTION", "Subscription", "REF", 402
    EnsureTranslationSeed db, "FR-FR", "REF.ARTICLE_TYPE.SUBSCRIPTION", "Abonnement", "REF", 402

    EnsureTranslationSeed db, "DE-CH", "REF.ARTICLE_TYPE.FEE", "Gebuehr", "REF", 403
    EnsureTranslationSeed db, "EN-US", "REF.ARTICLE_TYPE.FEE", "Fee", "REF", 403
    EnsureTranslationSeed db, "FR-FR", "REF.ARTICLE_TYPE.FEE", "Frais", "REF", 403

    EnsureTranslationSeed db, "DE-CH", "REF.ARTICLE_TYPE.DISCOUNT", "Rabatt", "REF", 404
    EnsureTranslationSeed db, "EN-US", "REF.ARTICLE_TYPE.DISCOUNT", "Discount", "REF", 404
    EnsureTranslationSeed db, "FR-FR", "REF.ARTICLE_TYPE.DISCOUNT", "Remise", "REF", 404

    EnsureTranslationSeed db, "DE-CH", "REF.ARTICLE_TYPE.WINE", "Wein", "REF", 405
    EnsureTranslationSeed db, "EN-US", "REF.ARTICLE_TYPE.WINE", "Wine", "REF", 405
    EnsureTranslationSeed db, "FR-FR", "REF.ARTICLE_TYPE.WINE", "Vin", "REF", 405

    EnsureTranslationSeed db, "DE-CH", "REF.ARTICLE_TYPE.CUSTOM_SIZE", "Massanfertigung", "REF", 406
    EnsureTranslationSeed db, "EN-US", "REF.ARTICLE_TYPE.CUSTOM_SIZE", "Custom Size", "REF", 406
    EnsureTranslationSeed db, "FR-FR", "REF.ARTICLE_TYPE.CUSTOM_SIZE", "Sur mesure", "REF", 406

    EnsureTranslationSeed db, "DE-CH", "REF.ARTICLE_TYPE.APPAREL_SIZE", "Kleidergroesse", "REF", 407
    EnsureTranslationSeed db, "EN-US", "REF.ARTICLE_TYPE.APPAREL_SIZE", "Apparel Size", "REF", 407
    EnsureTranslationSeed db, "FR-FR", "REF.ARTICLE_TYPE.APPAREL_SIZE", "Taille de vetement", "REF", 407

    MsgBox "ref_article_type_code und ARTICLE_TYPE-Uebersetzungen wurden erfolgreich initialisiert.", vbInformation, MODULE_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Initialisieren von ref_article_type_code: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub InsertAddressType( _
    ByVal db As DAO.Database, _
    ByVal addressTypeCode As String, _
    ByVal translationKey As String, _
    ByVal sortOrder As Long, _
    ByVal isActive As Boolean)

    Dim sqlStatement As String

    sqlStatement = "INSERT INTO ref_address_type " & _
                   "(address_type_code, translation_key, sort_order, is_active, created_at, created_by, updated_at, updated_by) " & _
                   "VALUES (" & _
                   SqlText(addressTypeCode) & ", " & _
                   SqlText(translationKey) & ", " & _
                   CStr(sortOrder) & ", " & _
                   IIf(isActive, "True", "False") & ", " & _
                   "Now(), 'SYSTEM', Now(), 'SYSTEM')"

    db.Execute sqlStatement, dbFailOnError
End Sub

Private Sub InsertSalutation( _
    ByVal db As DAO.Database, _
    ByVal salutationCode As String, _
    ByVal translationKey As String, _
    ByVal sortOrder As Long, _
    ByVal isActive As Boolean)

    Dim sqlStatement As String

    sqlStatement = "INSERT INTO ref_salutation " & _
                   "(salutation_code, translation_key, sort_order, is_active, created_at, created_by, updated_at, updated_by) " & _
                   "VALUES (" & _
                   SqlText(salutationCode) & ", " & _
                   SqlText(translationKey) & ", " & _
                   CStr(sortOrder) & ", " & _
                   IIf(isActive, "True", "False") & ", " & _
                   "Now(), 'SYSTEM', Now(), 'SYSTEM')"

    db.Execute sqlStatement, dbFailOnError
End Sub

Private Sub InsertAddressingMode( _
    ByVal db As DAO.Database, _
    ByVal addressingModeCode As String, _
    ByVal translationKey As String, _
    ByVal sortOrder As Long, _
    ByVal isActive As Boolean)

    Dim sqlStatement As String

    sqlStatement = "INSERT INTO ref_addressing_mode " & _
                   "(addressing_mode_code, translation_key, sort_order, is_active, created_at, created_by, updated_at, updated_by) " & _
                   "VALUES (" & _
                   SqlText(addressingModeCode) & ", " & _
                   SqlText(translationKey) & ", " & _
                   CStr(sortOrder) & ", " & _
                   IIf(isActive, "True", "False") & ", " & _
                   "Now(), 'SYSTEM', Now(), 'SYSTEM')"

    db.Execute sqlStatement, dbFailOnError
End Sub

Private Sub InsertContactType( _
    ByVal db As DAO.Database, _
    ByVal contactTypeCode As String, _
    ByVal translationKey As String, _
    ByVal sortOrder As Long, _
    ByVal isActive As Boolean)

    Dim sqlStatement As String

    sqlStatement = "INSERT INTO ref_contact_type " & _
                   "(contact_type_code, translation_key, sort_order, is_active, created_at, created_by, updated_at, updated_by) " & _
                   "VALUES (" & _
                   SqlText(contactTypeCode) & ", " & _
                   SqlText(translationKey) & ", " & _
                   CStr(sortOrder) & ", " & _
                   IIf(isActive, "True", "False") & ", " & _
                   "Now(), 'SYSTEM', Now(), 'SYSTEM')"

    db.Execute sqlStatement, dbFailOnError
End Sub

Private Sub InsertUnit( _
    ByVal db As DAO.Database, _
    ByVal unitCode As String, _
    ByVal translationKey As String, _
    ByVal sortOrder As Long, _
    ByVal isActive As Boolean)

    Dim sqlStatement As String

    sqlStatement = "INSERT INTO ref_unit " & _
                   "(unit_code, translation_key, sort_order, is_active, created_at, created_by, updated_at, updated_by) " & _
                   "VALUES (" & _
                   SqlText(unitCode) & ", " & _
                   SqlText(translationKey) & ", " & _
                   CStr(sortOrder) & ", " & _
                   IIf(isActive, "True", "False") & ", " & _
                   "Now(), 'SYSTEM', Now(), 'SYSTEM')"

    db.Execute sqlStatement, dbFailOnError
End Sub

Private Sub InsertArticleType( _
    ByVal db As DAO.Database, _
    ByVal articleTypeCode As String, _
    ByVal articleTypeName As String, _
    ByVal translationKey As String, _
    ByVal descriptionText As String, _
    ByVal sortOrder As Long, _
    ByVal isActive As Boolean)

    Dim sqlStatement As String

    sqlStatement = "INSERT INTO ref_article_type_code " & _
                   "(article_type_code, article_type_name, translation_key, description_text, sort_order, is_active, created_at, created_by, updated_at, updated_by) " & _
                   "VALUES (" & _
                   SqlText(articleTypeCode) & ", " & _
                   SqlText(articleTypeName) & ", " & _
                   SqlText(translationKey) & ", " & _
                   SqlNullableText(descriptionText) & ", " & _
                   CStr(sortOrder) & ", " & _
                   IIf(isActive, "True", "False") & ", " & _
                   "Now(), 'SYSTEM', Now(), 'SYSTEM')"

    db.Execute sqlStatement, dbFailOnError
End Sub

Private Sub EnsureArticleTypeSeed( _
    ByVal db As DAO.Database, _
    ByVal articleTypeCode As String, _
    ByVal articleTypeName As String, _
    ByVal translationKey As String, _
    ByVal descriptionText As String, _
    ByVal sortOrder As Long, _
    ByVal isActive As Boolean)
    On Error GoTo ErrorHandler

    Dim criteria As String
    Dim updateSql As String

    criteria = "article_type_code = " & SqlText(articleTypeCode)

    If DCount("*", "ref_article_type_code", criteria) > 0 Then
        updateSql = "UPDATE ref_article_type_code SET " & _
                    "article_type_name = " & SqlText(articleTypeName) & ", " & _
                    "translation_key = " & SqlText(translationKey) & ", " & _
                    "description_text = " & SqlNullableText(descriptionText) & ", " & _
                    "sort_order = " & CStr(sortOrder) & ", " & _
                    "is_active = " & IIf(isActive, "True", "False") & ", " & _
                    "updated_at = Now(), " & _
                    "updated_by = 'SYSTEM' " & _
                    "WHERE " & criteria & ";"
        db.Execute updateSql, dbFailOnError
    Else
        InsertArticleType db, articleTypeCode, articleTypeName, translationKey, descriptionText, sortOrder, isActive
    End If
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & ".EnsureArticleTypeSeed", Err.description
End Sub

Private Sub InsertVatCode( _
    ByVal db As DAO.Database, _
    ByVal vatCode As String, _
    ByVal translationKey As String, _
    ByVal vatRate As Double, _
    ByVal countryCode As String, _
    ByVal validFrom As Date, _
    ByVal validTo As Variant, _
    ByVal sortOrder As Long, _
    ByVal isActive As Boolean)

    Dim sqlStatement As String

    sqlStatement = "INSERT INTO ref_vat_code " & _
                   "(vat_code, translation_key, vat_rate, country_code, valid_from, valid_to, sort_order, is_active, created_at, created_by, updated_at, updated_by) " & _
                   "VALUES (" & _
                   SqlText(vatCode) & ", " & _
                   SqlText(translationKey) & ", " & _
                   Replace(CStr(vatRate), ",", ".") & ", " & _
                   SqlText(countryCode) & ", " & _
                   "#" & Format$(validFrom, "yyyy-mm-dd") & "#, " & _
                   SqlDateOrNull(validTo) & ", " & _
                   CStr(sortOrder) & ", " & _
                   IIf(isActive, "True", "False") & ", " & _
                   "Now(), 'SYSTEM', Now(), 'SYSTEM')"

    db.Execute sqlStatement, dbFailOnError
End Sub

Private Sub InsertAddressTypeTranslations(ByVal db As DAO.Database)
    InsertTranslation db, "DE-CH", "ADDRESS_TYPE.PRIVATE", "Privat", True
    InsertTranslation db, "DE-DE", "ADDRESS_TYPE.PRIVATE", "Privat", True
    InsertTranslation db, "FR-FR", "ADDRESS_TYPE.PRIVATE", "Prive", True
    InsertTranslation db, "IT-CH", "ADDRESS_TYPE.PRIVATE", "Privato", True
    InsertTranslation db, "EN-US", "ADDRESS_TYPE.PRIVATE", "Private", True

    InsertTranslation db, "DE-CH", "ADDRESS_TYPE.COMPANY", "Firma", True
    InsertTranslation db, "DE-DE", "ADDRESS_TYPE.COMPANY", "Unternehmen", True
    InsertTranslation db, "FR-FR", "ADDRESS_TYPE.COMPANY", "Entreprise", True
    InsertTranslation db, "IT-CH", "ADDRESS_TYPE.COMPANY", "Azienda", True
    InsertTranslation db, "EN-US", "ADDRESS_TYPE.COMPANY", "Company", True

    InsertTranslation db, "DE-CH", "ADDRESS_TYPE.CUSTOMER", "Kunde", True
    InsertTranslation db, "DE-DE", "ADDRESS_TYPE.CUSTOMER", "Kunde", True
    InsertTranslation db, "FR-FR", "ADDRESS_TYPE.CUSTOMER", "Client", True
    InsertTranslation db, "IT-CH", "ADDRESS_TYPE.CUSTOMER", "Cliente", True
    InsertTranslation db, "EN-US", "ADDRESS_TYPE.CUSTOMER", "Customer", True

    InsertTranslation db, "DE-CH", "ADDRESS_TYPE.SUPPLIER", "Lieferant", True
    InsertTranslation db, "DE-DE", "ADDRESS_TYPE.SUPPLIER", "Lieferant", True
    InsertTranslation db, "FR-FR", "ADDRESS_TYPE.SUPPLIER", "Fournisseur", True
    InsertTranslation db, "IT-CH", "ADDRESS_TYPE.SUPPLIER", "Fornitore", True
    InsertTranslation db, "EN-US", "ADDRESS_TYPE.SUPPLIER", "Supplier", True

    InsertTranslation db, "DE-CH", "ADDRESS_TYPE.PARTNER", "Partner", True
    InsertTranslation db, "DE-DE", "ADDRESS_TYPE.PARTNER", "Partner", True
    InsertTranslation db, "FR-FR", "ADDRESS_TYPE.PARTNER", "Partenaire", True
    InsertTranslation db, "IT-CH", "ADDRESS_TYPE.PARTNER", "Partner", True
    InsertTranslation db, "EN-US", "ADDRESS_TYPE.PARTNER", "Partner", True

    InsertTranslation db, "DE-CH", "ADDRESS_TYPE.EMPLOYEE", "Mitarbeiter", True
    InsertTranslation db, "DE-DE", "ADDRESS_TYPE.EMPLOYEE", "Mitarbeiter", True
    InsertTranslation db, "FR-FR", "ADDRESS_TYPE.EMPLOYEE", "Employe", True
    InsertTranslation db, "IT-CH", "ADDRESS_TYPE.EMPLOYEE", "Collaboratore", True
    InsertTranslation db, "EN-US", "ADDRESS_TYPE.EMPLOYEE", "Employee", True

    InsertTranslation db, "DE-CH", "ADDRESS_TYPE.OTHER", "Andere", True
    InsertTranslation db, "DE-DE", "ADDRESS_TYPE.OTHER", "Sonstige", True
    InsertTranslation db, "FR-FR", "ADDRESS_TYPE.OTHER", "Autre", True
    InsertTranslation db, "IT-CH", "ADDRESS_TYPE.OTHER", "Altro", True
    InsertTranslation db, "EN-US", "ADDRESS_TYPE.OTHER", "Other", True
End Sub

Private Sub InsertSalutationTranslations(ByVal db As DAO.Database)
    InsertTranslation db, "DE-CH", "SALUTATION.MR", "Herr", True
    InsertTranslation db, "DE-DE", "SALUTATION.MR", "Herr", True
    InsertTranslation db, "FR-FR", "SALUTATION.MR", "Monsieur", True
    InsertTranslation db, "IT-CH", "SALUTATION.MR", "Signor", True
    InsertTranslation db, "EN-US", "SALUTATION.MR", "Mr.", True

    InsertTranslation db, "DE-CH", "SALUTATION.MS", "Frau", True
    InsertTranslation db, "DE-DE", "SALUTATION.MS", "Frau", True
    InsertTranslation db, "FR-FR", "SALUTATION.MS", "Madame", True
    InsertTranslation db, "IT-CH", "SALUTATION.MS", "Signora", True
    InsertTranslation db, "EN-US", "SALUTATION.MS", "Ms.", True

    InsertTranslation db, "DE-CH", "SALUTATION.COMPANY", "Firma", True
    InsertTranslation db, "DE-DE", "SALUTATION.COMPANY", "Unternehmen", True
    InsertTranslation db, "FR-FR", "SALUTATION.COMPANY", "Entreprise", True
    InsertTranslation db, "IT-CH", "SALUTATION.COMPANY", "Azienda", True
    InsertTranslation db, "EN-US", "SALUTATION.COMPANY", "Company", True

    InsertTranslation db, "DE-CH", "SALUTATION.NEUTRAL", "Neutral", True
    InsertTranslation db, "DE-DE", "SALUTATION.NEUTRAL", "Neutral", True
    InsertTranslation db, "FR-FR", "SALUTATION.NEUTRAL", "Neutre", True
    InsertTranslation db, "IT-CH", "SALUTATION.NEUTRAL", "Neutrale", True
    InsertTranslation db, "EN-US", "SALUTATION.NEUTRAL", "Neutral", True
End Sub

Private Sub InsertAddressingModeTranslations(ByVal db As DAO.Database)
    InsertTranslation db, "DE-CH", "ADDRESSING_MODE.FORMAL", "Formal", True
    InsertTranslation db, "DE-DE", "ADDRESSING_MODE.FORMAL", "Formal", True
    InsertTranslation db, "FR-FR", "ADDRESSING_MODE.FORMAL", "Formel", True
    InsertTranslation db, "IT-CH", "ADDRESSING_MODE.FORMAL", "Formale", True
    InsertTranslation db, "EN-US", "ADDRESSING_MODE.FORMAL", "Formal", True

    InsertTranslation db, "DE-CH", "ADDRESSING_MODE.INFORMAL", "Informell", True
    InsertTranslation db, "DE-DE", "ADDRESSING_MODE.INFORMAL", "Informell", True
    InsertTranslation db, "FR-FR", "ADDRESSING_MODE.INFORMAL", "Informel", True
    InsertTranslation db, "IT-CH", "ADDRESSING_MODE.INFORMAL", "Informale", True
    InsertTranslation db, "EN-US", "ADDRESSING_MODE.INFORMAL", "Informal", True
End Sub

Private Sub InsertContactTypeTranslations(ByVal db As DAO.Database)
    InsertTranslation db, "DE-CH", "CONTACT_TYPE.EMAIL", "E-Mail", True
    InsertTranslation db, "DE-DE", "CONTACT_TYPE.EMAIL", "E-Mail", True
    InsertTranslation db, "FR-FR", "CONTACT_TYPE.EMAIL", "E-mail", True
    InsertTranslation db, "IT-CH", "CONTACT_TYPE.EMAIL", "E-mail", True
    InsertTranslation db, "EN-US", "CONTACT_TYPE.EMAIL", "E-Mail", True

    InsertTranslation db, "DE-CH", "CONTACT_TYPE.PHONE", "Telefon", True
    InsertTranslation db, "DE-DE", "CONTACT_TYPE.PHONE", "Telefon", True
    InsertTranslation db, "FR-FR", "CONTACT_TYPE.PHONE", "Telephone", True
    InsertTranslation db, "IT-CH", "CONTACT_TYPE.PHONE", "Telefono", True
    InsertTranslation db, "EN-US", "CONTACT_TYPE.PHONE", "Phone", True

    InsertTranslation db, "DE-CH", "CONTACT_TYPE.MOBILE", "Mobil", True
    InsertTranslation db, "DE-DE", "CONTACT_TYPE.MOBILE", "Mobil", True
    InsertTranslation db, "FR-FR", "CONTACT_TYPE.MOBILE", "Mobile", True
    InsertTranslation db, "IT-CH", "CONTACT_TYPE.MOBILE", "Mobile", True
    InsertTranslation db, "EN-US", "CONTACT_TYPE.MOBILE", "Mobile", True

    InsertTranslation db, "DE-CH", "CONTACT_TYPE.WEBSITE", "Webseite", True
    InsertTranslation db, "DE-DE", "CONTACT_TYPE.WEBSITE", "Webseite", True
    InsertTranslation db, "FR-FR", "CONTACT_TYPE.WEBSITE", "Site web", True
    InsertTranslation db, "IT-CH", "CONTACT_TYPE.WEBSITE", "Sito web", True
    InsertTranslation db, "EN-US", "CONTACT_TYPE.WEBSITE", "Website", True

    InsertTranslation db, "DE-CH", "CONTACT_TYPE.FAX", "Fax", True
    InsertTranslation db, "DE-DE", "CONTACT_TYPE.FAX", "Fax", True
    InsertTranslation db, "FR-FR", "CONTACT_TYPE.FAX", "Fax", True
    InsertTranslation db, "IT-CH", "CONTACT_TYPE.FAX", "Fax", True
    InsertTranslation db, "EN-US", "CONTACT_TYPE.FAX", "Fax", True

    InsertTranslation db, "DE-CH", "CONTACT_TYPE.OTHER", "Sonstige", True
    InsertTranslation db, "DE-DE", "CONTACT_TYPE.OTHER", "Sonstige", True
    InsertTranslation db, "FR-FR", "CONTACT_TYPE.OTHER", "Autre", True
    InsertTranslation db, "IT-CH", "CONTACT_TYPE.OTHER", "Altro", True
    InsertTranslation db, "EN-US", "CONTACT_TYPE.OTHER", "Other", True
End Sub

Private Sub InsertUnitTranslations(ByVal db As DAO.Database)
    InsertTranslation db, "DE-CH", "UNIT.PCS", "Stk", True
    InsertTranslation db, "DE-DE", "UNIT.PCS", "Stk", True
    InsertTranslation db, "FR-FR", "UNIT.PCS", "pcs", True
    InsertTranslation db, "IT-CH", "UNIT.PCS", "pz", True
    InsertTranslation db, "EN-US", "UNIT.PCS", "pcs", True

    InsertTranslation db, "DE-CH", "UNIT.HOUR", "Stunde", True
    InsertTranslation db, "DE-DE", "UNIT.HOUR", "Stunde", True
    InsertTranslation db, "FR-FR", "UNIT.HOUR", "Heure", True
    InsertTranslation db, "IT-CH", "UNIT.HOUR", "Ora", True
    InsertTranslation db, "EN-US", "UNIT.HOUR", "Hour", True

    InsertTranslation db, "DE-CH", "UNIT.KG", "Kilogramm", True
    InsertTranslation db, "DE-DE", "UNIT.KG", "Kilogramm", True
    InsertTranslation db, "FR-FR", "UNIT.KG", "Kilogramme", True
    InsertTranslation db, "IT-CH", "UNIT.KG", "Chilogrammo", True
    InsertTranslation db, "EN-US", "UNIT.KG", "Kilogram", True

    InsertTranslation db, "DE-CH", "UNIT.METER", "Meter", True
    InsertTranslation db, "DE-DE", "UNIT.METER", "Meter", True
    InsertTranslation db, "FR-FR", "UNIT.METER", "Metre", True
    InsertTranslation db, "IT-CH", "UNIT.METER", "Metro", True
    InsertTranslation db, "EN-US", "UNIT.METER", "Meter", True

    InsertTranslation db, "DE-CH", "UNIT.LITER", "Liter", True
    InsertTranslation db, "DE-DE", "UNIT.LITER", "Liter", True
    InsertTranslation db, "FR-FR", "UNIT.LITER", "Litre", True
    InsertTranslation db, "IT-CH", "UNIT.LITER", "Litro", True
    InsertTranslation db, "EN-US", "UNIT.LITER", "Liter", True

    InsertTranslation db, "DE-CH", "UNIT.PACKAGE", "Paket", True
    InsertTranslation db, "DE-DE", "UNIT.PACKAGE", "Paket", True
    InsertTranslation db, "FR-FR", "UNIT.PACKAGE", "Colis", True
    InsertTranslation db, "IT-CH", "UNIT.PACKAGE", "Pacco", True
    InsertTranslation db, "EN-US", "UNIT.PACKAGE", "Package", True
End Sub

Private Sub InsertVatCodeTranslations(ByVal db As DAO.Database)
    InsertTranslation db, "DE-CH", "VAT.CH.STANDARD", "Normalsatz", True
    InsertTranslation db, "DE-DE", "VAT.CH.STANDARD", "Standardsatz", True
    InsertTranslation db, "FR-FR", "VAT.CH.STANDARD", "Taux normal", True
    InsertTranslation db, "IT-CH", "VAT.CH.STANDARD", "Aliquota normale", True
    InsertTranslation db, "EN-US", "VAT.CH.STANDARD", "Standard rate", True

    InsertTranslation db, "DE-CH", "VAT.CH.REDUCED", "Reduzierter Satz", True
    InsertTranslation db, "DE-DE", "VAT.CH.REDUCED", "Ermaessigter Satz", True
    InsertTranslation db, "FR-FR", "VAT.CH.REDUCED", "Taux reduit", True
    InsertTranslation db, "IT-CH", "VAT.CH.REDUCED", "Aliquota ridotta", True
    InsertTranslation db, "EN-US", "VAT.CH.REDUCED", "Reduced rate", True

    InsertTranslation db, "DE-CH", "VAT.CH.SPECIAL", "Sondersatz", True
    InsertTranslation db, "DE-DE", "VAT.CH.SPECIAL", "Sondersatz", True
    InsertTranslation db, "FR-FR", "VAT.CH.SPECIAL", "Taux special", True
    InsertTranslation db, "IT-CH", "VAT.CH.SPECIAL", "Aliquota speciale", True
    InsertTranslation db, "EN-US", "VAT.CH.SPECIAL", "Special rate", True

    InsertTranslation db, "DE-CH", "VAT.CH.ZERO", "Nullsatz", True
    InsertTranslation db, "DE-DE", "VAT.CH.ZERO", "Nullsatz", True
    InsertTranslation db, "FR-FR", "VAT.CH.ZERO", "Taux zero", True
    InsertTranslation db, "IT-CH", "VAT.CH.ZERO", "Aliquota zero", True
    InsertTranslation db, "EN-US", "VAT.CH.ZERO", "Zero rate", True
End Sub

Private Sub InsertShellTranslations(ByVal db As DAO.Database)
    EnsureTranslationSeed db, "DE-CH", "NAV.GROUP.ADDRESSES", "Adressen", "NAVIGATION", 10
    EnsureTranslationSeed db, "EN-US", "NAV.GROUP.ADDRESSES", "Addresses", "NAVIGATION", 10
    EnsureTranslationSeed db, "DE-CH", "NAV.ADDRESS_LIST", "Adressliste", "NAVIGATION", 20
    EnsureTranslationSeed db, "EN-US", "NAV.ADDRESS_LIST", "Address list", "NAVIGATION", 20
    EnsureTranslationSeed db, "DE-CH", "NAV.NEW_ADDRESS", "Neue Adresse", "NAVIGATION", 30
    EnsureTranslationSeed db, "EN-US", "NAV.NEW_ADDRESS", "New address", "NAVIGATION", 30
    EnsureTranslationSeed db, "DE-CH", "NAV.GROUP.DOCUMENTS", "Dokumente", "NAVIGATION", 40
    EnsureTranslationSeed db, "EN-US", "NAV.GROUP.DOCUMENTS", "Documents", "NAVIGATION", 40
    EnsureTranslationSeed db, "DE-CH", "NAV.DOCUMENT_PREVIEW", "Dokumentvorschau", "NAVIGATION", 50
    EnsureTranslationSeed db, "EN-US", "NAV.DOCUMENT_PREVIEW", "Document preview", "NAVIGATION", 50
    EnsureTranslationSeed db, "DE-CH", "NAV.GROUP.FRAMEWORK", "Framework", "NAVIGATION", 60
    EnsureTranslationSeed db, "EN-US", "NAV.GROUP.FRAMEWORK", "Framework", "NAVIGATION", 60
    EnsureTranslationSeed db, "DE-CH", "NAV.TRANSLATIONS", "Uebersetzungen", "NAVIGATION", 70
    EnsureTranslationSeed db, "EN-US", "NAV.TRANSLATIONS", "Translations", "NAVIGATION", 70
    EnsureTranslationSeed db, "DE-CH", "NAV.FW_TRANSLATIONS", "Uebersetzungen pflegen", "NAVIGATION", 80
    EnsureTranslationSeed db, "EN-US", "NAV.FW_TRANSLATIONS", "Maintain translations", "NAVIGATION", 80
    EnsureTranslationSeed db, "DE-CH", "NAV.LOCALISATION", "Lokalisierung", "NAVIGATION", 90
    EnsureTranslationSeed db, "EN-US", "NAV.LOCALISATION", "Localisation", "NAVIGATION", 90
    EnsureTranslationSeed db, "DE-CH", "NAV.TAGS", "Tags", "NAVIGATION", 100
    EnsureTranslationSeed db, "EN-US", "NAV.TAGS", "Tags", "NAVIGATION", 100
    EnsureTranslationSeed db, "DE-CH", "NAV.TAG_HELP", "Tag-Hilfe", "NAVIGATION", 110
    EnsureTranslationSeed db, "EN-US", "NAV.TAG_HELP", "Tag help", "NAVIGATION", 110
    EnsureTranslationSeed db, "DE-CH", "NAV.GROUP.ORDERS", "Bestellungen", "NAVIGATION", 120
    EnsureTranslationSeed db, "EN-US", "NAV.GROUP.ORDERS", "Orders", "NAVIGATION", 120
    EnsureTranslationSeed db, "DE-CH", "NAV.GROUP.FINANCE", "Finanzen", "NAVIGATION", 130
    EnsureTranslationSeed db, "EN-US", "NAV.GROUP.FINANCE", "Finance", "NAVIGATION", 130
    EnsureTranslationSeed db, "DE-CH", "NAV.GROUP.REPORTING", "Auswertungen", "NAVIGATION", 140
    EnsureTranslationSeed db, "EN-US", "NAV.GROUP.REPORTING", "Reports", "NAVIGATION", 140
    EnsureTranslationSeed db, "DE-CH", "NAV.GROUP.TENANT", "Mandant", "NAVIGATION", 150
    EnsureTranslationSeed db, "EN-US", "NAV.GROUP.TENANT", "Tenant", "NAVIGATION", 150
    EnsureTranslationSeed db, "DE-CH", "NAV.ARTICLE_GROUPS", "Artikelgruppen", "NAVIGATION", 155
    EnsureTranslationSeed db, "EN-US", "NAV.ARTICLE_GROUPS", "Article groups", "NAVIGATION", 155
    EnsureTranslationSeed db, "DE-CH", "NAV.NEW_ARTICLE_GROUP", "Neue Artikelgruppe", "NAVIGATION", 156
    EnsureTranslationSeed db, "EN-US", "NAV.NEW_ARTICLE_GROUP", "New article group", "NAVIGATION", 156
    EnsureTranslationSeed db, "DE-CH", "NAV.ARTICLES", "Artikel", "NAVIGATION", 157
    EnsureTranslationSeed db, "EN-US", "NAV.ARTICLES", "Articles", "NAVIGATION", 157
    EnsureTranslationSeed db, "DE-CH", "NAV.GROUP.SYSTEM", "System", "NAVIGATION", 160
    EnsureTranslationSeed db, "EN-US", "NAV.GROUP.SYSTEM", "System", "NAVIGATION", 160
    EnsureTranslationSeed db, "DE-CH", "NAV.FW_NAVIGATION_ADMIN", "Navigation verwalten", "NAVIGATION", 170
    EnsureTranslationSeed db, "EN-US", "NAV.FW_NAVIGATION_ADMIN", "Manage navigation", "NAVIGATION", 170

    EnsureTranslationSeed db, "DE-CH", "COMMON.NEW", "Neu", "COMMON", 171
    EnsureTranslationSeed db, "EN-US", "COMMON.NEW", "New", "COMMON", 171
    EnsureTranslationSeed db, "DE-CH", "COMMON.EDIT", "Bearbeiten", "COMMON", 172
    EnsureTranslationSeed db, "EN-US", "COMMON.EDIT", "Edit", "COMMON", 172
    EnsureTranslationSeed db, "DE-CH", "COMMON.REFRESH", "Aktualisieren", "COMMON", 173
    EnsureTranslationSeed db, "EN-US", "COMMON.REFRESH", "Refresh", "COMMON", 173
    EnsureTranslationSeed db, "DE-CH", "COMMON.SEARCH", "Suche", "COMMON", 174
    EnsureTranslationSeed db, "EN-US", "COMMON.SEARCH", "Search", "COMMON", 174
    EnsureTranslationSeed db, "DE-CH", "COMMON.CLEAR_SEARCH", "Leeren", "COMMON", 175
    EnsureTranslationSeed db, "EN-US", "COMMON.CLEAR_SEARCH", "Clear", "COMMON", 175
    EnsureTranslationSeed db, "DE-CH", "COMMON.HOME", "Home", "COMMON", 176
    EnsureTranslationSeed db, "EN-US", "COMMON.HOME", "Home", "COMMON", 176
    EnsureTranslationSeed db, "DE-CH", "COMMON.BACK", "Zurueck", "COMMON", 177
    EnsureTranslationSeed db, "EN-US", "COMMON.BACK", "Back", "COMMON", 177

    EnsureTranslationSeed db, "DE-CH", "FORM.FRMAPPSHELL.APP_TITLE", "EASIS v4", "FORM", 10
    EnsureTranslationSeed db, "EN-US", "FORM.FRMAPPSHELL.APP_TITLE", "EASIS v4", "FORM", 10
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMAPPSHELL.APP_SUBTITLE", "Access Framework", "FORM", 20
    EnsureTranslationSeed db, "EN-US", "FORM.FRMAPPSHELL.APP_SUBTITLE", "Access Framework", "FORM", 20
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMAPPSHELL.USER", "Benutzer", "FORM", 30
    EnsureTranslationSeed db, "EN-US", "FORM.FRMAPPSHELL.USER", "User", "FORM", 30
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMAPPSHELL.TENANT", "Mandant", "FORM", 40
    EnsureTranslationSeed db, "EN-US", "FORM.FRMAPPSHELL.TENANT", "Tenant", "FORM", 40
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMAPPSHELL.ROLE", "Rolle", "FORM", 50
    EnsureTranslationSeed db, "EN-US", "FORM.FRMAPPSHELL.ROLE", "Role", "FORM", 50
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMAPPSHELL.ENVIRONMENT", "Umgebung", "FORM", 60
    EnsureTranslationSeed db, "EN-US", "FORM.FRMAPPSHELL.ENVIRONMENT", "Environment", "FORM", 60
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMAPPSHELL.BACKEND", "Backend", "FORM", 70
    EnsureTranslationSeed db, "EN-US", "FORM.FRMAPPSHELL.BACKEND", "Backend", "FORM", 70

    EnsureTranslationSeed db, "DE-CH", "FORM.FRMAPPDASHBOARD.TENANT", "Mandant", "FORM", 80
    EnsureTranslationSeed db, "EN-US", "FORM.FRMAPPDASHBOARD.TENANT", "Tenant", "FORM", 80
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMAPPDASHBOARD.USER", "Benutzer", "FORM", 90
    EnsureTranslationSeed db, "EN-US", "FORM.FRMAPPDASHBOARD.USER", "User", "FORM", 90
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMAPPDASHBOARD.BACKEND", "Backend", "FORM", 100
    EnsureTranslationSeed db, "EN-US", "FORM.FRMAPPDASHBOARD.BACKEND", "Backend", "FORM", 100
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMAPPDASHBOARD.FRAMEWORK", "Framework", "FORM", 110
    EnsureTranslationSeed db, "EN-US", "FORM.FRMAPPDASHBOARD.FRAMEWORK", "Framework", "FORM", 110
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMAPPDASHBOARD.STATUS", "Status", "FORM", 120
    EnsureTranslationSeed db, "EN-US", "FORM.FRMAPPDASHBOARD.STATUS", "Status", "FORM", 120

    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWNAVIGATIONADMIN.FORM_TITLE", "Navigation verwalten", "FORM", 130
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWNAVIGATIONADMIN.FORM_TITLE", "Manage navigation", "FORM", 130
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWNAVIGATIONADMIN.SAVE", "Speichern", "FORM", 140
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWNAVIGATIONADMIN.SAVE", "Save", "FORM", 140
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWNAVIGATIONADMIN.REFRESH", "Aktualisieren", "FORM", 150
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWNAVIGATIONADMIN.REFRESH", "Refresh", "FORM", 150
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWNAVIGATIONADMIN.NEW_GROUP", "Neue Gruppe", "FORM", 160
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWNAVIGATIONADMIN.NEW_GROUP", "New group", "FORM", 160
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWNAVIGATIONADMIN.NEW_ITEM", "Neuer Eintrag", "FORM", 170
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWNAVIGATIONADMIN.NEW_ITEM", "New item", "FORM", 170
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWNAVIGATIONADMIN.DEACTIVATE", "Deaktivieren", "FORM", 180
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWNAVIGATIONADMIN.DEACTIVATE", "Deactivate", "FORM", 180
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWNAVIGATIONADMIN.HIDE", "Ausblenden", "FORM", 190
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWNAVIGATIONADMIN.HIDE", "Hide", "FORM", 190
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWNAVIGATIONADMIN.SHOW", "Einblenden", "FORM", 200
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWNAVIGATIONADMIN.SHOW", "Show", "FORM", 200

    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEGROUPLIST.FORM_TITLE", "Artikelgruppen", "FORM", 205
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEGROUPLIST.FORM_TITLE", "Article groups", "FORM", 205
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEGROUPLIST.SEARCH", "Suche", "FORM", 206
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEGROUPLIST.SEARCH", "Search", "FORM", 206
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEGROUPLIST.EDIT", "Bearbeiten", "FORM", 207
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEGROUPLIST.EDIT", "Edit", "FORM", 207
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEGROUPLIST.REFRESH", "Aktualisieren", "FORM", 208
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEGROUPLIST.REFRESH", "Refresh", "FORM", 208

    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLELIST.FORM_TITLE", "Artikel", "FORM", 2081
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLELIST.FORM_TITLE", "Articles", "FORM", 2081
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLELIST.SEARCH", "Suche", "FORM", 2082
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLELIST.SEARCH", "Search", "FORM", 2082
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLELIST.REFRESH", "Aktualisieren", "FORM", 2083
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLELIST.REFRESH", "Refresh", "FORM", 2083
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLELIST.ARTICLE_NO", "Artikel-Nr.", "FORM", 2084
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLELIST.ARTICLE_NO", "Article no.", "FORM", 2084
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLELIST.ARTICLE_NAME", "Artikelname", "FORM", 2085
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLELIST.ARTICLE_NAME", "Article name", "FORM", 2085
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLELIST.PRODUCT_GROUP", "Artikelgruppe", "FORM", 2086
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLELIST.PRODUCT_GROUP", "Product group", "FORM", 2086
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLELIST.UNIT_CODE", "Einheit", "FORM", 2087
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLELIST.UNIT_CODE", "Unit", "FORM", 2087
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLELIST.VAT_CODE", "MWST-Code", "FORM", 2088
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLELIST.VAT_CODE", "VAT code", "FORM", 2088
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLELIST.SALES_PRICE", "Verkaufspreis", "FORM", 2089
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLELIST.SALES_PRICE", "Sales price", "FORM", 2089
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLELIST.IS_ACTIVE", "Aktiv", "FORM", 2090
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLELIST.IS_ACTIVE", "Active", "FORM", 2090

    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEDETAIL.FORM_TITLE", "Artikel", "FORM", 2091
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEDETAIL.FORM_TITLE", "Article", "FORM", 2091
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEDETAIL.ARTICLE_NO", "Artikel-Nr.", "FORM", 2092
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEDETAIL.ARTICLE_NO", "Article no.", "FORM", 2092
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEDETAIL.ARTICLE_NAME", "Artikelname", "FORM", 2093
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEDETAIL.ARTICLE_NAME", "Article name", "FORM", 2093
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEDETAIL.PRODUCT_GROUP", "Artikelgruppe", "FORM", 2094
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEDETAIL.PRODUCT_GROUP", "Product group", "FORM", 2094
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEDETAIL.ARTICLE_TYPE_CODE", "Artikeltyp", "FORM", 2095
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEDETAIL.ARTICLE_TYPE_CODE", "Article type", "FORM", 2095
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEDETAIL.UNIT_CODE", "Einheit", "FORM", 2096
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEDETAIL.UNIT_CODE", "Unit", "FORM", 2096
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEDETAIL.VAT_CODE", "MWST-Code", "FORM", 2097
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEDETAIL.VAT_CODE", "VAT code", "FORM", 2097
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEDETAIL.PURCHASE_PRICE", "Einkaufspreis", "FORM", 2098
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEDETAIL.PURCHASE_PRICE", "Purchase price", "FORM", 2098
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEDETAIL.SALES_PRICE", "Verkaufspreis", "FORM", 2099
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEDETAIL.SALES_PRICE", "Sales price", "FORM", 2099
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEDETAIL.BARCODE", "Barcode", "FORM", 2100
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEDETAIL.BARCODE", "Barcode", "FORM", 2100
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEDETAIL.DESCRIPTION_TEXT", "Beschreibung", "FORM", 2101
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEDETAIL.DESCRIPTION_TEXT", "Description", "FORM", 2101
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEDETAIL.IS_ACTIVE", "Aktiv", "FORM", 2102
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEDETAIL.IS_ACTIVE", "Active", "FORM", 2102
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEDETAIL.SAVE", "Speichern", "FORM", 2103
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEDETAIL.SAVE", "Save", "FORM", 2103
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEDETAIL.CANCEL", "Abbrechen", "FORM", 2104
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEDETAIL.CANCEL", "Cancel", "FORM", 2104

    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEGROUPDETAIL.FORM_TITLE", "Artikelgruppe", "FORM", 209
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEGROUPDETAIL.FORM_TITLE", "Article group", "FORM", 209
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEGROUPDETAIL.PRODUCT_GROUP_CODE", "Artikelgruppen-Code", "FORM", 210
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEGROUPDETAIL.PRODUCT_GROUP_CODE", "Article group code", "FORM", 210
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEGROUPDETAIL.PRODUCT_GROUP_NAME", "Artikelgruppen-Name", "FORM", 211
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEGROUPDETAIL.PRODUCT_GROUP_NAME", "Article group name", "FORM", 211
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEGROUPDETAIL.DESCRIPTION_TEXT", "Beschreibung", "FORM", 212
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEGROUPDETAIL.DESCRIPTION_TEXT", "Description", "FORM", 212
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEGROUPDETAIL.ARTICLE_GROUP_CODE", "Artikelgruppen-Code", "FORM", 210
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEGROUPDETAIL.ARTICLE_GROUP_CODE", "Article group code", "FORM", 210
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEGROUPDETAIL.ARTICLE_GROUP_NAME", "Artikelgruppen-Name", "FORM", 211
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEGROUPDETAIL.ARTICLE_GROUP_NAME", "Article group name", "FORM", 211
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEGROUPDETAIL.DESCRIPTION", "Beschreibung", "FORM", 212
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEGROUPDETAIL.DESCRIPTION", "Description", "FORM", 212
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEGROUPDETAIL.IS_ACTIVE", "Aktiv", "FORM", 213
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEGROUPDETAIL.IS_ACTIVE", "Active", "FORM", 213
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEGROUPDETAIL.SORT_ORDER", "Sortierung", "FORM", 214
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEGROUPDETAIL.SORT_ORDER", "Sort order", "FORM", 214
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEGROUPDETAIL.SAVE", "Speichern", "FORM", 215
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEGROUPDETAIL.SAVE", "Save", "FORM", 215
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMARTICLEGROUPDETAIL.CANCEL", "Abbrechen", "FORM", 216
    EnsureTranslationSeed db, "EN-US", "FORM.FRMARTICLEGROUPDETAIL.CANCEL", "Cancel", "FORM", 216

    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_GROUP_CODE_REQUIRED", "Artikelgruppen-Code ist erforderlich.", "MSG", 217
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_GROUP_CODE_REQUIRED", "Article group code is required.", "MSG", 217
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_GROUP_NAME_REQUIRED", "Artikelgruppen-Name ist erforderlich.", "MSG", 218
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_GROUP_NAME_REQUIRED", "Article group name is required.", "MSG", 218
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_GROUP_DUPLICATE_CODE", "Artikelgruppen-Code existiert bereits.", "MSG", 219
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_GROUP_DUPLICATE_CODE", "Article group code already exists.", "MSG", 219
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_GROUP_SELECT_FIRST", "Bitte zuerst eine Artikelgruppe auswaehlen.", "MSG", 221
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_GROUP_SELECT_FIRST", "Please select an article group first.", "MSG", 221
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_GROUP_DISCARD_CHANGES", "Ungespeicherte Aenderungen verwerfen?", "MSG", 222
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_GROUP_DISCARD_CHANGES", "Discard unsaved changes?", "MSG", 222
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_GROUP_CANCEL_CONFIRM", "Aenderungen verwerfen?", "MSG", 223
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_GROUP_CANCEL_CONFIRM", "Discard changes?", "MSG", 223
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_GROUP_SAVE_OR_CANCEL_FIRST", "Bitte zuerst speichern oder abbrechen.", "MSG", 224
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_GROUP_SAVE_OR_CANCEL_FIRST", "Please save or cancel first.", "MSG", 224
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_GROUP_LIST_LOAD_ERROR", "Fehler beim Laden der Artikelgruppenliste.", "MSG", 225
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_GROUP_LIST_LOAD_ERROR", "Error loading the article group list.", "MSG", 225
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_GROUP_DETAIL_LOAD_ERROR", "Fehler beim Laden der Artikelgruppendetails.", "MSG", 226
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_GROUP_DETAIL_LOAD_ERROR", "Error loading the article group details.", "MSG", 226
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_GROUP_SAVE_ERROR", "Fehler beim Speichern der Artikelgruppe.", "MSG", 227
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_GROUP_SAVE_ERROR", "Error saving the article group.", "MSG", 227
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_LIST_LOAD_ERROR", "Fehler beim Laden der Artikelliste.", "MSG", 228
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_LIST_LOAD_ERROR", "Error loading the article list.", "MSG", 228
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_NO_REQUIRED", "Artikel-Nr. ist erforderlich.", "MSG", 229
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_NO_REQUIRED", "Article no. is required.", "MSG", 229
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_NAME_REQUIRED", "Artikelname ist erforderlich.", "MSG", 230
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_NAME_REQUIRED", "Article name is required.", "MSG", 230
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_GROUP_REQUIRED", "Artikelgruppe ist erforderlich.", "MSG", 231
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_GROUP_REQUIRED", "Product group is required.", "MSG", 231
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_UNIT_REQUIRED", "Einheit ist erforderlich.", "MSG", 232
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_UNIT_REQUIRED", "Unit is required.", "MSG", 232
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_VAT_REQUIRED", "MWST-Code ist erforderlich.", "MSG", 233
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_VAT_REQUIRED", "VAT code is required.", "MSG", 233
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_SALES_PRICE_REQUIRED", "Verkaufspreis ist erforderlich.", "MSG", 234
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_SALES_PRICE_REQUIRED", "Sales price is required.", "MSG", 234
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_DUPLICATE_NO", "Artikel-Nr. existiert bereits.", "MSG", 235
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_DUPLICATE_NO", "Article no. already exists.", "MSG", 235
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_SAVE_ERROR", "Fehler beim Speichern des Artikels.", "MSG", 236
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_SAVE_ERROR", "Error saving the article.", "MSG", 236
    EnsureTranslationSeed db, "DE-CH", "MSG.ARTICLE_CANCEL_CONFIRM", "Aenderungen verwerfen?", "MSG", 237
    EnsureTranslationSeed db, "EN-US", "MSG.ARTICLE_CANCEL_CONFIRM", "Discard changes?", "MSG", 237

    EnsureTranslationSeed db, "DE-CH", "STATUS.READY", "Bereit", "STATUS", 238
    EnsureTranslationSeed db, "EN-US", "STATUS.READY", "Ready", "STATUS", 238
    EnsureTranslationSeed db, "FR-FR", "STATUS.READY", "Pret", "STATUS", 238

    EnsureTranslationSeed db, "DE-CH", "COMMON.ALL", "Alle", "COMMON", 239
    EnsureTranslationSeed db, "EN-US", "COMMON.ALL", "All", "COMMON", 239
    EnsureTranslationSeed db, "FR-FR", "COMMON.ALL", "Tous", "COMMON", 239

    EnsureTranslationSeed db, "DE-CH", "NAV.FW_TRANSLATION_AUDIT", "Uebersetzungs-Audit", "NAVIGATION", 240
    EnsureTranslationSeed db, "EN-US", "NAV.FW_TRANSLATION_AUDIT", "Translation audit", "NAVIGATION", 240
    EnsureTranslationSeed db, "FR-FR", "NAV.FW_TRANSLATION_AUDIT", "Audit des traductions", "NAVIGATION", 240
    EnsureTranslationSeed db, "DE-CH", "NAV.FW_TRANSLATION_TAG_GENERATOR", "Translation-Tags", "NAVIGATION", 240
    EnsureTranslationSeed db, "EN-US", "NAV.FW_TRANSLATION_TAG_GENERATOR", "Translation tags", "NAVIGATION", 240
    EnsureTranslationSeed db, "FR-FR", "NAV.FW_TRANSLATION_TAG_GENERATOR", "Balises de traduction", "NAVIGATION", 240

    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONAUDIT.FORM_TITLE", "Uebersetzungs-Audit", "FORM", 241
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONAUDIT.FORM_TITLE", "Translation audit", "FORM", 241
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONAUDIT.FORM_TITLE", "Audit des traductions", "FORM", 241
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONAUDIT.SCOPE_CODE", "Bereich", "FORM", 242
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONAUDIT.SCOPE_CODE", "Scope", "FORM", 242
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONAUDIT.SCOPE_CODE", "Portee", "FORM", 242
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONAUDIT.LANGUAGE_CODE", "Sprache", "FORM", 243
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONAUDIT.LANGUAGE_CODE", "Language", "FORM", 243
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONAUDIT.LANGUAGE_CODE", "Langue", "FORM", 243
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONAUDIT.AUDIT_STATUS", "Audit-Status", "FORM", 244
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONAUDIT.AUDIT_STATUS", "Audit status", "FORM", 244
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONAUDIT.AUDIT_STATUS", "Statut d'audit", "FORM", 244
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONAUDIT.SEARCH", "Suche", "FORM", 245
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONAUDIT.SEARCH", "Search", "FORM", 245
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONAUDIT.SEARCH", "Recherche", "FORM", 245
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONAUDIT.COVERAGE_SUMMARY", "Abdeckungsuebersicht", "FORM", 246
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONAUDIT.COVERAGE_SUMMARY", "Coverage summary", "FORM", 246
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONAUDIT.COVERAGE_SUMMARY", "Resume de couverture", "FORM", 246
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONAUDIT.REFRESH_AUDIT", "Audit pruefen", "FORM", 247
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONAUDIT.REFRESH_AUDIT", "Run audit", "FORM", 247
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONAUDIT.REFRESH_AUDIT", "Verifier l'audit", "FORM", 247
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONAUDIT.CREATE_MISSING_ROWS", "Fehlende Eintraege erzeugen", "FORM", 248
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONAUDIT.CREATE_MISSING_ROWS", "Create missing rows", "FORM", 248
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONAUDIT.CREATE_MISSING_ROWS", "Creer les lignes manquantes", "FORM", 248
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONAUDIT.OPEN_TRANSLATION", "Uebersetzung oeffnen", "FORM", 249
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONAUDIT.OPEN_TRANSLATION", "Open translation", "FORM", 249
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONAUDIT.OPEN_TRANSLATION", "Ouvrir la traduction", "FORM", 249
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONAUDIT.CLEAR_FILTERS", "Filter loeschen", "FORM", 250
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONAUDIT.CLEAR_FILTERS", "Clear filters", "FORM", 250
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONAUDIT.CLEAR_FILTERS", "Effacer les filtres", "FORM", 250

    EnsureTranslationSeed db, "DE-CH", "MSG.FW_TRANSLATION_AUDIT_LOAD_ERROR", "Fehler beim Laden des Uebersetzungs-Audits.", "MSG", 251
    EnsureTranslationSeed db, "EN-US", "MSG.FW_TRANSLATION_AUDIT_LOAD_ERROR", "Error loading the translation audit.", "MSG", 251
    EnsureTranslationSeed db, "FR-FR", "MSG.FW_TRANSLATION_AUDIT_LOAD_ERROR", "Erreur lors du chargement de l'audit des traductions.", "MSG", 251
    EnsureTranslationSeed db, "DE-CH", "MSG.FW_TRANSLATION_AUDIT_REFRESH_ERROR", "Fehler beim Aktualisieren des Uebersetzungs-Audits.", "MSG", 252
    EnsureTranslationSeed db, "EN-US", "MSG.FW_TRANSLATION_AUDIT_REFRESH_ERROR", "Error refreshing the translation audit.", "MSG", 252
    EnsureTranslationSeed db, "FR-FR", "MSG.FW_TRANSLATION_AUDIT_REFRESH_ERROR", "Erreur lors de l'actualisation de l'audit des traductions.", "MSG", 252
    EnsureTranslationSeed db, "DE-CH", "MSG.FW_TRANSLATION_AUDIT_CREATE_MISSING_ERROR", "Fehler beim Erzeugen fehlender Uebersetzungseintraege.", "MSG", 253
    EnsureTranslationSeed db, "EN-US", "MSG.FW_TRANSLATION_AUDIT_CREATE_MISSING_ERROR", "Error creating missing translation rows.", "MSG", 253
    EnsureTranslationSeed db, "FR-FR", "MSG.FW_TRANSLATION_AUDIT_CREATE_MISSING_ERROR", "Erreur lors de la creation des lignes de traduction manquantes.", "MSG", 253
    EnsureTranslationSeed db, "DE-CH", "MSG.FW_TRANSLATION_AUDIT_OPEN_ERROR", "Fehler beim Oeffnen der Uebersetzung.", "MSG", 254
    EnsureTranslationSeed db, "EN-US", "MSG.FW_TRANSLATION_AUDIT_OPEN_ERROR", "Error opening the translation.", "MSG", 254
    EnsureTranslationSeed db, "FR-FR", "MSG.FW_TRANSLATION_AUDIT_OPEN_ERROR", "Erreur lors de l'ouverture de la traduction.", "MSG", 254
    EnsureTranslationSeed db, "DE-CH", "MSG.FW_TRANSLATION_AUDIT_SELECT_FIRST", "Bitte zuerst einen Uebersetzungseintrag auswaehlen.", "MSG", 255
    EnsureTranslationSeed db, "EN-US", "MSG.FW_TRANSLATION_AUDIT_SELECT_FIRST", "Please select a translation entry first.", "MSG", 255
    EnsureTranslationSeed db, "FR-FR", "MSG.FW_TRANSLATION_AUDIT_SELECT_FIRST", "Veuillez d'abord selectionner une entree de traduction.", "MSG", 255

    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONEDIT.FORM_TITLE", "Uebersetzung bearbeiten", "FORM", 256
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONEDIT.FORM_TITLE", "Edit translation", "FORM", 256
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONEDIT.FORM_TITLE", "Modifier la traduction", "FORM", 256
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONEDIT.TRANSLATION_KEY", "Uebersetzungsschluessel", "FORM", 257
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONEDIT.TRANSLATION_KEY", "Translation key", "FORM", 257
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONEDIT.TRANSLATION_KEY", "Cle de traduction", "FORM", 257
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONEDIT.SCOPE_CODE", "Bereich", "FORM", 258
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONEDIT.SCOPE_CODE", "Scope", "FORM", 258
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONEDIT.SCOPE_CODE", "Portee", "FORM", 258
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONEDIT.AUDIT_STATUS", "Audit-Status", "FORM", 259
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONEDIT.AUDIT_STATUS", "Audit status", "FORM", 259
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONEDIT.AUDIT_STATUS", "Statut d'audit", "FORM", 259
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_TYPE", "Quelltyp", "FORM", 260
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_TYPE", "Source type", "FORM", 260
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_TYPE", "Type de source", "FORM", 260
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_OBJECT", "Quellobjekt", "FORM", 261
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_OBJECT", "Source object", "FORM", 261
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_OBJECT", "Objet source", "FORM", 261
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_CONTROL", "Quellsteuerelement", "FORM", 262
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_CONTROL", "Source control", "FORM", 262
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_CONTROL", "Controle source", "FORM", 262
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONEDIT.FALLBACK_TEXT", "Fallback-Text", "FORM", 263
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONEDIT.FALLBACK_TEXT", "Fallback text", "FORM", 263
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONEDIT.FALLBACK_TEXT", "Texte de secours", "FORM", 263
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONEDIT.DE_CH", "Deutsch (Schweiz)", "FORM", 264
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONEDIT.DE_CH", "German (Switzerland)", "FORM", 264
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONEDIT.DE_CH", "Allemand (Suisse)", "FORM", 264
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONEDIT.EN_US", "Englisch (USA)", "FORM", 265
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONEDIT.EN_US", "English (US)", "FORM", 265
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONEDIT.EN_US", "Anglais (Etats-Unis)", "FORM", 265
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONEDIT.FR_FR", "Franzoesisch (Frankreich)", "FORM", 266
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONEDIT.FR_FR", "French (France)", "FORM", 266
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONEDIT.FR_FR", "Francais (France)", "FORM", 266
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONEDIT.SAVE", "Speichern", "FORM", 267
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONEDIT.SAVE", "Save", "FORM", 267
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONEDIT.SAVE", "Enregistrer", "FORM", 267
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONEDIT.CANCEL", "Abbrechen", "FORM", 268
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONEDIT.CANCEL", "Cancel", "FORM", 268
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONEDIT.CANCEL", "Annuler", "FORM", 268
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONEDIT.DEEPL_SUGGESTION", "DeepL Vorschlag", "FORM", 269
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONEDIT.DEEPL_SUGGESTION", "DeepL suggestion", "FORM", 269
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONEDIT.DEEPL_SUGGESTION", "Suggestion DeepL", "FORM", 269

    EnsureTranslationSeed db, "DE-CH", "MSG.TRANSLATION_KEY_REQUIRED", "Uebersetzungsschluessel ist erforderlich.", "MSG", 270
    EnsureTranslationSeed db, "EN-US", "MSG.TRANSLATION_KEY_REQUIRED", "Translation key is required.", "MSG", 270
    EnsureTranslationSeed db, "FR-FR", "MSG.TRANSLATION_KEY_REQUIRED", "La cle de traduction est obligatoire.", "MSG", 270
    EnsureTranslationSeed db, "DE-CH", "MSG.TRANSLATION_EDIT_LOAD_ERROR", "Fehler beim Laden der Uebersetzung.", "MSG", 271
    EnsureTranslationSeed db, "EN-US", "MSG.TRANSLATION_EDIT_LOAD_ERROR", "Error loading the translation.", "MSG", 271
    EnsureTranslationSeed db, "FR-FR", "MSG.TRANSLATION_EDIT_LOAD_ERROR", "Erreur lors du chargement de la traduction.", "MSG", 271
    EnsureTranslationSeed db, "DE-CH", "MSG.TRANSLATION_EDIT_SAVE_ERROR", "Fehler beim Speichern der Uebersetzung.", "MSG", 272
    EnsureTranslationSeed db, "EN-US", "MSG.TRANSLATION_EDIT_SAVE_ERROR", "Error saving the translation.", "MSG", 272
    EnsureTranslationSeed db, "FR-FR", "MSG.TRANSLATION_EDIT_SAVE_ERROR", "Erreur lors de l'enregistrement de la traduction.", "MSG", 272
    EnsureTranslationSeed db, "DE-CH", "MSG.TRANSLATION_EDIT_CANCEL_CONFIRM", "Ungespeicherte Aenderungen verwerfen?", "MSG", 273
    EnsureTranslationSeed db, "EN-US", "MSG.TRANSLATION_EDIT_CANCEL_CONFIRM", "Discard unsaved changes?", "MSG", 273
    EnsureTranslationSeed db, "FR-FR", "MSG.TRANSLATION_EDIT_CANCEL_CONFIRM", "Ignorer les modifications non enregistrees ?", "MSG", 273
    EnsureTranslationSeed db, "DE-CH", "MSG.DEEPL_API_KEY_MISSING", "DeepL API-Schluessel ist nicht konfiguriert.", "MSG", 274
    EnsureTranslationSeed db, "EN-US", "MSG.DEEPL_API_KEY_MISSING", "DeepL API key is not configured.", "MSG", 274
    EnsureTranslationSeed db, "FR-FR", "MSG.DEEPL_API_KEY_MISSING", "La cle API DeepL n'est pas configuree.", "MSG", 274
    EnsureTranslationSeed db, "DE-CH", "MSG.TRANSLATION_EDIT_DEEPL_ERROR", "Fehler beim Abrufen der DeepL-Vorschlaege.", "MSG", 275
    EnsureTranslationSeed db, "EN-US", "MSG.TRANSLATION_EDIT_DEEPL_ERROR", "Error retrieving DeepL suggestions.", "MSG", 275
    EnsureTranslationSeed db, "FR-FR", "MSG.TRANSLATION_EDIT_DEEPL_ERROR", "Erreur lors de la recuperation des suggestions DeepL.", "MSG", 275
    EnsureTranslationSeed db, "DE-CH", "MSG.TRANSLATION_EDIT_DEEPL_SOURCE_REQUIRED", "Bitte zuerst einen DE-CH Ausgangstext erfassen.", "MSG", 276
    EnsureTranslationSeed db, "EN-US", "MSG.TRANSLATION_EDIT_DEEPL_SOURCE_REQUIRED", "Please enter a DE-CH source text first.", "MSG", 276
    EnsureTranslationSeed db, "FR-FR", "MSG.TRANSLATION_EDIT_DEEPL_SOURCE_REQUIRED", "Veuillez d'abord saisir un texte source DE-CH.", "MSG", 276
    EnsureTranslationSeed db, "DE-CH", "MSG.TRANSLATION_EDIT_DEEPL_OVERWRITE_CONFIRM", "Bestehende Zieltexte mit DeepL-Vorschlaegen ersetzen?", "MSG", 277
    EnsureTranslationSeed db, "EN-US", "MSG.TRANSLATION_EDIT_DEEPL_OVERWRITE_CONFIRM", "Replace existing target texts with DeepL suggestions?", "MSG", 277
    EnsureTranslationSeed db, "FR-FR", "MSG.TRANSLATION_EDIT_DEEPL_OVERWRITE_CONFIRM", "Remplacer les textes cibles existants par des suggestions DeepL ?", "MSG", 277

    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.FORM_TITLE", "Translation-Tag-Generator", "FORM", 278
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.FORM_TITLE", "Translation tag generator", "FORM", 278
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.FORM_TITLE", "Generateur de balises de traduction", "FORM", 278
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.FORM_NAME", "Formular", "FORM", 279
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.FORM_NAME", "Form", "FORM", 279
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.FORM_NAME", "Formulaire", "FORM", 279
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.INCLUDE_HIDDEN", "Versteckte Controls einschliessen", "FORM", 280
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.INCLUDE_HIDDEN", "Include hidden controls", "FORM", 280
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.INCLUDE_HIDDEN", "Inclure les controles masques", "FORM", 280
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SHOW_ALL_CONTROLS", "Alle Controls anzeigen", "FORM", 281
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SHOW_ALL_CONTROLS", "Show all controls", "FORM", 281
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SHOW_ALL_CONTROLS", "Afficher tous les controles", "FORM", 281
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.LOAD_FORM", "Formular laden", "FORM", 282
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.LOAD_FORM", "Load form", "FORM", 282
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.LOAD_FORM", "Charger le formulaire", "FORM", 282
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.GENERATE_SUGGESTIONS", "Vorschlaege generieren", "FORM", 283
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.GENERATE_SUGGESTIONS", "Generate suggestions", "FORM", 283
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.GENERATE_SUGGESTIONS", "Generer les propositions", "FORM", 283
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SET_MISSING_KEYS", "Fehlende Keys setzen", "FORM", 284
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SET_MISSING_KEYS", "Set missing keys", "FORM", 284
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SET_MISSING_KEYS", "Definir les cles manquantes", "FORM", 284
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SET_SELECTED_KEY", "Key fuer Control setzen", "FORM", 285
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SET_SELECTED_KEY", "Set key for control", "FORM", 285
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SET_SELECTED_KEY", "Definir la cle pour le controle", "FORM", 285
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.REMOVE_SELECTED_KEY", "Key fuer Control entfernen", "FORM", 286
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.REMOVE_SELECTED_KEY", "Remove key from control", "FORM", 286
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.REMOVE_SELECTED_KEY", "Supprimer la cle du controle", "FORM", 286
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SAVE", "Aenderungen speichern", "FORM", 287
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SAVE", "Save changes", "FORM", 287
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SAVE", "Enregistrer les modifications", "FORM", 287
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CANCEL", "Schliessen / verwerfen", "FORM", 288
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CANCEL", "Close / discard", "FORM", 288
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CANCEL", "Fermer / abandonner", "FORM", 288
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CONTROL_NAME", "Control-Name", "FORM", 289
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CONTROL_NAME", "Control name", "FORM", 289
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CONTROL_NAME", "Nom du controle", "FORM", 289
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CONTROL_TYPE", "Control-Typ", "FORM", 290
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CONTROL_TYPE", "Control type", "FORM", 290
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CONTROL_TYPE", "Type de controle", "FORM", 290
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SOURCE_TEXT", "Quelltext", "FORM", 291
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SOURCE_TEXT", "Source text", "FORM", 291
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SOURCE_TEXT", "Texte source", "FORM", 291
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CURRENT_TAG", "Aktueller Tag", "FORM", 292
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CURRENT_TAG", "Current tag", "FORM", 292
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CURRENT_TAG", "Balise actuelle", "FORM", 292
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CURRENT_TRANSLATION_KEY", "Vorhandener Translation-Key", "FORM", 293
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CURRENT_TRANSLATION_KEY", "Existing translation key", "FORM", 293
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CURRENT_TRANSLATION_KEY", "Cle de traduction existante", "FORM", 293
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SUGGESTED_TRANSLATION_KEY", "Vorgeschlagener Translation-Key", "FORM", 294
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SUGGESTED_TRANSLATION_KEY", "Suggested translation key", "FORM", 294
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SUGGESTED_TRANSLATION_KEY", "Cle de traduction proposee", "FORM", 294
    EnsureTranslationSeed db, "DE-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.TAG_STATUS", "Tag-Status", "FORM", 295
    EnsureTranslationSeed db, "EN-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.TAG_STATUS", "Tag status", "FORM", 295
    EnsureTranslationSeed db, "FR-FR", "FORM.FRMFWTRANSLATIONTAGGENERATOR.TAG_STATUS", "Statut de balise", "FORM", 295

    EnsureTranslationSeed db, "DE-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_LOAD_ERROR", "Fehler beim Laden des Translation-Tag-Generators.", "MSG", 295
    EnsureTranslationSeed db, "EN-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_LOAD_ERROR", "Error loading the translation tag generator.", "MSG", 295
    EnsureTranslationSeed db, "FR-FR", "MSG.FW_TRANSLATION_TAG_GENERATOR_LOAD_ERROR", "Erreur lors du chargement du generateur de balises de traduction.", "MSG", 295
    EnsureTranslationSeed db, "DE-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_LOAD_FORM_ERROR", "Fehler beim Laden des ausgewaehlten Formulars.", "MSG", 296
    EnsureTranslationSeed db, "EN-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_LOAD_FORM_ERROR", "Error loading the selected form.", "MSG", 296
    EnsureTranslationSeed db, "FR-FR", "MSG.FW_TRANSLATION_TAG_GENERATOR_LOAD_FORM_ERROR", "Erreur lors du chargement du formulaire selectionne.", "MSG", 296
    EnsureTranslationSeed db, "DE-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_FORM_REQUIRED", "Bitte zuerst ein Formular auswaehlen.", "MSG", 297
    EnsureTranslationSeed db, "EN-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_FORM_REQUIRED", "Please select a form first.", "MSG", 297
    EnsureTranslationSeed db, "FR-FR", "MSG.FW_TRANSLATION_TAG_GENERATOR_FORM_REQUIRED", "Veuillez d'abord selectionner un formulaire.", "MSG", 297
    EnsureTranslationSeed db, "DE-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_SELECT_CONTROL", "Bitte zuerst ein Control auswaehlen.", "MSG", 298
    EnsureTranslationSeed db, "EN-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_SELECT_CONTROL", "Please select a control first.", "MSG", 298
    EnsureTranslationSeed db, "FR-FR", "MSG.FW_TRANSLATION_TAG_GENERATOR_SELECT_CONTROL", "Veuillez d'abord selectionner un controle.", "MSG", 298
    EnsureTranslationSeed db, "DE-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_NOT_TRANSLATABLE", "Das markierte Control ist nicht fuer einen Translation-Key geeignet.", "MSG", 299
    EnsureTranslationSeed db, "EN-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_NOT_TRANSLATABLE", "The selected control is not suitable for a translation key.", "MSG", 299
    EnsureTranslationSeed db, "FR-FR", "MSG.FW_TRANSLATION_TAG_GENERATOR_NOT_TRANSLATABLE", "Le controle selectionne ne convient pas pour une cle de traduction.", "MSG", 299
    EnsureTranslationSeed db, "DE-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_NO_MISSING_KEYS", "Keine fehlenden Translation-Keys zum Setzen gefunden.", "MSG", 300
    EnsureTranslationSeed db, "EN-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_NO_MISSING_KEYS", "No missing translation keys were found to set.", "MSG", 300
    EnsureTranslationSeed db, "FR-FR", "MSG.FW_TRANSLATION_TAG_GENERATOR_NO_MISSING_KEYS", "Aucune cle de traduction manquante a definir n'a ete trouvee.", "MSG", 300
    EnsureTranslationSeed db, "DE-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_SAVE_SUCCESS", "Translation-Tags wurden gespeichert.", "MSG", 301
    EnsureTranslationSeed db, "EN-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_SAVE_SUCCESS", "Translation tags were saved.", "MSG", 301
    EnsureTranslationSeed db, "FR-FR", "MSG.FW_TRANSLATION_TAG_GENERATOR_SAVE_SUCCESS", "Les balises de traduction ont ete enregistrees.", "MSG", 301
    EnsureTranslationSeed db, "DE-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_SAVE_ERROR", "Fehler beim Speichern der Translation-Tags.", "MSG", 302
    EnsureTranslationSeed db, "EN-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_SAVE_ERROR", "Error saving the translation tags.", "MSG", 302
    EnsureTranslationSeed db, "FR-FR", "MSG.FW_TRANSLATION_TAG_GENERATOR_SAVE_ERROR", "Erreur lors de l'enregistrement des balises de traduction.", "MSG", 302
    EnsureTranslationSeed db, "DE-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_DISCARD_CONFIRM", "Ungespeicherte Tag-Aenderungen verwerfen?", "MSG", 303
    EnsureTranslationSeed db, "EN-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_DISCARD_CONFIRM", "Discard unsaved tag changes?", "MSG", 303
    EnsureTranslationSeed db, "FR-FR", "MSG.FW_TRANSLATION_TAG_GENERATOR_DISCARD_CONFIRM", "Abandonner les modifications de balises non enregistrees ?", "MSG", 303
End Sub

Private Sub EnsureTranslationSeed( _
    ByVal db As DAO.Database, _
    ByVal LanguageCode As String, _
    ByVal translationKey As String, _
    ByVal TranslationValue As String, _
    ByVal moduleCode As String, _
    ByVal sortOrder As Long)

    If Not TranslationSeedExists(db, LanguageCode, translationKey) Then
        InsertTranslation db, LanguageCode, translationKey, TranslationValue, True, moduleCode, sortOrder
    End If
End Sub

Private Sub UpdateTranslationSeed( _
    ByVal db As DAO.Database, _
    ByVal LanguageCode As String, _
    ByVal translationKey As String, _
    ByVal TranslationValue As String, _
    ByVal isActive As Boolean, _
    ByVal moduleCode As String, _
    ByVal sortOrder As Long)

    Dim sqlStatement As String

    sqlStatement = "UPDATE fw_translation SET " & _
                   "translation_value = " & SqlText(TranslationValue) & ", " & _
                   "is_active = " & IIf(isActive, "True", "False") & ", " & _
                   "module_code = " & SqlNullableText(moduleCode) & ", " & _
                   "sort_order = " & CStr(sortOrder) & ", " & _
                   "updated_at = Now(), " & _
                   "updated_by = 'SYSTEM' " & _
                   "WHERE language_code = " & SqlText(LanguageCode) & " " & _
                   "AND translation_key = " & SqlText(translationKey) & ";"

    db.Execute sqlStatement, dbFailOnError
End Sub

Private Function TranslationSeedExists( _
    ByVal db As DAO.Database, _
    ByVal LanguageCode As String, _
    ByVal translationKey As String) As Boolean
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlStatement As String

    sqlStatement = "SELECT TOP 1 translation_key " & _
                   "FROM fw_translation " & _
                   "WHERE language_code = " & SqlText(LanguageCode) & " " & _
                   "AND translation_key = " & SqlText(translationKey) & ";"

    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)
    TranslationSeedExists = Not (rs.BOF And rs.EOF)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    TranslationSeedExists = False
End Function

Public Sub SeedTagHelp()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    db.Execute "DELETE FROM fw_tag_help", dbFailOnError

    InsertTagHelp db, "REQUIRED", "VALIDATION", "REQUIRED", _
        "Feld ist ein Pflichtfeld. Leere Werte sind nicht erlaubt.", _
        "REQUIRED", _
        "Markiert das zugehoerige Label mit einem *.", _
        10, True

    InsertTagHelp db, "NUMERIC", "VALIDATION", "NUMERIC", _
        "Wert muss numerisch sein, sofern ein Wert eingegeben wurde.", _
        "NUMERIC;MIN:0;MAX:100", _
        "Leere Werte sind erlaubt, solange REQUIRED nicht zusaetzlich gesetzt ist.", _
        20, True

    InsertTagHelp db, "INTEGER", "VALIDATION", "INTEGER", _
        "Wert muss eine ganze Zahl sein.", _
        "INTEGER;MIN:1;MAX:10", _
        "Ganzzahlen wie 1, 2 oder -5 sind gueltig; 1.5 ist ungueltig.", _
        30, True

    InsertTagHelp db, "DATE", "VALIDATION", "DATE", _
        "Wert muss ein gueltiges Datum sein.", _
        "REQUIRED;DATE", _
        "Leere Werte sind erlaubt, solange REQUIRED nicht zusaetzlich gesetzt ist.", _
        40, True

    InsertTagHelp db, "MIN", "VALIDATION", "MIN:<value>", _
        "Minimalwert fuer numerische Eingaben.", _
        "NUMERIC;MIN:0", _
        "Sollte zusammen mit NUMERIC oder INTEGER verwendet werden.", _
        50, True

    InsertTagHelp db, "MAX", "VALIDATION", "MAX:<value>", _
        "Maximalwert fuer numerische Eingaben.", _
        "NUMERIC;MAX:100", _
        "Sollte zusammen mit NUMERIC oder INTEGER verwendet werden.", _
        60, True

    InsertTagHelp db, "MINLEN", "VALIDATION", "MINLEN:<value>", _
        "Mindestlaenge fuer Texteingaben.", _
        "REQUIRED;MINLEN:3", _
        "Wirkt nur auf Textwerte.", _
        70, True

    InsertTagHelp db, "MAXLEN", "VALIDATION", "MAXLEN:<value>", _
        "Maximallaenge fuer Texteingaben.", _
        "MAXLEN:50", _
        "Wirkt nur auf Textwerte.", _
        80, True

    InsertTagHelp db, "HIDDEN", "BEHAVIOR", "HIDDEN", _
        "Blendet das Control aus.", _
        "HIDDEN", _
        "Ausgeblendete Controls werden aktuell nicht validiert.", _
        90, True

    InsertTagHelp db, "DISABLED", "BEHAVIOR", "DISABLED", _
        "Deaktiviert das Control.", _
        "DISABLED", _
        "Deaktivierte Controls werden aktuell nicht validiert.", _
        100, True

    InsertTagHelp db, "LOCKED", "BEHAVIOR", "LOCKED", _
        "Sperrt das Control fuer Bearbeitung.", _
        "LOCKED", _
        "Gesperrte Controls bleiben sichtbar und koennen weiterhin validiert werden.", _
        110, True

    InsertTagHelp db, "SETFOCUS", "BEHAVIOR", "SETFOCUS", _
        "Setzt beim Initialisieren den Fokus auf dieses Control.", _
        "SETFOCUS", _
        "Sinnvoll bei Formularen mit gesteuerter Startnavigation.", _
        120, True

    InsertTagHelp db, "READONLY", "FORM", "READONLY", _
        "Setzt das gesamte Formular in den Nur-Lesen-Modus.", _
        "READONLY", _
        "Betroffen sind Edits, Additions und Deletions.", _
        130, True

    InsertTagHelp db, "ROLE", "ACCESS", "ROLE:<role1,role2,...>", _
        "Steuert Sichtbarkeit oder Zugriff anhand von Rollen.", _
        "ROLE:ADMIN,ACCOUNTING", _
        "Eine passende Rolle reicht aus.", _
        140, True

    InsertTagHelp db, "MOD", "ACCESS", "MOD:<modulecode>", _
        "Bindet Formular oder Control an ein aktives Modul.", _
        "MOD:PROPERTY_MGMT", _
        "Wenn Modul nicht aktiv ist, wird Zugriff oder Initialisierung verhindert.", _
        150, True

    InsertTagHelp db, "TR", "I18N", "TR:<translationkey>", _
        "Verweist im Control.Tag auf einen Uebersetzungsschluessel.", _
        "TR:FORM.FRMAPPSHELL.APP_SUBTITLE", _
        "Der TR:-Marker wird vom Uebersetzungswerkzeug verwaltet, andere Tag-Segmente bleiben erhalten.", _
        160, True

    MsgBox "fw_tag_help wurde erfolgreich initialisiert.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Initialisieren von fw_tag_help: " & Err.description, vbExclamation
End Sub

Private Sub InsertTagHelp( _
    ByVal db As DAO.Database, _
    ByVal TokenKey As String, _
    ByVal Category As String, _
    ByVal SyntaxText As String, _
    ByVal DescriptionText As String, _
    ByVal ExampleText As String, _
    ByVal NotesText As String, _
    ByVal sortOrder As Long, _
    ByVal isActive As Boolean)

    Dim sqlStatement As String

    sqlStatement = "INSERT INTO fw_tag_help " & _
                   "(token_key, category, syntax_text, description_text, example_text, notes_text, sort_order, is_active, created_at, created_by, updated_at, updated_by) " & _
                   "VALUES (" & _
                   SqlText(TokenKey) & ", " & _
                   SqlText(Category) & ", " & _
                   SqlText(SyntaxText) & ", " & _
                   SqlText(DescriptionText) & ", " & _
                   SqlText(ExampleText) & ", " & _
                   SqlText(NotesText) & ", " & _
                   CStr(sortOrder) & ", " & _
                   IIf(isActive, "True", "False") & ", " & _
                   "Now(), 'SYSTEM', Now(), 'SYSTEM')"

    db.Execute sqlStatement, dbFailOnError
End Sub

