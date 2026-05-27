Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFwSetup
' Purpose   : Provides initialization and seeding routines for framework data
'             such as translations, tag help definitions, and demo content.
' Author    : Codex
' Version   : 1.6.0
' Notes     : Safe to re-run. Existing data will be replaced.
'===============================================================================

Private Const MODULE_NAME As String = "modFwSetup"

Public Sub SeedTranslations()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    db.Execute "DELETE FROM fw_translation", dbFailOnError

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

    MsgBox "fw_translation wurde erfolgreich initialisiert.", vbInformation
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
    db.Execute "DELETE FROM fw_translation WHERE translation_key Like 'VAT.CH.*'", dbFailOnError

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
    db.Execute "DELETE FROM fw_translation WHERE translation_key Like 'UNIT.*'", dbFailOnError

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
    db.Execute "DELETE FROM fw_translation WHERE translation_key Like 'ADDRESS_TYPE.*'", dbFailOnError

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
    db.Execute "DELETE FROM fw_translation WHERE translation_key Like 'SALUTATION.*'", dbFailOnError
    db.Execute "DELETE FROM fw_translation WHERE translation_key Like 'ADDRESSING_MODE.*'", dbFailOnError

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
    db.Execute "DELETE FROM fw_translation WHERE translation_key Like 'CONTACT_TYPE.*'", dbFailOnError

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
    EnsureTranslationSeed db, "DE-CH", "NAV.GROUP.SYSTEM", "System", "NAVIGATION", 160
    EnsureTranslationSeed db, "EN-US", "NAV.GROUP.SYSTEM", "System", "NAVIGATION", 160
    EnsureTranslationSeed db, "DE-CH", "NAV.FW_NAVIGATION_ADMIN", "Navigation verwalten", "NAVIGATION", 170
    EnsureTranslationSeed db, "EN-US", "NAV.FW_NAVIGATION_ADMIN", "Manage navigation", "NAVIGATION", 170

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

    EnsureTranslationSeed db, "DE-CH", "STATUS.READY", "Bereit", "STATUS", 230
    EnsureTranslationSeed db, "EN-US", "STATUS.READY", "Ready", "STATUS", 230
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

