Attribute VB_Name = "modFwSetup"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFwSetup
' Purpose   : Provides initialization and seeding routines for framework data
'             such as translations, tag help definitions, and demo content.
' Author    : Codex
' Version   : 1.1.0
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

    MsgBox "fw_translation wurde erfolgreich initialisiert.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Initialisieren von fw_translation: " & Err.description, vbExclamation
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

Private Sub InsertTranslation( _
    ByVal db As DAO.Database, _
    ByVal LanguageCode As String, _
    ByVal TranslationKey As String, _
    ByVal TranslationValue As String, _
    ByVal IsActive As Boolean)

    Dim sqlStmt As String

    sqlStmt = "INSERT INTO fw_translation " & _
              "(language_code, translation_key, translation_value, is_active) " & _
              "VALUES (" & _
              "'" & EscapeSqlText(LanguageCode) & "', " & _
              "'" & EscapeSqlText(TranslationKey) & "', " & _
              "'" & EscapeSqlText(TranslationValue) & "', " & _
              IIf(IsActive, "True", "False") & ")"

    db.Execute sqlStmt, dbFailOnError
End Sub

Private Sub InsertAddressType( _
    ByVal db As DAO.Database, _
    ByVal addressTypeCode As String, _
    ByVal TranslationKey As String, _
    ByVal SortOrder As Long, _
    ByVal IsActive As Boolean)

    Dim sqlStatement As String

    sqlStatement = "INSERT INTO ref_address_type " & _
                   "(address_type_code, translation_key, sort_order, is_active, created_at, created_by, updated_at, updated_by) " & _
                   "VALUES (" & _
                   "'" & EscapeSqlText(addressTypeCode) & "', " & _
                   "'" & EscapeSqlText(TranslationKey) & "', " & _
                   CStr(SortOrder) & ", " & _
                   IIf(IsActive, "True", "False") & ", " & _
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
        "Verweist auf einen Uebersetzungsschluessel.", _
        "TR:LBL_CUSTOMER", _
        "Soll vom Tag-Composer erhalten, aber nicht ueberschrieben werden.", _
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
    ByVal SortOrder As Long, _
    ByVal IsActive As Boolean)

    Dim sqlStmt As String

    sqlStmt = "INSERT INTO fw_tag_help " & _
              "(TokenKey, Category, SyntaxText, DescriptionText, ExampleText, NotesText, SortOrder, IsActive) " & _
              "VALUES (" & _
              "'" & EscapeSqlText(TokenKey) & "', " & _
              "'" & EscapeSqlText(Category) & "', " & _
              "'" & EscapeSqlText(SyntaxText) & "', " & _
              "'" & EscapeSqlText(DescriptionText) & "', " & _
              "'" & EscapeSqlText(ExampleText) & "', " & _
              "'" & EscapeSqlText(NotesText) & "', " & _
              CStr(SortOrder) & ", " & _
              IIf(IsActive, "True", "False") & ")"

    db.Execute sqlStmt, dbFailOnError
End Sub

Private Function EscapeSqlText(ByVal Value As String) As String
    EscapeSqlText = Replace(Nz(Value, vbNullString), "'", "''")
End Function
