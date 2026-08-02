Attribute VB_Name = "modFwSetup"
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
Private Const FORM_TRANSLATION_TAG_GENERATOR As String = "frmFwTranslationTagGenerator"
Private Const FORM_ADDRESS_COCKPIT As String = "frmAddressCockpit"
Private Const FORM_ORDER_LINES_SUBFORM As String = "sfrmOrderLines"

Public Function EnsureOrderSchema() As Boolean
    On Error GoTo ErrorHandler

    Dim workingDb As DAO.Database

    EnsureOrderSchema = False
    Set workingDb = modDb.GetSystemDatabase()
    If workingDb Is Nothing Then
        Exit Function
    End If

    If Not modBasicModuleSchema.EnsureOrderPhase1Schema() Then
        Exit Function
    End If

    If Not modMigrationPaymentTerms.ApplyPaymentTermsMigration() Then
        Exit Function
    End If

    If Not modOrderRepository.EnsureSalesOrderNumberRange(Year(Date)) Then
        Exit Function
    End If

    If Not modBasicModuleSchema.EnsureSystemLanguageReferenceSchema() Then
        Exit Function
    End If

    EnsureOrderDetailTranslations workingDb
    EnsureOrderLinesTranslations workingDb

    If Not TranslationSeedExists(workingDb, "de-CH", "FORM.FRMORDERDETAIL.VAT_MODE") Then
        Err.Raise vbObjectError + 6120, MODULE_NAME & ".EnsureOrderSchema", _
            "Order detail VAT-mode translations could not be ensured."
    End If

    EnsureOrderSchema = True
    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureOrderSchema", _
        "Sales-order schema, SO number range, ref_language, and order translations ensured successfully."
    Exit Function

ErrorHandler:
    EnsureOrderSchema = False
    modErrorHandler.HandleError MODULE_NAME, "EnsureOrderSchema", Err
End Function

Public Function EnsureBasicModuleSchema() As Boolean
    EnsureBasicModuleSchema = EnsureOrderSchema()
End Function

Public Sub SeedTranslations()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = modDb.GetSystemDatabase()
    If db Is Nothing Then
        Exit Sub
    End If

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
    EnsureTranslationTagGeneratorTranslations db
    EnsureTranslationTagGeneratorTags
    EnsureAddressCockpitTranslations db
    EnsureAddressCockpitTags
    EnsureOrderDetailTranslations db
    EnsureOrderLinesTranslations db
    EnsureOrderLinesTags

    MsgBox "fw_translation wurde erfolgreich ergaenzt.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Initialisieren von fw_translation: " & Err.description, vbExclamation
End Sub

Public Sub EnsureTranslationTagGeneratorTags()
    On Error GoTo ErrorHandler

    Dim metadataItems As Collection
    Dim metadata As Variant
    Dim controlTagMap As Object
    Dim ControlName As String
    Dim currentTag As String
    Dim updatedTag As String
    Dim updatedCount As Long

    Set metadataItems = modFwComposerService.GetFormControlMetadata(FORM_TRANSLATION_TAG_GENERATOR, True)
    Set controlTagMap = CreateObject("Scripting.Dictionary")
    controlTagMap.CompareMode = vbTextCompare

    For Each metadata In metadataItems
        ControlName = Trim$(modDaoHelper.NzString(metadata("control_name")))
        If LenB(ControlName) = 0 Then
            GoTo NextControl
        End If

        currentTag = modDaoHelper.NzString(metadata("current_tag"))
        updatedTag = modFwTranslationRuntime.SetTranslationKeyInTag( _
            currentTag, _
            BuildTagGeneratorTranslationKey(ControlName))

        If StrComp(updatedTag, currentTag, vbBinaryCompare) <> 0 Then
            controlTagMap(ControlName) = updatedTag
        End If

NextControl:
    Next metadata

    If controlTagMap.count > 0 Then
        If Not modFwComposerService.SaveControlTagsToObject(modFwComposerService.OBJECT_TYPE_FORM, FORM_TRANSLATION_TAG_GENERATOR, controlTagMap, updatedCount) Then
            Err.Raise vbObjectError + 6110, MODULE_NAME & ".EnsureTranslationTagGeneratorTags", _
                "Failed to persist frmFwTranslationTagGenerator tags."
        End If
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureTranslationTagGeneratorTags", _
        "Translation tags ensured for frmFwTranslationTagGenerator. updated_count=" & CStr(updatedCount) & "."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureTranslationTagGeneratorTags", Err
End Sub

Public Sub EnsureTranslationTagGeneratorTranslations(Optional ByVal db As DAO.Database = Nothing)
    On Error GoTo ErrorHandler

    Dim workingDb As DAO.Database

    If db Is Nothing Then
        Set workingDb = modDb.GetSystemDatabase()
    Else
        Set workingDb = db
    End If

    EnsureTranslationSeed workingDb, "de-CH", "NAV.FW_TRANSLATION_TAG_GENERATOR", "Translation-Tags", "NAVIGATION", 240
    EnsureTranslationSeed workingDb, "en-US", "NAV.FW_TRANSLATION_TAG_GENERATOR", "Translation tags", "NAVIGATION", 240
    EnsureTranslationSeed workingDb, "fr-CH", "NAV.FW_TRANSLATION_TAG_GENERATOR", "Balises de traduction", "NAVIGATION", 240

    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.FORM_TITLE", "Translation-Tag-Generator", "FORM", 278
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.FORM_TITLE", "Translation tag generator", "FORM", 278
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.FORM_TITLE", "Generateur de balises de traduction", "FORM", 278
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.FORM_NAME", "Formular", "FORM", 279
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.FORM_NAME", "Form", "FORM", 279
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.FORM_NAME", "Formulaire", "FORM", 279
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.INCLUDE_HIDDEN", "Versteckte Controls einschliessen", "FORM", 280
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.INCLUDE_HIDDEN", "Include hidden controls", "FORM", 280
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.INCLUDE_HIDDEN", "Inclure les controles masques", "FORM", 280
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SHOW_ALL_CONTROLS", "Alle Controls anzeigen", "FORM", 281
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SHOW_ALL_CONTROLS", "Show all controls", "FORM", 281
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SHOW_ALL_CONTROLS", "Afficher tous les controles", "FORM", 281
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.LOAD_FORM", "Formular laden", "FORM", 282
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.LOAD_FORM", "Load form", "FORM", 282
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.LOAD_FORM", "Charger le formulaire", "FORM", 282
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.GENERATE_SUGGESTIONS", "Vorschlaege generieren", "FORM", 283
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.GENERATE_SUGGESTIONS", "Generate suggestions", "FORM", 283
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.GENERATE_SUGGESTIONS", "Generer les propositions", "FORM", 283
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SET_MISSING_KEYS", "Fehlende Keys setzen", "FORM", 284
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SET_MISSING_KEYS", "Set missing keys", "FORM", 284
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SET_MISSING_KEYS", "Definir les cles manquantes", "FORM", 284
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SET_SELECTED_KEY", "Key fuer Control setzen", "FORM", 285
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SET_SELECTED_KEY", "Set key for control", "FORM", 285
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SET_SELECTED_KEY", "Definir la cle pour le controle", "FORM", 285
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.REMOVE_SELECTED_KEY", "Key fuer Control entfernen", "FORM", 286
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.REMOVE_SELECTED_KEY", "Remove key from control", "FORM", 286
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.REMOVE_SELECTED_KEY", "Supprimer la cle du controle", "FORM", 286
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SAVE", "Aenderungen speichern", "FORM", 287
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SAVE", "Save changes", "FORM", 287
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SAVE", "Enregistrer les modifications", "FORM", 287
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CANCEL", "Schliessen / verwerfen", "FORM", 288
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CANCEL", "Close / discard", "FORM", 288
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CANCEL", "Fermer / abandonner", "FORM", 288
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CONTROL_NAME", "Control-Name", "FORM", 289
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CONTROL_NAME", "Control name", "FORM", 289
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CONTROL_NAME", "Nom du controle", "FORM", 289
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CONTROL_TYPE", "Control-Typ", "FORM", 290
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CONTROL_TYPE", "Control type", "FORM", 290
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CONTROL_TYPE", "Type de controle", "FORM", 290
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SOURCE_TEXT", "Quelltext", "FORM", 291
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SOURCE_TEXT", "Source text", "FORM", 291
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SOURCE_TEXT", "Texte source", "FORM", 291
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CURRENT_TAG", "Aktueller Tag", "FORM", 292
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CURRENT_TAG", "Current tag", "FORM", 292
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CURRENT_TAG", "Balise actuelle", "FORM", 292
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CURRENT_TRANSLATION_KEY", "Vorhandener Translation-Key", "FORM", 293
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CURRENT_TRANSLATION_KEY", "Existing translation key", "FORM", 293
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.CURRENT_TRANSLATION_KEY", "Cle de traduction existante", "FORM", 293
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SUGGESTED_TRANSLATION_KEY", "Vorgeschlagener Translation-Key", "FORM", 294
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SUGGESTED_TRANSLATION_KEY", "Suggested translation key", "FORM", 294
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.SUGGESTED_TRANSLATION_KEY", "Cle de traduction proposee", "FORM", 294
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.TAG_STATUS", "Tag-Status", "FORM", 295
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMFWTRANSLATIONTAGGENERATOR.TAG_STATUS", "Tag status", "FORM", 295
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMFWTRANSLATIONTAGGENERATOR.TAG_STATUS", "Statut de balise", "FORM", 295

    EnsureTranslationSeed workingDb, "de-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_LOAD_ERROR", "Fehler beim Laden des Translation-Tag-Generators.", "MSG", 295
    EnsureTranslationSeed workingDb, "en-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_LOAD_ERROR", "Error loading the translation tag generator.", "MSG", 295
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_LOAD_ERROR", "Erreur lors du chargement du generateur de balises de traduction.", "MSG", 295
    EnsureTranslationSeed workingDb, "de-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_LOAD_FORM_ERROR", "Fehler beim Laden des ausgewaehlten Formulars.", "MSG", 296
    EnsureTranslationSeed workingDb, "en-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_LOAD_FORM_ERROR", "Error loading the selected form.", "MSG", 296
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_LOAD_FORM_ERROR", "Erreur lors du chargement du formulaire selectionne.", "MSG", 296
    EnsureTranslationSeed workingDb, "de-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_FORM_REQUIRED", "Bitte zuerst ein Formular auswaehlen.", "MSG", 297
    EnsureTranslationSeed workingDb, "en-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_FORM_REQUIRED", "Please select a form first.", "MSG", 297
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_FORM_REQUIRED", "Veuillez d'abord selectionner un formulaire.", "MSG", 297
    EnsureTranslationSeed workingDb, "de-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_SELECT_CONTROL", "Bitte zuerst ein Control auswaehlen.", "MSG", 298
    EnsureTranslationSeed workingDb, "en-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_SELECT_CONTROL", "Please select a control first.", "MSG", 298
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_SELECT_CONTROL", "Veuillez d'abord selectionner un controle.", "MSG", 298
    EnsureTranslationSeed workingDb, "de-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_NOT_TRANSLATABLE", "Das markierte Control ist nicht fuer einen Translation-Key geeignet.", "MSG", 299
    EnsureTranslationSeed workingDb, "en-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_NOT_TRANSLATABLE", "The selected control is not suitable for a translation key.", "MSG", 299
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_NOT_TRANSLATABLE", "Le controle selectionne ne convient pas pour une cle de traduction.", "MSG", 299
    EnsureTranslationSeed workingDb, "de-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_NO_MISSING_KEYS", "Keine fehlenden Translation-Keys zum Setzen gefunden.", "MSG", 300
    EnsureTranslationSeed workingDb, "en-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_NO_MISSING_KEYS", "No missing translation keys were found to set.", "MSG", 300
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_NO_MISSING_KEYS", "Aucune cle de traduction manquante a definir n'a ete trouvee.", "MSG", 300
    EnsureTranslationSeed workingDb, "de-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_SAVE_SUCCESS", "Translation-Tags wurden gespeichert.", "MSG", 301
    EnsureTranslationSeed workingDb, "en-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_SAVE_SUCCESS", "Translation tags were saved.", "MSG", 301
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_SAVE_SUCCESS", "Les balises de traduction ont ete enregistrees.", "MSG", 301
    EnsureTranslationSeed workingDb, "de-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_SAVE_ERROR", "Fehler beim Speichern der Translation-Tags.", "MSG", 302
    EnsureTranslationSeed workingDb, "en-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_SAVE_ERROR", "Error saving the translation tags.", "MSG", 302
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_SAVE_ERROR", "Erreur lors de l'enregistrement des balises de traduction.", "MSG", 302
    EnsureTranslationSeed workingDb, "de-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_DISCARD_CONFIRM", "Ungespeicherte Tag-Aenderungen verwerfen?", "MSG", 303
    EnsureTranslationSeed workingDb, "en-US", "MSG.FW_TRANSLATION_TAG_GENERATOR_DISCARD_CONFIRM", "Discard unsaved tag changes?", "MSG", 303
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.FW_TRANSLATION_TAG_GENERATOR_DISCARD_CONFIRM", "Abandonner les modifications de balises non enregistrees ?", "MSG", 303

    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureTranslationTagGeneratorTranslations", _
        "Translation seeds ensured for frmFwTranslationTagGenerator."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureTranslationTagGeneratorTranslations", Err
End Sub

Public Sub EnsureAddressCockpitTranslations(Optional ByVal db As DAO.Database = Nothing)
    On Error GoTo ErrorHandler

    Dim workingDb As DAO.Database

    If db Is Nothing Then
        Set workingDb = modDb.GetSystemDatabase()
    Else
        Set workingDb = db
    End If

    EnsureTranslationSeed workingDb, "de-CH", "NAV.ADDRESS_COCKPIT", "Adress-Cockpit", "NAVIGATION", 21
    EnsureTranslationSeed workingDb, "en-US", "NAV.ADDRESS_COCKPIT", "Address cockpit", "NAVIGATION", 21
    EnsureTranslationSeed workingDb, "fr-CH", "NAV.ADDRESS_COCKPIT", "Cockpit adresse", "NAVIGATION", 21

    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.FORM_TITLE", "Adress-Cockpit", "FORM", 320
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.FORM_TITLE", "Address cockpit", "FORM", 320
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.FORM_TITLE", "Cockpit adresse", "FORM", 320
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.ADDRESS", "Adresse", "FORM", 321
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.ADDRESS", "Address", "FORM", 321
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.ADDRESS", "Adresse", "FORM", 321
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.CONTACT", "Kontakt", "FORM", 322
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.CONTACT", "Contact", "FORM", 322
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.CONTACT", "Contact", "FORM", 322
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.STATUS", "Status", "FORM", 323
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.STATUS", "Status", "FORM", 323
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.STATUS", "Statut", "FORM", 323
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.LOCK_HINT", "Adresse ist gesperrt", "FORM", 324
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.LOCK_HINT", "Address is locked", "FORM", 324
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.LOCK_HINT", "L'adresse est bloquee", "FORM", 324
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.OPEN_INVOICES", "Offene Rechnungen", "FORM", 325
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.OPEN_INVOICES", "Open invoices", "FORM", 325
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.OPEN_INVOICES", "Factures ouvertes", "FORM", 325
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.OVERDUE_ITEMS", "Ueberfaellige Posten", "FORM", 326
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.OVERDUE_ITEMS", "Overdue items", "FORM", 326
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.OVERDUE_ITEMS", "Postes echus", "FORM", 326
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.DUNNINGS", "Mahnungen", "FORM", 327
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.DUNNINGS", "Dunnings", "FORM", 327
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.DUNNINGS", "Rappels", "FORM", 327
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.SALES_CURRENT_YEAR", "Umsatz laufendes Jahr", "FORM", 328
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.SALES_CURRENT_YEAR", "Sales current year", "FORM", 328
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.SALES_CURRENT_YEAR", "Chiffre d'affaires annee en cours", "FORM", 328
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.OPEN_ORDERS", "Offene Auftraege", "FORM", 329
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.OPEN_ORDERS", "Open orders", "FORM", 329
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.OPEN_ORDERS", "Commandes ouvertes", "FORM", 329
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.ACTIVE_SUBSCRIPTIONS", "Aktive Abos", "FORM", 330
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.ACTIVE_SUBSCRIPTIONS", "Active subscriptions", "FORM", 330
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.ACTIVE_SUBSCRIPTIONS", "Abonnements actifs", "FORM", 330
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.LAST_ACTIVITY", "Letzte Aktivitaet", "FORM", 331
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.LAST_ACTIVITY", "Last activity", "FORM", 331
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.LAST_ACTIVITY", "Derniere activite", "FORM", 331
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.NEW_ORDER", "Neue Bestellung", "FORM", 332
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.NEW_ORDER", "New order", "FORM", 332
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.NEW_ORDER", "Nouvelle commande", "FORM", 332
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.MANAGE_ORDERS", "Bestellungen verwalten", "FORM", 333
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.MANAGE_ORDERS", "Manage orders", "FORM", 333
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.MANAGE_ORDERS", "Gerer les commandes", "FORM", 333
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.NEW_SUBSCRIPTION", "Neues Abo", "FORM", 334
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.NEW_SUBSCRIPTION", "New subscription", "FORM", 334
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.NEW_SUBSCRIPTION", "Nouvel abonnement", "FORM", 334
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.MANAGE_SUBSCRIPTIONS", "Abos verwalten", "FORM", 335
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.MANAGE_SUBSCRIPTIONS", "Manage subscriptions", "FORM", 335
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.MANAGE_SUBSCRIPTIONS", "Gerer les abonnements", "FORM", 335
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.ACCOUNT_STATEMENT", "Kontoauszug", "FORM", 336
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.ACCOUNT_STATEMENT", "Account statement", "FORM", 336
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.ACCOUNT_STATEMENT", "Releve de compte", "FORM", 336
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.CAPTURE_PAYMENT", "Zahlungseingang erfassen", "FORM", 337
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.CAPTURE_PAYMENT", "Capture payment", "FORM", 337
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.CAPTURE_PAYMENT", "Saisir un paiement", "FORM", 337
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.DUNNINGS_ACTION", "Mahnungen", "FORM", 338
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.DUNNINGS_ACTION", "Dunnings", "FORM", 338
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.DUNNINGS_ACTION", "Rappels", "FORM", 338
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.NOTES", "Notizen", "FORM", 339
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.NOTES", "Notes", "FORM", 339
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.NOTES", "Notes", "FORM", 339
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.EMAILS", "E-Mails", "FORM", 340
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.EMAILS", "Emails", "FORM", 340
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.EMAILS", "E-mails", "FORM", 340
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.REPORTS", "Reports", "FORM", 341
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.REPORTS", "Reports", "FORM", 341
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.REPORTS", "Rapports", "FORM", 341
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.N_A", "n/a", "FORM", 342
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.N_A", "n/a", "FORM", 342
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.N_A", "n/a", "FORM", 342
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.STATUS_ACTIVE", "Aktiv", "FORM", 343
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.STATUS_ACTIVE", "Active", "FORM", 343
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.STATUS_ACTIVE", "Actif", "FORM", 343
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMADDRESSCOCKPIT.STATUS_INACTIVE", "Inaktiv", "FORM", 344
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMADDRESSCOCKPIT.STATUS_INACTIVE", "Inactive", "FORM", 344
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMADDRESSCOCKPIT.STATUS_INACTIVE", "Inactif", "FORM", 344

    EnsureTranslationSeed workingDb, "de-CH", "MSG.ADDRESS_COCKPIT_SELECT_ADDRESS_FIRST", "Bitte zuerst eine Adresse auswaehlen.", "MSG", 345
    EnsureTranslationSeed workingDb, "en-US", "MSG.ADDRESS_COCKPIT_SELECT_ADDRESS_FIRST", "Please select an address first.", "MSG", 345
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.ADDRESS_COCKPIT_SELECT_ADDRESS_FIRST", "Veuillez d'abord selectionner une adresse.", "MSG", 345
    EnsureTranslationSeed workingDb, "de-CH", "MSG.ADDRESS_COCKPIT_ORDER_CREATE_FAILED", "Die Bestellung konnte nicht erstellt werden.", "MSG", 346
    EnsureTranslationSeed workingDb, "en-US", "MSG.ADDRESS_COCKPIT_ORDER_CREATE_FAILED", "The order could not be created.", "MSG", 346
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.ADDRESS_COCKPIT_ORDER_CREATE_FAILED", "La commande n'a pas pu etre creee.", "MSG", 346
    EnsureTranslationSeed workingDb, "de-CH", "MSG.ADDRESS_COCKPIT_ORDER_DETAIL_MISSING", "Die Bestellung wurde erstellt, aber frmOrderDetail ist nicht verfuegbar.", "MSG", 347
    EnsureTranslationSeed workingDb, "en-US", "MSG.ADDRESS_COCKPIT_ORDER_DETAIL_MISSING", "The order was created, but frmOrderDetail is not available.", "MSG", 347
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.ADDRESS_COCKPIT_ORDER_DETAIL_MISSING", "La commande a ete creee, mais frmOrderDetail n'est pas disponible.", "MSG", 347

    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureAddressCockpitTranslations", _
        "Translation seeds ensured for frmAddressCockpit."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureAddressCockpitTranslations", Err
End Sub

Public Sub EnsureAddressCockpitTags()
    On Error GoTo ErrorHandler

    Dim metadataItems As Collection
    Dim metadata As Variant
    Dim controlTagMap As Object
    Dim ControlName As String
    Dim currentTag As String
    Dim updatedTag As String
    Dim updatedCount As Long

    If Not FormObjectExists(FORM_ADDRESS_COCKPIT) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".EnsureAddressCockpitTags", _
            "Form '" & FORM_ADDRESS_COCKPIT & "' is not available in the current Access project. Tag ensure skipped."
        Exit Sub
    End If

    Set metadataItems = modFwComposerService.GetFormControlMetadata(FORM_ADDRESS_COCKPIT, True)
    Set controlTagMap = CreateObject("Scripting.Dictionary")
    controlTagMap.CompareMode = vbTextCompare

    For Each metadata In metadataItems
        ControlName = Trim$(modDaoHelper.NzString(metadata("control_name")))
        If LenB(ControlName) = 0 Then
            GoTo NextControl
        End If

        currentTag = modDaoHelper.NzString(metadata("current_tag"))
        updatedTag = modFwTranslationRuntime.SetTranslationKeyInTag( _
            currentTag, _
            BuildAddressCockpitTranslationKey(ControlName))

        If StrComp(updatedTag, currentTag, vbBinaryCompare) <> 0 Then
            controlTagMap(ControlName) = updatedTag
        End If

NextControl:
    Next metadata

    If controlTagMap.count > 0 Then
        If Not modFwComposerService.SaveControlTagsToObject(modFwComposerService.OBJECT_TYPE_FORM, FORM_ADDRESS_COCKPIT, controlTagMap, updatedCount) Then
            Err.Raise vbObjectError + 6111, MODULE_NAME & ".EnsureAddressCockpitTags", _
                "Failed to persist frmAddressCockpit tags."
        End If
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureAddressCockpitTags", _
        "Translation tags ensured for frmAddressCockpit. updated_count=" & CStr(updatedCount) & "."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureAddressCockpitTags", Err
End Sub

Public Sub EnsureOrderDetailTranslations(Optional ByVal db As DAO.Database = Nothing)
    On Error GoTo ErrorHandler

    Dim workingDb As DAO.Database

    If db Is Nothing Then
        Set workingDb = modDb.GetSystemDatabase()
    Else
        Set workingDb = db
    End If

    EnsureTranslationSeed workingDb, "de-CH", "NAV.ORDER_DETAIL", "Bestellung", "NAVIGATION", 348
    EnsureTranslationSeed workingDb, "en-US", "NAV.ORDER_DETAIL", "Order", "NAVIGATION", 348
    EnsureTranslationSeed workingDb, "fr-CH", "NAV.ORDER_DETAIL", "Commande", "NAVIGATION", 348

    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.FORM_TITLE", "Bestellung", "FORM", 349
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.FORM_TITLE", "Order", "FORM", 349
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.FORM_TITLE", "Commande", "FORM", 349
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.ORDER_NO", "Bestell-Nr.", "FORM", 350
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.ORDER_NO", "Order no.", "FORM", 350
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.ORDER_NO", "No de commande", "FORM", 350
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.CUSTOMER_NAME", "Kunde", "FORM", 351
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.CUSTOMER_NAME", "Customer", "FORM", 351
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.CUSTOMER_NAME", "Client", "FORM", 351
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.ORDER_DATE", "Bestelldatum", "FORM", 352
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.ORDER_DATE", "Order date", "FORM", 352
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.ORDER_DATE", "Date de commande", "FORM", 352
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.DELIVERY_DATE", "Lieferdatum", "FORM", 353
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.DELIVERY_DATE", "Delivery date", "FORM", 353
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.DELIVERY_DATE", "Date de livraison", "FORM", 353
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.VALID_UNTIL", "Gueltig bis", "FORM", 354
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.VALID_UNTIL", "Valid until", "FORM", 354
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.VALID_UNTIL", "Valable jusqu'au", "FORM", 354
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.REFERENCE_TEXT", "Referenz", "FORM", 355
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.REFERENCE_TEXT", "Reference", "FORM", 355
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.REFERENCE_TEXT", "Reference", "FORM", 355
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.EXTERNAL_REFERENCE", "Externe Referenz", "FORM", 356
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.EXTERNAL_REFERENCE", "External reference", "FORM", 356
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.EXTERNAL_REFERENCE", "Reference externe", "FORM", 356
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.LANGUAGE_CODE", "Sprache", "FORM", 357
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.LANGUAGE_CODE", "Language", "FORM", 357
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.LANGUAGE_CODE", "Langue", "FORM", 357
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.CURRENCY_CODE", "Waehrung", "FORM", 358
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.CURRENCY_CODE", "Currency", "FORM", 358
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.CURRENCY_CODE", "Devise", "FORM", 358
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.PAYMENT_TERM_CODE", "Zahlungsbedingung", "FORM", 359
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.PAYMENT_TERM_CODE", "Payment term", "FORM", 359
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.PAYMENT_TERM_CODE", "Condition de paiement", "FORM", 359
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.VAT_MODE", "MWST-Modus", "FORM", 360
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.VAT_MODE", "VAT mode", "FORM", 360
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.VAT_MODE", "Mode TVA", "FORM", 360
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.VAT_MODE_EXCLUSIVE", "Exklusive MwSt.", "FORM", 3601
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.VAT_MODE_EXCLUSIVE", "Exclusive VAT", "FORM", 3601
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.VAT_MODE_EXCLUSIVE", "TVA exclue", "FORM", 3601
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.VAT_MODE_INCLUSIVE", "Inklusive MwSt.", "FORM", 3602
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.VAT_MODE_INCLUSIVE", "Inclusive VAT", "FORM", 3602
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.VAT_MODE_INCLUSIVE", "TVA incluse", "FORM", 3602
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.VAT_MODE_NONE", "Keine MwSt.", "FORM", 3603
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.VAT_MODE_NONE", "No VAT", "FORM", 3603
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.VAT_MODE_NONE", "Sans TVA", "FORM", 3603
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.ORDER_STATUS_CODE", "Status", "FORM", 361
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.ORDER_STATUS_CODE", "Status", "FORM", 361
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.ORDER_STATUS_CODE", "Statut", "FORM", 361
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.NOTES_TEXT", "Notizen", "FORM", 362
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.NOTES_TEXT", "Notes", "FORM", 362
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.NOTES_TEXT", "Notes", "FORM", 362
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.INTERNAL_NOTES_TEXT", "Interne Notizen", "FORM", 363
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.INTERNAL_NOTES_TEXT", "Internal notes", "FORM", 363
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.INTERNAL_NOTES_TEXT", "Notes internes", "FORM", 363
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.SUBTOTAL_NET_AMOUNT", "Zwischentotal netto", "FORM", 364
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.SUBTOTAL_NET_AMOUNT", "Subtotal net", "FORM", 364
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.SUBTOTAL_NET_AMOUNT", "Sous-total net", "FORM", 364
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.NET_AMOUNT", "Netto", "FORM", 365
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.NET_AMOUNT", "Net", "FORM", 365
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.NET_AMOUNT", "Net", "FORM", 365
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.VAT_AMOUNT", "MWST", "FORM", 366
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.VAT_AMOUNT", "VAT", "FORM", 366
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.VAT_AMOUNT", "TVA", "FORM", 366
    EnsureTranslationSeed workingDb, "de-CH", "FORM.FRMORDERDETAIL.GROSS_AMOUNT", "Brutto", "FORM", 367
    EnsureTranslationSeed workingDb, "en-US", "FORM.FRMORDERDETAIL.GROSS_AMOUNT", "Gross", "FORM", 367
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.FRMORDERDETAIL.GROSS_AMOUNT", "Brut", "FORM", 367

    EnsureTranslationSeed workingDb, "de-CH", "MSG.ORDER_DETAIL_MISSING_ID", "Keine Bestellung uebergeben.", "MSG", 368
    EnsureTranslationSeed workingDb, "en-US", "MSG.ORDER_DETAIL_MISSING_ID", "No order was provided.", "MSG", 368
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.ORDER_DETAIL_MISSING_ID", "Aucune commande n'a ete transmise.", "MSG", 368
    EnsureTranslationSeed workingDb, "de-CH", "MSG.ORDER_DETAIL_INVALID_ID", "Die uebergebene Bestell-ID ist ungueltig.", "MSG", 369
    EnsureTranslationSeed workingDb, "en-US", "MSG.ORDER_DETAIL_INVALID_ID", "The provided order id is invalid.", "MSG", 369
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.ORDER_DETAIL_INVALID_ID", "L'identifiant de commande transmis est invalide.", "MSG", 369
    EnsureTranslationSeed workingDb, "de-CH", "MSG.ORDER_DETAIL_NOT_FOUND", "Die Bestellung konnte nicht gefunden werden.", "MSG", 370
    EnsureTranslationSeed workingDb, "en-US", "MSG.ORDER_DETAIL_NOT_FOUND", "The order could not be found.", "MSG", 370
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.ORDER_DETAIL_NOT_FOUND", "La commande n'a pas pu etre trouvee.", "MSG", 370
    EnsureTranslationSeed workingDb, "de-CH", "MSG.ORDER_DETAIL_SAVE_ERROR", "Fehler beim Speichern der Bestellung.", "MSG", 371
    EnsureTranslationSeed workingDb, "en-US", "MSG.ORDER_DETAIL_SAVE_ERROR", "Error saving the order.", "MSG", 371
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.ORDER_DETAIL_SAVE_ERROR", "Erreur lors de l'enregistrement de la commande.", "MSG", 371
    EnsureTranslationSeed workingDb, "de-CH", "MSG.ORDER_DETAIL_CANCEL_CONFIRM", "Aenderungen verwerfen?", "MSG", 372
    EnsureTranslationSeed workingDb, "en-US", "MSG.ORDER_DETAIL_CANCEL_CONFIRM", "Discard changes?", "MSG", 372
    EnsureTranslationSeed workingDb, "fr-CH", "MSG.ORDER_DETAIL_CANCEL_CONFIRM", "Abandonner les modifications ?", "MSG", 372

    EnsureTranslationSeed workingDb, "de-CH", "COMMON.SAVE", "Speichern", "COMMON", 373
    EnsureTranslationSeed workingDb, "en-US", "COMMON.SAVE", "Save", "COMMON", 373
    EnsureTranslationSeed workingDb, "fr-CH", "COMMON.SAVE", "Enregistrer", "COMMON", 373
    EnsureTranslationSeed workingDb, "de-CH", "COMMON.CANCEL", "Abbrechen", "COMMON", 374
    EnsureTranslationSeed workingDb, "en-US", "COMMON.CANCEL", "Cancel", "COMMON", 374
    EnsureTranslationSeed workingDb, "fr-CH", "COMMON.CANCEL", "Annuler", "COMMON", 374

    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureOrderDetailTranslations", _
        "Translation seeds ensured for frmOrderDetail."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureOrderDetailTranslations", Err
End Sub

Public Sub EnsureOrderLinesTranslations(Optional ByVal db As DAO.Database = Nothing)
    On Error GoTo ErrorHandler

    Dim workingDb As DAO.Database

    If db Is Nothing Then
        Set workingDb = modDb.GetSystemDatabase()
    Else
        Set workingDb = db
    End If

    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.LINE_NO", "Pos.", "FORM", 375
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.LINE_NO", "Line", "FORM", 375
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.LINE_NO", "Pos.", "FORM", 375
    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.ARTICLE_ID", "Artikel-ID", "FORM", 376
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.ARTICLE_ID", "Article id", "FORM", 376
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.ARTICLE_ID", "Id article", "FORM", 376
    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.ARTICLE_NO", "Artikel-Nr.", "FORM", 377
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.ARTICLE_NO", "Article no.", "FORM", 377
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.ARTICLE_NO", "No article", "FORM", 377
    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.DESCRIPTION_TEXT", "Beschreibung", "FORM", 378
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.DESCRIPTION_TEXT", "Description", "FORM", 378
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.DESCRIPTION_TEXT", "Description", "FORM", 378
    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.QUANTITY", "Menge", "FORM", 379
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.QUANTITY", "Quantity", "FORM", 379
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.QUANTITY", "Quantite", "FORM", 379
    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.UNIT_CODE", "Einheit", "FORM", 380
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.UNIT_CODE", "Unit", "FORM", 380
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.UNIT_CODE", "Unite", "FORM", 380
    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.UNIT_PRICE", "Preis", "FORM", 381
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.UNIT_PRICE", "Price", "FORM", 381
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.UNIT_PRICE", "Prix", "FORM", 381
    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.DISCOUNT_TYPE", "Rabattart", "FORM", 382
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.DISCOUNT_TYPE", "Discount type", "FORM", 382
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.DISCOUNT_TYPE", "Type de rabais", "FORM", 382
    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.DISCOUNT_VALUE", "Rabatt", "FORM", 383
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.DISCOUNT_VALUE", "Discount", "FORM", 383
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.DISCOUNT_VALUE", "Rabais", "FORM", 383
    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.SURCHARGE_TYPE", "Zuschlagsart", "FORM", 384
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.SURCHARGE_TYPE", "Surcharge type", "FORM", 384
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.SURCHARGE_TYPE", "Type de supplement", "FORM", 384
    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.SURCHARGE_VALUE", "Zuschlag", "FORM", 385
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.SURCHARGE_VALUE", "Surcharge", "FORM", 385
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.SURCHARGE_VALUE", "Supplement", "FORM", 385
    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.VAT_CODE", "MWST-Code", "FORM", 386
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.VAT_CODE", "VAT code", "FORM", 386
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.VAT_CODE", "Code TVA", "FORM", 386
    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.VAT_RATE", "MWST-Satz", "FORM", 387
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.VAT_RATE", "VAT rate", "FORM", 387
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.VAT_RATE", "Taux TVA", "FORM", 387
    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.LINE_NET_AMOUNT", "Netto", "FORM", 388
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.LINE_NET_AMOUNT", "Net", "FORM", 388
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.LINE_NET_AMOUNT", "Net", "FORM", 388
    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.LINE_VAT_AMOUNT", "MWST", "FORM", 389
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.LINE_VAT_AMOUNT", "VAT", "FORM", 389
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.LINE_VAT_AMOUNT", "TVA", "FORM", 389
    EnsureTranslationSeed workingDb, "de-CH", "FORM.SFRMORDERLINES.LINE_GROSS_AMOUNT", "Brutto", "FORM", 390
    EnsureTranslationSeed workingDb, "en-US", "FORM.SFRMORDERLINES.LINE_GROSS_AMOUNT", "Gross", "FORM", 390
    EnsureTranslationSeed workingDb, "fr-CH", "FORM.SFRMORDERLINES.LINE_GROSS_AMOUNT", "Brut", "FORM", 390

    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureOrderLinesTranslations", _
        "Translation seeds ensured for sfrmOrderLines."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureOrderLinesTranslations", Err
End Sub

Public Sub EnsureOrderLinesTags()
    On Error GoTo ErrorHandler

    Dim metadataItems As Collection
    Dim metadata As Variant
    Dim controlTagMap As Object
    Dim ControlName As String
    Dim currentTag As String
    Dim updatedTag As String
    Dim updatedCount As Long

    If Not FormObjectExists(FORM_ORDER_LINES_SUBFORM) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".EnsureOrderLinesTags", _
            "Form '" & FORM_ORDER_LINES_SUBFORM & "' is not available in the current Access project. Tag ensure skipped."
        Exit Sub
    End If

    Set metadataItems = modFwComposerService.GetFormControlMetadata(FORM_ORDER_LINES_SUBFORM, True)
    Set controlTagMap = CreateObject("Scripting.Dictionary")
    controlTagMap.CompareMode = vbTextCompare

    For Each metadata In metadataItems
        ControlName = Trim$(modDaoHelper.NzString(metadata("control_name")))
        If LenB(ControlName) = 0 Then
            GoTo NextControl
        End If

        currentTag = modDaoHelper.NzString(metadata("current_tag"))
        updatedTag = modFwTranslationRuntime.SetTranslationKeyInTag( _
            currentTag, _
            BuildOrderLinesTranslationKey(ControlName))

        If StrComp(updatedTag, currentTag, vbBinaryCompare) <> 0 Then
            controlTagMap(ControlName) = updatedTag
        End If

NextControl:
    Next metadata

    If controlTagMap.count > 0 Then
        If Not modFwComposerService.SaveControlTagsToObject(modFwComposerService.OBJECT_TYPE_FORM, FORM_ORDER_LINES_SUBFORM, controlTagMap, updatedCount) Then
            Err.Raise vbObjectError + 6112, MODULE_NAME & ".EnsureOrderLinesTags", _
                "Failed to persist sfrmOrderLines tags."
        End If
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureOrderLinesTags", _
        "Translation tags ensured for sfrmOrderLines. updated_count=" & CStr(updatedCount) & "."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureOrderLinesTags", Err
End Sub

Private Function BuildTagGeneratorTranslationKey(ByVal ControlName As String) As String
    ControlName = UCase$(Trim$(modDaoHelper.NzString(ControlName)))
    BuildTagGeneratorTranslationKey = "FORM." & UCase$(FORM_TRANSLATION_TAG_GENERATOR) & "." & ControlName
End Function

Private Function BuildAddressCockpitTranslationKey(ByVal ControlName As String) As String
    ControlName = UCase$(Trim$(modDaoHelper.NzString(ControlName)))
    BuildAddressCockpitTranslationKey = "FORM." & UCase$(FORM_ADDRESS_COCKPIT) & "." & ControlName
End Function

Private Function BuildOrderLinesTranslationKey(ByVal ControlName As String) As String
    ControlName = UCase$(Trim$(modDaoHelper.NzString(ControlName)))
    BuildOrderLinesTranslationKey = "FORM." & UCase$(FORM_ORDER_LINES_SUBFORM) & "." & ControlName
End Function

Private Function FormObjectExists(ByVal FormName As String) As Boolean
    On Error GoTo SafeExit

    Dim accessObject As accessObject

    For Each accessObject In CurrentProject.AllForms
        If StrComp(accessObject.Name, FormName, vbTextCompare) = 0 Then
            FormObjectExists = True
            Exit Function
        End If
    Next accessObject

SafeExit:
End Function

Public Sub SeedShellTranslations()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = currentDb

    InsertShellTranslations db
    MsgBox "Shell-, Dashboard- und Navigations-Uebersetzungen wurden erfolgreich initialisiert.", vbInformation, MODULE_NAME
    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Initialisieren der Shell-Uebersetzungen: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Public Sub SeedVatCodeReference()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = currentDb

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
    Set db = currentDb

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
    Set db = currentDb

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
    Set db = currentDb

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
    Set db = currentDb

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
    ByVal languageCode As String, _
    ByVal translationKey As String, _
    ByVal TranslationValue As String, _
    ByVal isActive As Boolean, _
    Optional ByVal moduleCode As String = "", _
    Optional ByVal sortOrder As Long = 0)

    Dim sqlStatement As String

    languageCode = NormalizeSeedLanguageCode(languageCode)

    If TranslationSeedExists(db, languageCode, translationKey) Then
        Exit Sub
    End If

    sqlStatement = "INSERT INTO fw_translation " & _
                   "(language_code, translation_key, translation_value, is_active, module_code, sort_order, created_at, created_by, updated_at, updated_by) " & _
                   "VALUES (" & _
                   SqlText(languageCode) & ", " & _
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
    Set db = currentDb

    EnsureArticleTypeSeed db, "PRODUCT", "Product", "REF.ARTICLE_TYPE.PRODUCT", "Standard article or physical product.", 10, True
    EnsureArticleTypeSeed db, "SERVICE", "Service", "REF.ARTICLE_TYPE.SERVICE", "Service article without stock behavior.", 20, True
    EnsureArticleTypeSeed db, "SUBSCRIPTION", "Subscription", "REF.ARTICLE_TYPE.SUBSCRIPTION", "Recurring service or subscription article.", 30, True
    EnsureArticleTypeSeed db, "FEE", "Fee", "REF.ARTICLE_TYPE.FEE", "Fee article for fixed charges.", 40, True
    EnsureArticleTypeSeed db, "DISCOUNT", "Discount", "REF.ARTICLE_TYPE.DISCOUNT", "Discount article for explicit reductions.", 50, True
    EnsureArticleTypeSeed db, "WINE", "Wine", "REF.ARTICLE_TYPE.WINE", "Wine article prepared for future wine-specific extensions.", 60, True
    EnsureArticleTypeSeed db, "CUSTOM_SIZE", "Custom Size", "REF.ARTICLE_TYPE.CUSTOM_SIZE", "Article with future dimensional specialization.", 70, True
    EnsureArticleTypeSeed db, "APPAREL_SIZE", "Apparel Size", "REF.ARTICLE_TYPE.APPAREL_SIZE", "Article with future apparel size specialization.", 80, True

    EnsureTranslationSeed db, "de-CH", "REF.ARTICLE_TYPE.PRODUCT", "Produkt", "REF", 400
    EnsureTranslationSeed db, "en-US", "REF.ARTICLE_TYPE.PRODUCT", "Product", "REF", 400
    EnsureTranslationSeed db, "fr-CH", "REF.ARTICLE_TYPE.PRODUCT", "Produit", "REF", 400

    EnsureTranslationSeed db, "de-CH", "REF.ARTICLE_TYPE.SERVICE", "Dienstleistung", "REF", 401
    EnsureTranslationSeed db, "en-US", "REF.ARTICLE_TYPE.SERVICE", "Service", "REF", 401
    EnsureTranslationSeed db, "fr-CH", "REF.ARTICLE_TYPE.SERVICE", "Service", "REF", 401

    EnsureTranslationSeed db, "de-CH", "REF.ARTICLE_TYPE.SUBSCRIPTION", "Abonnement", "REF", 402
    EnsureTranslationSeed db, "en-US", "REF.ARTICLE_TYPE.SUBSCRIPTION", "Subscription", "REF", 402
    EnsureTranslationSeed db, "fr-CH", "REF.ARTICLE_TYPE.SUBSCRIPTION", "Abonnement", "REF", 402

    EnsureTranslationSeed db, "de-CH", "REF.ARTICLE_TYPE.FEE", "Gebuehr", "REF", 403
    EnsureTranslationSeed db, "en-US", "REF.ARTICLE_TYPE.FEE", "Fee", "REF", 403
    EnsureTranslationSeed db, "fr-CH", "REF.ARTICLE_TYPE.FEE", "Frais", "REF", 403

    EnsureTranslationSeed db, "de-CH", "REF.ARTICLE_TYPE.DISCOUNT", "Rabatt", "REF", 404
    EnsureTranslationSeed db, "en-US", "REF.ARTICLE_TYPE.DISCOUNT", "Discount", "REF", 404
    EnsureTranslationSeed db, "fr-CH", "REF.ARTICLE_TYPE.DISCOUNT", "Remise", "REF", 404

    EnsureTranslationSeed db, "de-CH", "REF.ARTICLE_TYPE.WINE", "Wein", "REF", 405
    EnsureTranslationSeed db, "en-US", "REF.ARTICLE_TYPE.WINE", "Wine", "REF", 405
    EnsureTranslationSeed db, "fr-CH", "REF.ARTICLE_TYPE.WINE", "Vin", "REF", 405

    EnsureTranslationSeed db, "de-CH", "REF.ARTICLE_TYPE.CUSTOM_SIZE", "Massanfertigung", "REF", 406
    EnsureTranslationSeed db, "en-US", "REF.ARTICLE_TYPE.CUSTOM_SIZE", "Custom Size", "REF", 406
    EnsureTranslationSeed db, "fr-CH", "REF.ARTICLE_TYPE.CUSTOM_SIZE", "Sur mesure", "REF", 406

    EnsureTranslationSeed db, "de-CH", "REF.ARTICLE_TYPE.APPAREL_SIZE", "Kleidergroesse", "REF", 407
    EnsureTranslationSeed db, "en-US", "REF.ARTICLE_TYPE.APPAREL_SIZE", "Apparel Size", "REF", 407
    EnsureTranslationSeed db, "fr-CH", "REF.ARTICLE_TYPE.APPAREL_SIZE", "Taille de vetement", "REF", 407

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
    ByVal UnitCode As String, _
    ByVal translationKey As String, _
    ByVal sortOrder As Long, _
    ByVal isActive As Boolean)

    Dim sqlStatement As String

    sqlStatement = "INSERT INTO ref_unit " & _
                   "(unit_code, translation_key, sort_order, is_active, created_at, created_by, updated_at, updated_by) " & _
                   "VALUES (" & _
                   SqlText(UnitCode) & ", " & _
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
    ByVal DescriptionText As String, _
    ByVal sortOrder As Long, _
    ByVal isActive As Boolean)

    Dim sqlStatement As String

    sqlStatement = "INSERT INTO ref_article_type_code " & _
                   "(article_type_code, article_type_name, translation_key, description_text, sort_order, is_active, created_at, created_by, updated_at, updated_by) " & _
                   "VALUES (" & _
                   SqlText(articleTypeCode) & ", " & _
                   SqlText(articleTypeName) & ", " & _
                   SqlText(translationKey) & ", " & _
                   SqlNullableText(DescriptionText) & ", " & _
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
    ByVal DescriptionText As String, _
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
                    "description_text = " & SqlNullableText(DescriptionText) & ", " & _
                    "sort_order = " & CStr(sortOrder) & ", " & _
                    "is_active = " & IIf(isActive, "True", "False") & ", " & _
                    "updated_at = Now(), " & _
                    "updated_by = 'SYSTEM' " & _
                    "WHERE " & criteria & ";"
        db.Execute updateSql, dbFailOnError
    Else
        InsertArticleType db, articleTypeCode, articleTypeName, translationKey, DescriptionText, sortOrder, isActive
    End If
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, MODULE_NAME & ".EnsureArticleTypeSeed", Err.description
End Sub

Private Sub InsertVatCode( _
    ByVal db As DAO.Database, _
    ByVal VatCode As String, _
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
                   SqlText(VatCode) & ", " & _
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
    InsertTranslation db, "de-CH", "ADDRESS_TYPE.PRIVATE", "Privat", True
    InsertTranslation db, "DE-DE", "ADDRESS_TYPE.PRIVATE", "Privat", True
    InsertTranslation db, "fr-CH", "ADDRESS_TYPE.PRIVATE", "Prive", True
    InsertTranslation db, "it-CH", "ADDRESS_TYPE.PRIVATE", "Privato", True
    InsertTranslation db, "en-US", "ADDRESS_TYPE.PRIVATE", "Private", True

    InsertTranslation db, "de-CH", "ADDRESS_TYPE.COMPANY", "Firma", True
    InsertTranslation db, "DE-DE", "ADDRESS_TYPE.COMPANY", "Unternehmen", True
    InsertTranslation db, "fr-CH", "ADDRESS_TYPE.COMPANY", "Entreprise", True
    InsertTranslation db, "it-CH", "ADDRESS_TYPE.COMPANY", "Azienda", True
    InsertTranslation db, "en-US", "ADDRESS_TYPE.COMPANY", "Company", True

    InsertTranslation db, "de-CH", "ADDRESS_TYPE.CUSTOMER", "Kunde", True
    InsertTranslation db, "DE-DE", "ADDRESS_TYPE.CUSTOMER", "Kunde", True
    InsertTranslation db, "fr-CH", "ADDRESS_TYPE.CUSTOMER", "Client", True
    InsertTranslation db, "it-CH", "ADDRESS_TYPE.CUSTOMER", "Cliente", True
    InsertTranslation db, "en-US", "ADDRESS_TYPE.CUSTOMER", "Customer", True

    InsertTranslation db, "de-CH", "ADDRESS_TYPE.SUPPLIER", "Lieferant", True
    InsertTranslation db, "DE-DE", "ADDRESS_TYPE.SUPPLIER", "Lieferant", True
    InsertTranslation db, "fr-CH", "ADDRESS_TYPE.SUPPLIER", "Fournisseur", True
    InsertTranslation db, "it-CH", "ADDRESS_TYPE.SUPPLIER", "Fornitore", True
    InsertTranslation db, "en-US", "ADDRESS_TYPE.SUPPLIER", "Supplier", True

    InsertTranslation db, "de-CH", "ADDRESS_TYPE.PARTNER", "Partner", True
    InsertTranslation db, "DE-DE", "ADDRESS_TYPE.PARTNER", "Partner", True
    InsertTranslation db, "fr-CH", "ADDRESS_TYPE.PARTNER", "Partenaire", True
    InsertTranslation db, "it-CH", "ADDRESS_TYPE.PARTNER", "Partner", True
    InsertTranslation db, "en-US", "ADDRESS_TYPE.PARTNER", "Partner", True

    InsertTranslation db, "de-CH", "ADDRESS_TYPE.EMPLOYEE", "Mitarbeiter", True
    InsertTranslation db, "DE-DE", "ADDRESS_TYPE.EMPLOYEE", "Mitarbeiter", True
    InsertTranslation db, "fr-CH", "ADDRESS_TYPE.EMPLOYEE", "Employe", True
    InsertTranslation db, "it-CH", "ADDRESS_TYPE.EMPLOYEE", "Collaboratore", True
    InsertTranslation db, "en-US", "ADDRESS_TYPE.EMPLOYEE", "Employee", True

    InsertTranslation db, "de-CH", "ADDRESS_TYPE.OTHER", "Andere", True
    InsertTranslation db, "DE-DE", "ADDRESS_TYPE.OTHER", "Sonstige", True
    InsertTranslation db, "fr-CH", "ADDRESS_TYPE.OTHER", "Autre", True
    InsertTranslation db, "it-CH", "ADDRESS_TYPE.OTHER", "Altro", True
    InsertTranslation db, "en-US", "ADDRESS_TYPE.OTHER", "Other", True
End Sub

Private Sub InsertSalutationTranslations(ByVal db As DAO.Database)
    InsertTranslation db, "de-CH", "SALUTATION.MR", "Herr", True
    InsertTranslation db, "DE-DE", "SALUTATION.MR", "Herr", True
    InsertTranslation db, "fr-CH", "SALUTATION.MR", "Monsieur", True
    InsertTranslation db, "it-CH", "SALUTATION.MR", "Signor", True
    InsertTranslation db, "en-US", "SALUTATION.MR", "Mr.", True

    InsertTranslation db, "de-CH", "SALUTATION.MS", "Frau", True
    InsertTranslation db, "DE-DE", "SALUTATION.MS", "Frau", True
    InsertTranslation db, "fr-CH", "SALUTATION.MS", "Madame", True
    InsertTranslation db, "it-CH", "SALUTATION.MS", "Signora", True
    InsertTranslation db, "en-US", "SALUTATION.MS", "Ms.", True

    InsertTranslation db, "de-CH", "SALUTATION.COMPANY", "Firma", True
    InsertTranslation db, "DE-DE", "SALUTATION.COMPANY", "Unternehmen", True
    InsertTranslation db, "fr-CH", "SALUTATION.COMPANY", "Entreprise", True
    InsertTranslation db, "it-CH", "SALUTATION.COMPANY", "Azienda", True
    InsertTranslation db, "en-US", "SALUTATION.COMPANY", "Company", True

    InsertTranslation db, "de-CH", "SALUTATION.NEUTRAL", "Neutral", True
    InsertTranslation db, "DE-DE", "SALUTATION.NEUTRAL", "Neutral", True
    InsertTranslation db, "fr-CH", "SALUTATION.NEUTRAL", "Neutre", True
    InsertTranslation db, "it-CH", "SALUTATION.NEUTRAL", "Neutrale", True
    InsertTranslation db, "en-US", "SALUTATION.NEUTRAL", "Neutral", True
End Sub

Private Sub InsertAddressingModeTranslations(ByVal db As DAO.Database)
    InsertTranslation db, "de-CH", "ADDRESSING_MODE.FORMAL", "Formal", True
    InsertTranslation db, "DE-DE", "ADDRESSING_MODE.FORMAL", "Formal", True
    InsertTranslation db, "fr-CH", "ADDRESSING_MODE.FORMAL", "Formel", True
    InsertTranslation db, "it-CH", "ADDRESSING_MODE.FORMAL", "Formale", True
    InsertTranslation db, "en-US", "ADDRESSING_MODE.FORMAL", "Formal", True

    InsertTranslation db, "de-CH", "ADDRESSING_MODE.INFORMAL", "Informell", True
    InsertTranslation db, "DE-DE", "ADDRESSING_MODE.INFORMAL", "Informell", True
    InsertTranslation db, "fr-CH", "ADDRESSING_MODE.INFORMAL", "Informel", True
    InsertTranslation db, "it-CH", "ADDRESSING_MODE.INFORMAL", "Informale", True
    InsertTranslation db, "en-US", "ADDRESSING_MODE.INFORMAL", "Informal", True
End Sub

Private Sub InsertContactTypeTranslations(ByVal db As DAO.Database)
    InsertTranslation db, "de-CH", "CONTACT_TYPE.EMAIL", "E-Mail", True
    InsertTranslation db, "DE-DE", "CONTACT_TYPE.EMAIL", "E-Mail", True
    InsertTranslation db, "fr-CH", "CONTACT_TYPE.EMAIL", "E-mail", True
    InsertTranslation db, "it-CH", "CONTACT_TYPE.EMAIL", "E-mail", True
    InsertTranslation db, "en-US", "CONTACT_TYPE.EMAIL", "E-Mail", True

    InsertTranslation db, "de-CH", "CONTACT_TYPE.PHONE", "Telefon", True
    InsertTranslation db, "DE-DE", "CONTACT_TYPE.PHONE", "Telefon", True
    InsertTranslation db, "fr-CH", "CONTACT_TYPE.PHONE", "Telephone", True
    InsertTranslation db, "it-CH", "CONTACT_TYPE.PHONE", "Telefono", True
    InsertTranslation db, "en-US", "CONTACT_TYPE.PHONE", "Phone", True

    InsertTranslation db, "de-CH", "CONTACT_TYPE.MOBILE", "Mobil", True
    InsertTranslation db, "DE-DE", "CONTACT_TYPE.MOBILE", "Mobil", True
    InsertTranslation db, "fr-CH", "CONTACT_TYPE.MOBILE", "Mobile", True
    InsertTranslation db, "it-CH", "CONTACT_TYPE.MOBILE", "Mobile", True
    InsertTranslation db, "en-US", "CONTACT_TYPE.MOBILE", "Mobile", True

    InsertTranslation db, "de-CH", "CONTACT_TYPE.WEBSITE", "Webseite", True
    InsertTranslation db, "DE-DE", "CONTACT_TYPE.WEBSITE", "Webseite", True
    InsertTranslation db, "fr-CH", "CONTACT_TYPE.WEBSITE", "Site web", True
    InsertTranslation db, "it-CH", "CONTACT_TYPE.WEBSITE", "Sito web", True
    InsertTranslation db, "en-US", "CONTACT_TYPE.WEBSITE", "Website", True

    InsertTranslation db, "de-CH", "CONTACT_TYPE.FAX", "Fax", True
    InsertTranslation db, "DE-DE", "CONTACT_TYPE.FAX", "Fax", True
    InsertTranslation db, "fr-CH", "CONTACT_TYPE.FAX", "Fax", True
    InsertTranslation db, "it-CH", "CONTACT_TYPE.FAX", "Fax", True
    InsertTranslation db, "en-US", "CONTACT_TYPE.FAX", "Fax", True

    InsertTranslation db, "de-CH", "CONTACT_TYPE.OTHER", "Sonstige", True
    InsertTranslation db, "DE-DE", "CONTACT_TYPE.OTHER", "Sonstige", True
    InsertTranslation db, "fr-CH", "CONTACT_TYPE.OTHER", "Autre", True
    InsertTranslation db, "it-CH", "CONTACT_TYPE.OTHER", "Altro", True
    InsertTranslation db, "en-US", "CONTACT_TYPE.OTHER", "Other", True
End Sub

Private Sub InsertUnitTranslations(ByVal db As DAO.Database)
    InsertTranslation db, "de-CH", "UNIT.PCS", "Stk", True
    InsertTranslation db, "DE-DE", "UNIT.PCS", "Stk", True
    InsertTranslation db, "fr-CH", "UNIT.PCS", "pcs", True
    InsertTranslation db, "it-CH", "UNIT.PCS", "pz", True
    InsertTranslation db, "en-US", "UNIT.PCS", "pcs", True

    InsertTranslation db, "de-CH", "UNIT.HOUR", "Stunde", True
    InsertTranslation db, "DE-DE", "UNIT.HOUR", "Stunde", True
    InsertTranslation db, "fr-CH", "UNIT.HOUR", "Heure", True
    InsertTranslation db, "it-CH", "UNIT.HOUR", "Ora", True
    InsertTranslation db, "en-US", "UNIT.HOUR", "Hour", True

    InsertTranslation db, "de-CH", "UNIT.KG", "Kilogramm", True
    InsertTranslation db, "DE-DE", "UNIT.KG", "Kilogramm", True
    InsertTranslation db, "fr-CH", "UNIT.KG", "Kilogramme", True
    InsertTranslation db, "it-CH", "UNIT.KG", "Chilogrammo", True
    InsertTranslation db, "en-US", "UNIT.KG", "Kilogram", True

    InsertTranslation db, "de-CH", "UNIT.METER", "Meter", True
    InsertTranslation db, "DE-DE", "UNIT.METER", "Meter", True
    InsertTranslation db, "fr-CH", "UNIT.METER", "Metre", True
    InsertTranslation db, "it-CH", "UNIT.METER", "Metro", True
    InsertTranslation db, "en-US", "UNIT.METER", "Meter", True

    InsertTranslation db, "de-CH", "UNIT.LITER", "Liter", True
    InsertTranslation db, "DE-DE", "UNIT.LITER", "Liter", True
    InsertTranslation db, "fr-CH", "UNIT.LITER", "Litre", True
    InsertTranslation db, "it-CH", "UNIT.LITER", "Litro", True
    InsertTranslation db, "en-US", "UNIT.LITER", "Liter", True

    InsertTranslation db, "de-CH", "UNIT.PACKAGE", "Paket", True
    InsertTranslation db, "DE-DE", "UNIT.PACKAGE", "Paket", True
    InsertTranslation db, "fr-CH", "UNIT.PACKAGE", "Colis", True
    InsertTranslation db, "it-CH", "UNIT.PACKAGE", "Pacco", True
    InsertTranslation db, "en-US", "UNIT.PACKAGE", "Package", True
End Sub

Private Sub InsertVatCodeTranslations(ByVal db As DAO.Database)
    InsertTranslation db, "de-CH", "VAT.CH.STANDARD", "Normalsatz", True
    InsertTranslation db, "DE-DE", "VAT.CH.STANDARD", "Standardsatz", True
    InsertTranslation db, "fr-CH", "VAT.CH.STANDARD", "Taux normal", True
    InsertTranslation db, "it-CH", "VAT.CH.STANDARD", "Aliquota normale", True
    InsertTranslation db, "en-US", "VAT.CH.STANDARD", "Standard rate", True

    InsertTranslation db, "de-CH", "VAT.CH.REDUCED", "Reduzierter Satz", True
    InsertTranslation db, "DE-DE", "VAT.CH.REDUCED", "Ermaessigter Satz", True
    InsertTranslation db, "fr-CH", "VAT.CH.REDUCED", "Taux reduit", True
    InsertTranslation db, "it-CH", "VAT.CH.REDUCED", "Aliquota ridotta", True
    InsertTranslation db, "en-US", "VAT.CH.REDUCED", "Reduced rate", True

    InsertTranslation db, "de-CH", "VAT.CH.SPECIAL", "Sondersatz", True
    InsertTranslation db, "DE-DE", "VAT.CH.SPECIAL", "Sondersatz", True
    InsertTranslation db, "fr-CH", "VAT.CH.SPECIAL", "Taux special", True
    InsertTranslation db, "it-CH", "VAT.CH.SPECIAL", "Aliquota speciale", True
    InsertTranslation db, "en-US", "VAT.CH.SPECIAL", "Special rate", True

    InsertTranslation db, "de-CH", "VAT.CH.ZERO", "Nullsatz", True
    InsertTranslation db, "DE-DE", "VAT.CH.ZERO", "Nullsatz", True
    InsertTranslation db, "fr-CH", "VAT.CH.ZERO", "Taux zero", True
    InsertTranslation db, "it-CH", "VAT.CH.ZERO", "Aliquota zero", True
    InsertTranslation db, "en-US", "VAT.CH.ZERO", "Zero rate", True
End Sub

Private Sub InsertShellTranslations(ByVal db As DAO.Database)
    EnsureTranslationSeed db, "de-CH", "NAV.GROUP.ADDRESSES", "Adressen", "NAVIGATION", 10
    EnsureTranslationSeed db, "en-US", "NAV.GROUP.ADDRESSES", "Addresses", "NAVIGATION", 10
    EnsureTranslationSeed db, "de-CH", "NAV.ADDRESS_LIST", "Adressliste", "NAVIGATION", 20
    EnsureTranslationSeed db, "en-US", "NAV.ADDRESS_LIST", "Address list", "NAVIGATION", 20
    EnsureTranslationSeed db, "de-CH", "NAV.NEW_ADDRESS", "Neue Adresse", "NAVIGATION", 30
    EnsureTranslationSeed db, "en-US", "NAV.NEW_ADDRESS", "New address", "NAVIGATION", 30
    EnsureTranslationSeed db, "fr-CH", "NAV.NEW_ADDRESS", "Nouvelle adresse", "NAVIGATION", 30
    EnsureTranslationSeed db, "de-CH", "NAV.GROUP.DOCUMENTS", "Dokumente", "NAVIGATION", 40
    EnsureTranslationSeed db, "en-US", "NAV.GROUP.DOCUMENTS", "Documents", "NAVIGATION", 40
    EnsureTranslationSeed db, "de-CH", "NAV.DOCUMENT_PREVIEW", "Dokumentvorschau", "NAVIGATION", 50
    EnsureTranslationSeed db, "en-US", "NAV.DOCUMENT_PREVIEW", "Document preview", "NAVIGATION", 50
    EnsureTranslationSeed db, "de-CH", "NAV.GROUP.FRAMEWORK", "Framework", "NAVIGATION", 60
    EnsureTranslationSeed db, "en-US", "NAV.GROUP.FRAMEWORK", "Framework", "NAVIGATION", 60
    EnsureTranslationSeed db, "de-CH", "NAV.TRANSLATIONS", "Uebersetzungen", "NAVIGATION", 70
    EnsureTranslationSeed db, "en-US", "NAV.TRANSLATIONS", "Translations", "NAVIGATION", 70
    EnsureTranslationSeed db, "de-CH", "NAV.FW_TRANSLATIONS", "Uebersetzungen pflegen", "NAVIGATION", 80
    EnsureTranslationSeed db, "en-US", "NAV.FW_TRANSLATIONS", "Maintain translations", "NAVIGATION", 80
    EnsureTranslationSeed db, "de-CH", "NAV.LOCALISATION", "Lokalisierung", "NAVIGATION", 90
    EnsureTranslationSeed db, "en-US", "NAV.LOCALISATION", "Localisation", "NAVIGATION", 90
    EnsureTranslationSeed db, "de-CH", "NAV.TAGS", "Tags", "NAVIGATION", 100
    EnsureTranslationSeed db, "en-US", "NAV.TAGS", "Tags", "NAVIGATION", 100
    EnsureTranslationSeed db, "de-CH", "NAV.TAG_HELP", "Tag-Hilfe", "NAVIGATION", 110
    EnsureTranslationSeed db, "en-US", "NAV.TAG_HELP", "Tag help", "NAVIGATION", 110
    EnsureTranslationSeed db, "de-CH", "NAV.GROUP.ORDERS", "Bestellungen", "NAVIGATION", 120
    EnsureTranslationSeed db, "en-US", "NAV.GROUP.ORDERS", "Orders", "NAVIGATION", 120
    EnsureTranslationSeed db, "de-CH", "NAV.GROUP.FINANCE", "Finanzen", "NAVIGATION", 130
    EnsureTranslationSeed db, "en-US", "NAV.GROUP.FINANCE", "Finance", "NAVIGATION", 130
    EnsureTranslationSeed db, "de-CH", "NAV.GROUP.REPORTING", "Auswertungen", "NAVIGATION", 140
    EnsureTranslationSeed db, "en-US", "NAV.GROUP.REPORTING", "Reports", "NAVIGATION", 140
    EnsureTranslationSeed db, "de-CH", "NAV.GROUP.TENANT", "Mandant", "NAVIGATION", 150
    EnsureTranslationSeed db, "en-US", "NAV.GROUP.TENANT", "Tenant", "NAVIGATION", 150
    EnsureTranslationSeed db, "de-CH", "NAV.ARTICLE_GROUPS", "Artikelgruppen", "NAVIGATION", 155
    EnsureTranslationSeed db, "en-US", "NAV.ARTICLE_GROUPS", "Article groups", "NAVIGATION", 155
    EnsureTranslationSeed db, "de-CH", "NAV.NEW_ARTICLE_GROUP", "Neue Artikelgruppe", "NAVIGATION", 156
    EnsureTranslationSeed db, "en-US", "NAV.NEW_ARTICLE_GROUP", "New article group", "NAVIGATION", 156
    EnsureTranslationSeed db, "de-CH", "NAV.ARTICLES", "Artikel", "NAVIGATION", 157
    EnsureTranslationSeed db, "en-US", "NAV.ARTICLES", "Articles", "NAVIGATION", 157
    EnsureTranslationSeed db, "de-CH", "NAV.GROUP.SYSTEM", "System", "NAVIGATION", 160
    EnsureTranslationSeed db, "en-US", "NAV.GROUP.SYSTEM", "System", "NAVIGATION", 160
    EnsureTranslationSeed db, "de-CH", "NAV.FW_NAVIGATION_ADMIN", "Navigation verwalten", "NAVIGATION", 170
    EnsureTranslationSeed db, "en-US", "NAV.FW_NAVIGATION_ADMIN", "Manage navigation", "NAVIGATION", 170

    EnsureTranslationSeed db, "de-CH", "COMMON.NEW", "Neu", "COMMON", 171
    EnsureTranslationSeed db, "en-US", "COMMON.NEW", "New", "COMMON", 171
    EnsureTranslationSeed db, "fr-CH", "COMMON.NEW", "Nouveau", "COMMON", 171
    EnsureTranslationSeed db, "de-CH", "COMMON.EDIT", "Bearbeiten", "COMMON", 172
    EnsureTranslationSeed db, "en-US", "COMMON.EDIT", "Edit", "COMMON", 172
    EnsureTranslationSeed db, "fr-CH", "COMMON.EDIT", "Modifier", "COMMON", 172
    EnsureTranslationSeed db, "de-CH", "COMMON.REFRESH", "Aktualisieren", "COMMON", 173
    EnsureTranslationSeed db, "en-US", "COMMON.REFRESH", "Refresh", "COMMON", 173
    EnsureTranslationSeed db, "fr-CH", "COMMON.REFRESH", "Actualiser", "COMMON", 173
    EnsureTranslationSeed db, "de-CH", "COMMON.SEARCH", "Suche", "COMMON", 174
    EnsureTranslationSeed db, "en-US", "COMMON.SEARCH", "Search", "COMMON", 174
    EnsureTranslationSeed db, "fr-CH", "COMMON.SEARCH", "Recherche", "COMMON", 174
    EnsureTranslationSeed db, "de-CH", "COMMON.CLEAR_SEARCH", "Leeren", "COMMON", 175
    EnsureTranslationSeed db, "en-US", "COMMON.CLEAR_SEARCH", "Clear", "COMMON", 175
    EnsureTranslationSeed db, "fr-CH", "COMMON.CLEAR_SEARCH", "Effacer", "COMMON", 175
    EnsureTranslationSeed db, "de-CH", "COMMON.HOME", "Home", "COMMON", 176
    EnsureTranslationSeed db, "en-US", "COMMON.HOME", "Home", "COMMON", 176
    EnsureTranslationSeed db, "fr-CH", "COMMON.HOME", "Accueil", "COMMON", 176
    EnsureTranslationSeed db, "de-CH", "COMMON.BACK", "Zurueck", "COMMON", 177
    EnsureTranslationSeed db, "en-US", "COMMON.BACK", "Back", "COMMON", 177
    EnsureTranslationSeed db, "fr-CH", "COMMON.BACK", "Retour", "COMMON", 177

    EnsureTranslationSeed db, "de-CH", "FORM.FRMAPPSHELL.APP_TITLE", "EASIS v4", "FORM", 10
    EnsureTranslationSeed db, "en-US", "FORM.FRMAPPSHELL.APP_TITLE", "EASIS v4", "FORM", 10
    EnsureTranslationSeed db, "de-CH", "FORM.FRMAPPSHELL.APP_SUBTITLE", "Access Framework", "FORM", 20
    EnsureTranslationSeed db, "en-US", "FORM.FRMAPPSHELL.APP_SUBTITLE", "Access Framework", "FORM", 20
    EnsureTranslationSeed db, "de-CH", "FORM.FRMAPPSHELL.USER", "Benutzer", "FORM", 30
    EnsureTranslationSeed db, "en-US", "FORM.FRMAPPSHELL.USER", "User", "FORM", 30
    EnsureTranslationSeed db, "de-CH", "FORM.FRMAPPSHELL.TENANT", "Mandant", "FORM", 40
    EnsureTranslationSeed db, "en-US", "FORM.FRMAPPSHELL.TENANT", "Tenant", "FORM", 40
    EnsureTranslationSeed db, "de-CH", "FORM.FRMAPPSHELL.ROLE", "Rolle", "FORM", 50
    EnsureTranslationSeed db, "en-US", "FORM.FRMAPPSHELL.ROLE", "Role", "FORM", 50
    EnsureTranslationSeed db, "de-CH", "FORM.FRMAPPSHELL.ENVIRONMENT", "Umgebung", "FORM", 60
    EnsureTranslationSeed db, "en-US", "FORM.FRMAPPSHELL.ENVIRONMENT", "Environment", "FORM", 60
    EnsureTranslationSeed db, "de-CH", "FORM.FRMAPPSHELL.BACKEND", "Backend", "FORM", 70
    EnsureTranslationSeed db, "en-US", "FORM.FRMAPPSHELL.BACKEND", "Backend", "FORM", 70

    EnsureTranslationSeed db, "de-CH", "FORM.FRMAPPDASHBOARD.TENANT", "Mandant", "FORM", 80
    EnsureTranslationSeed db, "en-US", "FORM.FRMAPPDASHBOARD.TENANT", "Tenant", "FORM", 80
    EnsureTranslationSeed db, "de-CH", "FORM.FRMAPPDASHBOARD.USER", "Benutzer", "FORM", 90
    EnsureTranslationSeed db, "en-US", "FORM.FRMAPPDASHBOARD.USER", "User", "FORM", 90
    EnsureTranslationSeed db, "de-CH", "FORM.FRMADDRESSDETAIL.TITLE_NEW", "Neue Adresse", "FORM", 91
    EnsureTranslationSeed db, "en-US", "FORM.FRMADDRESSDETAIL.TITLE_NEW", "New address", "FORM", 91
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMADDRESSDETAIL.TITLE_NEW", "Nouvelle adresse", "FORM", 91
    EnsureTranslationSeed db, "de-CH", "FORM.FRMADDRESSDETAIL.TITLE_EDIT", "Adresse bearbeiten", "FORM", 92
    EnsureTranslationSeed db, "en-US", "FORM.FRMADDRESSDETAIL.TITLE_EDIT", "Edit address", "FORM", 92
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMADDRESSDETAIL.TITLE_EDIT", "Modifier l'adresse", "FORM", 92
    EnsureTranslationSeed db, "de-CH", "FORM.FRMAPPDASHBOARD.BACKEND", "Backend", "FORM", 100
    EnsureTranslationSeed db, "en-US", "FORM.FRMAPPDASHBOARD.BACKEND", "Backend", "FORM", 100
    EnsureTranslationSeed db, "de-CH", "FORM.FRMAPPDASHBOARD.FRAMEWORK", "Framework", "FORM", 110
    EnsureTranslationSeed db, "en-US", "FORM.FRMAPPDASHBOARD.FRAMEWORK", "Framework", "FORM", 110
    EnsureTranslationSeed db, "de-CH", "FORM.FRMAPPDASHBOARD.STATUS", "Status", "FORM", 120
    EnsureTranslationSeed db, "en-US", "FORM.FRMAPPDASHBOARD.STATUS", "Status", "FORM", 120

    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWNAVIGATIONADMIN.FORM_TITLE", "Navigation verwalten", "FORM", 130
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWNAVIGATIONADMIN.FORM_TITLE", "Manage navigation", "FORM", 130
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWNAVIGATIONADMIN.SAVE", "Speichern", "FORM", 140
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWNAVIGATIONADMIN.SAVE", "Save", "FORM", 140
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWNAVIGATIONADMIN.REFRESH", "Aktualisieren", "FORM", 150
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWNAVIGATIONADMIN.REFRESH", "Refresh", "FORM", 150
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWNAVIGATIONADMIN.NEW_GROUP", "Neue Gruppe", "FORM", 160
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWNAVIGATIONADMIN.NEW_GROUP", "New group", "FORM", 160
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWNAVIGATIONADMIN.NEW_ITEM", "Neuer Eintrag", "FORM", 170
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWNAVIGATIONADMIN.NEW_ITEM", "New item", "FORM", 170
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWNAVIGATIONADMIN.DEACTIVATE", "Deaktivieren", "FORM", 180
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWNAVIGATIONADMIN.DEACTIVATE", "Deactivate", "FORM", 180
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWNAVIGATIONADMIN.HIDE", "Ausblenden", "FORM", 190
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWNAVIGATIONADMIN.HIDE", "Hide", "FORM", 190
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWNAVIGATIONADMIN.SHOW", "Einblenden", "FORM", 200
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWNAVIGATIONADMIN.SHOW", "Show", "FORM", 200

    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEGROUPLIST.FORM_TITLE", "Artikelgruppen", "FORM", 205
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEGROUPLIST.FORM_TITLE", "Article groups", "FORM", 205
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEGROUPLIST.SEARCH", "Suche", "FORM", 206
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEGROUPLIST.SEARCH", "Search", "FORM", 206
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEGROUPLIST.EDIT", "Bearbeiten", "FORM", 207
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEGROUPLIST.EDIT", "Edit", "FORM", 207
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEGROUPLIST.REFRESH", "Aktualisieren", "FORM", 208
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEGROUPLIST.REFRESH", "Refresh", "FORM", 208

    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLELIST.FORM_TITLE", "Artikel", "FORM", 2081
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLELIST.FORM_TITLE", "Articles", "FORM", 2081
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLELIST.SEARCH", "Suche", "FORM", 2082
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLELIST.SEARCH", "Search", "FORM", 2082
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLELIST.REFRESH", "Aktualisieren", "FORM", 2083
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLELIST.REFRESH", "Refresh", "FORM", 2083
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLELIST.ARTICLE_NO", "Artikel-Nr.", "FORM", 2084
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLELIST.ARTICLE_NO", "Article no.", "FORM", 2084
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLELIST.ARTICLE_NAME", "Artikelname", "FORM", 2085
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLELIST.ARTICLE_NAME", "Article name", "FORM", 2085
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLELIST.PRODUCT_GROUP", "Artikelgruppe", "FORM", 2086
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLELIST.PRODUCT_GROUP", "Product group", "FORM", 2086
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLELIST.UNIT_CODE", "Einheit", "FORM", 2087
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLELIST.UNIT_CODE", "Unit", "FORM", 2087
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLELIST.VAT_CODE", "MWST-Code", "FORM", 2088
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLELIST.VAT_CODE", "VAT code", "FORM", 2088
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLELIST.SALES_PRICE", "Verkaufspreis", "FORM", 2089
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLELIST.SALES_PRICE", "Sales price", "FORM", 2089
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLELIST.IS_ACTIVE", "Aktiv", "FORM", 2090
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLELIST.IS_ACTIVE", "Active", "FORM", 2090

    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEDETAIL.FORM_TITLE", "Artikel", "FORM", 2091
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEDETAIL.FORM_TITLE", "Article", "FORM", 2091
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEDETAIL.ARTICLE_NO", "Artikel-Nr.", "FORM", 2092
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEDETAIL.ARTICLE_NO", "Article no.", "FORM", 2092
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEDETAIL.ARTICLE_NAME", "Artikelname", "FORM", 2093
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEDETAIL.ARTICLE_NAME", "Article name", "FORM", 2093
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEDETAIL.PRODUCT_GROUP", "Artikelgruppe", "FORM", 2094
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEDETAIL.PRODUCT_GROUP", "Product group", "FORM", 2094
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEDETAIL.ARTICLE_TYPE_CODE", "Artikeltyp", "FORM", 2095
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEDETAIL.ARTICLE_TYPE_CODE", "Article type", "FORM", 2095
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEDETAIL.UNIT_CODE", "Einheit", "FORM", 2096
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEDETAIL.UNIT_CODE", "Unit", "FORM", 2096
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEDETAIL.VAT_CODE", "MWST-Code", "FORM", 2097
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEDETAIL.VAT_CODE", "VAT code", "FORM", 2097
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEDETAIL.PURCHASE_PRICE", "Einkaufspreis", "FORM", 2098
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEDETAIL.PURCHASE_PRICE", "Purchase price", "FORM", 2098
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEDETAIL.SALES_PRICE", "Verkaufspreis", "FORM", 2099
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEDETAIL.SALES_PRICE", "Sales price", "FORM", 2099
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEDETAIL.BARCODE", "Barcode", "FORM", 2100
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEDETAIL.BARCODE", "Barcode", "FORM", 2100
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEDETAIL.DESCRIPTION_TEXT", "Beschreibung", "FORM", 2101
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEDETAIL.DESCRIPTION_TEXT", "Description", "FORM", 2101
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEDETAIL.IS_ACTIVE", "Aktiv", "FORM", 2102
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEDETAIL.IS_ACTIVE", "Active", "FORM", 2102
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEDETAIL.SAVE", "Speichern", "FORM", 2103
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEDETAIL.SAVE", "Save", "FORM", 2103
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEDETAIL.CANCEL", "Abbrechen", "FORM", 2104
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEDETAIL.CANCEL", "Cancel", "FORM", 2104

    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEGROUPDETAIL.FORM_TITLE", "Artikelgruppe", "FORM", 209
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEGROUPDETAIL.FORM_TITLE", "Article group", "FORM", 209
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEGROUPDETAIL.PRODUCT_GROUP_CODE", "Artikelgruppen-Code", "FORM", 210
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEGROUPDETAIL.PRODUCT_GROUP_CODE", "Article group code", "FORM", 210
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEGROUPDETAIL.PRODUCT_GROUP_NAME", "Artikelgruppen-Name", "FORM", 211
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEGROUPDETAIL.PRODUCT_GROUP_NAME", "Article group name", "FORM", 211
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEGROUPDETAIL.DESCRIPTION_TEXT", "Beschreibung", "FORM", 212
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEGROUPDETAIL.DESCRIPTION_TEXT", "Description", "FORM", 212
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEGROUPDETAIL.ARTICLE_GROUP_CODE", "Artikelgruppen-Code", "FORM", 210
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEGROUPDETAIL.ARTICLE_GROUP_CODE", "Article group code", "FORM", 210
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEGROUPDETAIL.ARTICLE_GROUP_NAME", "Artikelgruppen-Name", "FORM", 211
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEGROUPDETAIL.ARTICLE_GROUP_NAME", "Article group name", "FORM", 211
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEGROUPDETAIL.DESCRIPTION", "Beschreibung", "FORM", 212
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEGROUPDETAIL.DESCRIPTION", "Description", "FORM", 212
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEGROUPDETAIL.IS_ACTIVE", "Aktiv", "FORM", 213
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEGROUPDETAIL.IS_ACTIVE", "Active", "FORM", 213
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEGROUPDETAIL.SORT_ORDER", "Sortierung", "FORM", 214
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEGROUPDETAIL.SORT_ORDER", "Sort order", "FORM", 214
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEGROUPDETAIL.SAVE", "Speichern", "FORM", 215
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEGROUPDETAIL.SAVE", "Save", "FORM", 215
    EnsureTranslationSeed db, "de-CH", "FORM.FRMARTICLEGROUPDETAIL.CANCEL", "Abbrechen", "FORM", 216
    EnsureTranslationSeed db, "en-US", "FORM.FRMARTICLEGROUPDETAIL.CANCEL", "Cancel", "FORM", 216

    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_GROUP_CODE_REQUIRED", "Artikelgruppen-Code ist erforderlich.", "MSG", 217
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_GROUP_CODE_REQUIRED", "Article group code is required.", "MSG", 217
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_GROUP_NAME_REQUIRED", "Artikelgruppen-Name ist erforderlich.", "MSG", 218
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_GROUP_NAME_REQUIRED", "Article group name is required.", "MSG", 218
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_GROUP_DUPLICATE_CODE", "Artikelgruppen-Code existiert bereits.", "MSG", 219
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_GROUP_DUPLICATE_CODE", "Article group code already exists.", "MSG", 219
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_GROUP_SELECT_FIRST", "Bitte zuerst eine Artikelgruppe auswaehlen.", "MSG", 221
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_GROUP_SELECT_FIRST", "Please select an article group first.", "MSG", 221
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_GROUP_DISCARD_CHANGES", "Ungespeicherte Aenderungen verwerfen?", "MSG", 222
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_GROUP_DISCARD_CHANGES", "Discard unsaved changes?", "MSG", 222
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_GROUP_CANCEL_CONFIRM", "Aenderungen verwerfen?", "MSG", 223
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_GROUP_CANCEL_CONFIRM", "Discard changes?", "MSG", 223
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_GROUP_SAVE_OR_CANCEL_FIRST", "Bitte zuerst speichern oder abbrechen.", "MSG", 224
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_GROUP_SAVE_OR_CANCEL_FIRST", "Please save or cancel first.", "MSG", 224
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_GROUP_LIST_LOAD_ERROR", "Fehler beim Laden der Artikelgruppenliste.", "MSG", 225
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_GROUP_LIST_LOAD_ERROR", "Error loading the article group list.", "MSG", 225
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_GROUP_DETAIL_LOAD_ERROR", "Fehler beim Laden der Artikelgruppendetails.", "MSG", 226
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_GROUP_DETAIL_LOAD_ERROR", "Error loading the article group details.", "MSG", 226
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_GROUP_SAVE_ERROR", "Fehler beim Speichern der Artikelgruppe.", "MSG", 227
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_GROUP_SAVE_ERROR", "Error saving the article group.", "MSG", 227
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_LIST_LOAD_ERROR", "Fehler beim Laden der Artikelliste.", "MSG", 228
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_LIST_LOAD_ERROR", "Error loading the article list.", "MSG", 228
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_NO_REQUIRED", "Artikel-Nr. ist erforderlich.", "MSG", 229
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_NO_REQUIRED", "Article no. is required.", "MSG", 229
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_NAME_REQUIRED", "Artikelname ist erforderlich.", "MSG", 230
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_NAME_REQUIRED", "Article name is required.", "MSG", 230
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_GROUP_REQUIRED", "Artikelgruppe ist erforderlich.", "MSG", 231
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_GROUP_REQUIRED", "Product group is required.", "MSG", 231
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_UNIT_REQUIRED", "Einheit ist erforderlich.", "MSG", 232
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_UNIT_REQUIRED", "Unit is required.", "MSG", 232
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_VAT_REQUIRED", "MWST-Code ist erforderlich.", "MSG", 233
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_VAT_REQUIRED", "VAT code is required.", "MSG", 233
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_SALES_PRICE_REQUIRED", "Verkaufspreis ist erforderlich.", "MSG", 234
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_SALES_PRICE_REQUIRED", "Sales price is required.", "MSG", 234
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_DUPLICATE_NO", "Artikel-Nr. existiert bereits.", "MSG", 235
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_DUPLICATE_NO", "Article no. already exists.", "MSG", 235
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_SAVE_ERROR", "Fehler beim Speichern des Artikels.", "MSG", 236
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_SAVE_ERROR", "Error saving the article.", "MSG", 236
    EnsureTranslationSeed db, "de-CH", "MSG.ARTICLE_CANCEL_CONFIRM", "Aenderungen verwerfen?", "MSG", 237
    EnsureTranslationSeed db, "en-US", "MSG.ARTICLE_CANCEL_CONFIRM", "Discard changes?", "MSG", 237

    EnsureTranslationSeed db, "de-CH", "STATUS.READY", "Bereit", "STATUS", 238
    EnsureTranslationSeed db, "en-US", "STATUS.READY", "Ready", "STATUS", 238
    EnsureTranslationSeed db, "fr-CH", "STATUS.READY", "Pret", "STATUS", 238

    EnsureTranslationSeed db, "de-CH", "COMMON.ALL", "Alle", "COMMON", 239
    EnsureTranslationSeed db, "en-US", "COMMON.ALL", "All", "COMMON", 239
    EnsureTranslationSeed db, "fr-CH", "COMMON.ALL", "Tous", "COMMON", 239

    EnsureTranslationSeed db, "de-CH", "NAV.FW_TRANSLATION_AUDIT", "Uebersetzungs-Audit", "NAVIGATION", 240
    EnsureTranslationSeed db, "en-US", "NAV.FW_TRANSLATION_AUDIT", "Translation audit", "NAVIGATION", 240
    EnsureTranslationSeed db, "fr-CH", "NAV.FW_TRANSLATION_AUDIT", "Audit des traductions", "NAVIGATION", 240
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONAUDIT.FORM_TITLE", "Uebersetzungs-Audit", "FORM", 241
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONAUDIT.FORM_TITLE", "Translation audit", "FORM", 241
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONAUDIT.FORM_TITLE", "Audit des traductions", "FORM", 241
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONAUDIT.SCOPE_CODE", "Bereich", "FORM", 242
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONAUDIT.SCOPE_CODE", "Scope", "FORM", 242
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONAUDIT.SCOPE_CODE", "Portee", "FORM", 242
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONAUDIT.LANGUAGE_CODE", "Sprache", "FORM", 243
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONAUDIT.LANGUAGE_CODE", "Language", "FORM", 243
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONAUDIT.LANGUAGE_CODE", "Langue", "FORM", 243
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONAUDIT.AUDIT_STATUS", "Audit-Status", "FORM", 244
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONAUDIT.AUDIT_STATUS", "Audit status", "FORM", 244
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONAUDIT.AUDIT_STATUS", "Statut d'audit", "FORM", 244
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONAUDIT.SEARCH", "Suche", "FORM", 245
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONAUDIT.SEARCH", "Search", "FORM", 245
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONAUDIT.SEARCH", "Recherche", "FORM", 245
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONAUDIT.COVERAGE_SUMMARY", "Abdeckungsuebersicht", "FORM", 246
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONAUDIT.COVERAGE_SUMMARY", "Coverage summary", "FORM", 246
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONAUDIT.COVERAGE_SUMMARY", "Resume de couverture", "FORM", 246
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONAUDIT.REFRESH_AUDIT", "Audit pruefen", "FORM", 247
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONAUDIT.REFRESH_AUDIT", "Run audit", "FORM", 247
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONAUDIT.REFRESH_AUDIT", "Verifier l'audit", "FORM", 247
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONAUDIT.CREATE_MISSING_ROWS", "Fehlende Eintraege erzeugen", "FORM", 248
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONAUDIT.CREATE_MISSING_ROWS", "Create missing rows", "FORM", 248
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONAUDIT.CREATE_MISSING_ROWS", "Creer les lignes manquantes", "FORM", 248
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONAUDIT.OPEN_TRANSLATION", "Uebersetzung oeffnen", "FORM", 249
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONAUDIT.OPEN_TRANSLATION", "Open translation", "FORM", 249
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONAUDIT.OPEN_TRANSLATION", "Ouvrir la traduction", "FORM", 249
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONAUDIT.CLEAR_FILTERS", "Filter loeschen", "FORM", 250
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONAUDIT.CLEAR_FILTERS", "Clear filters", "FORM", 250
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONAUDIT.CLEAR_FILTERS", "Effacer les filtres", "FORM", 250

    EnsureTranslationSeed db, "de-CH", "MSG.FW_TRANSLATION_AUDIT_LOAD_ERROR", "Fehler beim Laden des Uebersetzungs-Audits.", "MSG", 251
    EnsureTranslationSeed db, "en-US", "MSG.FW_TRANSLATION_AUDIT_LOAD_ERROR", "Error loading the translation audit.", "MSG", 251
    EnsureTranslationSeed db, "fr-CH", "MSG.FW_TRANSLATION_AUDIT_LOAD_ERROR", "Erreur lors du chargement de l'audit des traductions.", "MSG", 251
    EnsureTranslationSeed db, "de-CH", "MSG.FW_TRANSLATION_AUDIT_REFRESH_ERROR", "Fehler beim Aktualisieren des Uebersetzungs-Audits.", "MSG", 252
    EnsureTranslationSeed db, "en-US", "MSG.FW_TRANSLATION_AUDIT_REFRESH_ERROR", "Error refreshing the translation audit.", "MSG", 252
    EnsureTranslationSeed db, "fr-CH", "MSG.FW_TRANSLATION_AUDIT_REFRESH_ERROR", "Erreur lors de l'actualisation de l'audit des traductions.", "MSG", 252
    EnsureTranslationSeed db, "de-CH", "MSG.FW_TRANSLATION_AUDIT_CREATE_MISSING_ERROR", "Fehler beim Erzeugen fehlender Uebersetzungseintraege.", "MSG", 253
    EnsureTranslationSeed db, "en-US", "MSG.FW_TRANSLATION_AUDIT_CREATE_MISSING_ERROR", "Error creating missing translation rows.", "MSG", 253
    EnsureTranslationSeed db, "fr-CH", "MSG.FW_TRANSLATION_AUDIT_CREATE_MISSING_ERROR", "Erreur lors de la creation des lignes de traduction manquantes.", "MSG", 253
    EnsureTranslationSeed db, "de-CH", "MSG.FW_TRANSLATION_AUDIT_OPEN_ERROR", "Fehler beim Oeffnen der Uebersetzung.", "MSG", 254
    EnsureTranslationSeed db, "en-US", "MSG.FW_TRANSLATION_AUDIT_OPEN_ERROR", "Error opening the translation.", "MSG", 254
    EnsureTranslationSeed db, "fr-CH", "MSG.FW_TRANSLATION_AUDIT_OPEN_ERROR", "Erreur lors de l'ouverture de la traduction.", "MSG", 254
    EnsureTranslationSeed db, "de-CH", "MSG.FW_TRANSLATION_AUDIT_SELECT_FIRST", "Bitte zuerst einen Uebersetzungseintrag auswaehlen.", "MSG", 255
    EnsureTranslationSeed db, "en-US", "MSG.FW_TRANSLATION_AUDIT_SELECT_FIRST", "Please select a translation entry first.", "MSG", 255
    EnsureTranslationSeed db, "fr-CH", "MSG.FW_TRANSLATION_AUDIT_SELECT_FIRST", "Veuillez d'abord selectionner une entree de traduction.", "MSG", 255

    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONEDIT.FORM_TITLE", "Uebersetzung bearbeiten", "FORM", 256
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONEDIT.FORM_TITLE", "Edit translation", "FORM", 256
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONEDIT.FORM_TITLE", "Modifier la traduction", "FORM", 256
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONEDIT.TRANSLATION_KEY", "Uebersetzungsschluessel", "FORM", 257
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONEDIT.TRANSLATION_KEY", "Translation key", "FORM", 257
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONEDIT.TRANSLATION_KEY", "Cle de traduction", "FORM", 257
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONEDIT.SCOPE_CODE", "Bereich", "FORM", 258
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONEDIT.SCOPE_CODE", "Scope", "FORM", 258
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONEDIT.SCOPE_CODE", "Portee", "FORM", 258
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONEDIT.AUDIT_STATUS", "Audit-Status", "FORM", 259
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONEDIT.AUDIT_STATUS", "Audit status", "FORM", 259
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONEDIT.AUDIT_STATUS", "Statut d'audit", "FORM", 259
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_TYPE", "Quelltyp", "FORM", 260
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_TYPE", "Source type", "FORM", 260
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_TYPE", "Type de source", "FORM", 260
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_OBJECT", "Quellobjekt", "FORM", 261
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_OBJECT", "Source object", "FORM", 261
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_OBJECT", "Objet source", "FORM", 261
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_CONTROL", "Quellsteuerelement", "FORM", 262
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_CONTROL", "Source control", "FORM", 262
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_CONTROL", "Controle source", "FORM", 262
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONEDIT.FALLBACK_TEXT", "Fallback-Text", "FORM", 263
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONEDIT.FALLBACK_TEXT", "Fallback text", "FORM", 263
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONEDIT.FALLBACK_TEXT", "Texte de secours", "FORM", 263
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONEDIT.DE_CH", "Deutsch (Schweiz)", "FORM", 264
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONEDIT.DE_CH", "German (Switzerland)", "FORM", 264
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONEDIT.DE_CH", "Allemand (Suisse)", "FORM", 264
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONEDIT.EN_US", "Englisch (USA)", "FORM", 265
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONEDIT.EN_US", "English (US)", "FORM", 265
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONEDIT.EN_US", "Anglais (Etats-Unis)", "FORM", 265
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONEDIT.FR_FR", "Franzoesisch (Schweiz)", "FORM", 266
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONEDIT.FR_FR", "French (Switzerland)", "FORM", 266
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONEDIT.FR_FR", "Francais (Suisse)", "FORM", 266
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONEDIT.SAVE", "Speichern", "FORM", 267
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONEDIT.SAVE", "Save", "FORM", 267
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONEDIT.SAVE", "Enregistrer", "FORM", 267
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONEDIT.CANCEL", "Abbrechen", "FORM", 268
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONEDIT.CANCEL", "Cancel", "FORM", 268
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONEDIT.CANCEL", "Annuler", "FORM", 268
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONEDIT.DEEPL_SUGGESTION", "DeepL Vorschlag", "FORM", 269
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONEDIT.DEEPL_SUGGESTION", "DeepL suggestion", "FORM", 269
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONEDIT.DEEPL_SUGGESTION", "Suggestion DeepL", "FORM", 269
    EnsureTranslationSeed db, "de-CH", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_LANGUAGE", "Ausgangssprache", "FORM", 270
    EnsureTranslationSeed db, "en-US", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_LANGUAGE", "Source language", "FORM", 270
    EnsureTranslationSeed db, "fr-CH", "FORM.FRMFWTRANSLATIONEDIT.SOURCE_LANGUAGE", "Langue source", "FORM", 270

    EnsureTranslationSeed db, "de-CH", "MSG.TRANSLATION_KEY_REQUIRED", "Uebersetzungsschluessel ist erforderlich.", "MSG", 270
    EnsureTranslationSeed db, "en-US", "MSG.TRANSLATION_KEY_REQUIRED", "Translation key is required.", "MSG", 270
    EnsureTranslationSeed db, "fr-CH", "MSG.TRANSLATION_KEY_REQUIRED", "La cle de traduction est obligatoire.", "MSG", 270
    EnsureTranslationSeed db, "de-CH", "MSG.TRANSLATION_EDIT_LOAD_ERROR", "Fehler beim Laden der Uebersetzung.", "MSG", 271
    EnsureTranslationSeed db, "en-US", "MSG.TRANSLATION_EDIT_LOAD_ERROR", "Error loading the translation.", "MSG", 271
    EnsureTranslationSeed db, "fr-CH", "MSG.TRANSLATION_EDIT_LOAD_ERROR", "Erreur lors du chargement de la traduction.", "MSG", 271
    EnsureTranslationSeed db, "de-CH", "MSG.TRANSLATION_EDIT_SAVE_ERROR", "Fehler beim Speichern der Uebersetzung.", "MSG", 272
    EnsureTranslationSeed db, "en-US", "MSG.TRANSLATION_EDIT_SAVE_ERROR", "Error saving the translation.", "MSG", 272
    EnsureTranslationSeed db, "fr-CH", "MSG.TRANSLATION_EDIT_SAVE_ERROR", "Erreur lors de l'enregistrement de la traduction.", "MSG", 272
    EnsureTranslationSeed db, "de-CH", "MSG.TRANSLATION_EDIT_CANCEL_CONFIRM", "Ungespeicherte Aenderungen verwerfen?", "MSG", 273
    EnsureTranslationSeed db, "en-US", "MSG.TRANSLATION_EDIT_CANCEL_CONFIRM", "Discard unsaved changes?", "MSG", 273
    EnsureTranslationSeed db, "fr-CH", "MSG.TRANSLATION_EDIT_CANCEL_CONFIRM", "Ignorer les modifications non enregistrees ?", "MSG", 273
    EnsureTranslationSeed db, "de-CH", "MSG.DEEPL_API_KEY_MISSING", "DeepL API-Schluessel ist nicht konfiguriert.", "MSG", 274
    EnsureTranslationSeed db, "en-US", "MSG.DEEPL_API_KEY_MISSING", "DeepL API key is not configured.", "MSG", 274
    EnsureTranslationSeed db, "fr-CH", "MSG.DEEPL_API_KEY_MISSING", "La cle API DeepL n'est pas configuree.", "MSG", 274
    EnsureTranslationSeed db, "de-CH", "MSG.TRANSLATION_EDIT_DEEPL_ERROR", "Fehler beim Abrufen der DeepL-Vorschlaege.", "MSG", 275
    EnsureTranslationSeed db, "en-US", "MSG.TRANSLATION_EDIT_DEEPL_ERROR", "Error retrieving DeepL suggestions.", "MSG", 275
    EnsureTranslationSeed db, "fr-CH", "MSG.TRANSLATION_EDIT_DEEPL_ERROR", "Erreur lors de la recuperation des suggestions DeepL.", "MSG", 275
    EnsureTranslationSeed db, "de-CH", "MSG.TRANSLATION_EDIT_DEEPL_SOURCE_REQUIRED", "Bitte zuerst einen Text in der Ausgangssprache erfassen.", "MSG", 276
    EnsureTranslationSeed db, "en-US", "MSG.TRANSLATION_EDIT_DEEPL_SOURCE_REQUIRED", "Please enter a text in the source language first.", "MSG", 276
    EnsureTranslationSeed db, "fr-CH", "MSG.TRANSLATION_EDIT_DEEPL_SOURCE_REQUIRED", "Veuillez d'abord saisir un texte dans la langue source.", "MSG", 276
    EnsureTranslationSeed db, "de-CH", "MSG.TRANSLATION_EDIT_DEEPL_OVERWRITE_CONFIRM", "Bestehende Zieltexte mit DeepL-Vorschlaegen ersetzen?", "MSG", 277
    EnsureTranslationSeed db, "en-US", "MSG.TRANSLATION_EDIT_DEEPL_OVERWRITE_CONFIRM", "Replace existing target texts with DeepL suggestions?", "MSG", 277
    EnsureTranslationSeed db, "fr-CH", "MSG.TRANSLATION_EDIT_DEEPL_OVERWRITE_CONFIRM", "Remplacer les textes cibles existants par des suggestions DeepL ?", "MSG", 277

End Sub

Public Sub EnsureTranslationSeed( _
    ByVal db As DAO.Database, _
    ByVal languageCode As String, _
    ByVal translationKey As String, _
    ByVal TranslationValue As String, _
    ByVal moduleCode As String, _
    ByVal sortOrder As Long)

    languageCode = NormalizeSeedLanguageCode(languageCode)

    If Not TranslationSeedExists(db, languageCode, translationKey) Then
        InsertTranslation db, languageCode, translationKey, TranslationValue, True, moduleCode, sortOrder
    End If
End Sub

Private Sub UpdateTranslationSeed( _
    ByVal db As DAO.Database, _
    ByVal languageCode As String, _
    ByVal translationKey As String, _
    ByVal TranslationValue As String, _
    ByVal isActive As Boolean, _
    ByVal moduleCode As String, _
    ByVal sortOrder As Long)

    Dim sqlStatement As String

    languageCode = NormalizeSeedLanguageCode(languageCode)

    sqlStatement = "UPDATE fw_translation SET " & _
                   "translation_value = " & SqlText(TranslationValue) & ", " & _
                   "is_active = " & IIf(isActive, "True", "False") & ", " & _
                   "module_code = " & SqlNullableText(moduleCode) & ", " & _
                   "sort_order = " & CStr(sortOrder) & ", " & _
                   "updated_at = Now(), " & _
                   "updated_by = 'SYSTEM' " & _
                   "WHERE language_code = " & SqlText(languageCode) & " " & _
                   "AND translation_key = " & SqlText(translationKey) & ";"

    db.Execute sqlStatement, dbFailOnError
End Sub

Public Function TranslationSeedExists( _
    ByVal db As DAO.Database, _
    ByVal languageCode As String, _
    ByVal translationKey As String) As Boolean
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlStatement As String

    languageCode = NormalizeSeedLanguageCode(languageCode)

    sqlStatement = "SELECT TOP 1 translation_key " & _
                   "FROM fw_translation " & _
                   "WHERE language_code = " & SqlText(languageCode) & " " & _
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

Private Function NormalizeSeedLanguageCode(ByVal languageCode As String) As String
    NormalizeSeedLanguageCode = Trim$(modFwTranslationRuntime.NormalizeProjectLanguageCode(languageCode))
End Function

Public Sub NormalizeLanguageCodeData()
    On Error GoTo ErrorHandler

    Dim systemDb As DAO.Database
    Dim tenantDb As DAO.Database
    Dim workspace As DAO.Workspace

    Set systemDb = modDb.GetSystemDatabase()
    Set tenantDb = modDb.GetCurrentTenantDatabase()
    Set workspace = DBEngine.Workspaces(0)

    workspace.BeginTrans

    NormalizeFwTranslationLanguageCodes systemDb
    NormalizeLanguageCodesForTable systemDb, "ref_language", "language_code", Array()
    NormalizeLanguageCodesForTable tenantDb, "adr_address", "language_code", Array("address_id")

    If Not modBasicModuleSchema.EnsureSystemLanguageReferenceSchema() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".NormalizeLanguageCodeData", _
            "ref_language could not be fully ensured after language normalization."
    End If

    workspace.CommitTrans

    modLoggingHandler.LogInfo MODULE_NAME & ".NormalizeLanguageCodeData", _
        "Language-code normalization finished."
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not workspace Is Nothing Then
        workspace.Rollback
    End If
    modErrorHandler.HandleError MODULE_NAME, "NormalizeLanguageCodeData", Err
End Sub

Private Sub NormalizeLanguageCodesForTable( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal languageFieldName As String, _
    ByVal keyFields As Variant)
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sourceLanguageCode As String
    Dim targetLanguageCode As String
    Dim updatedCount As Long
    Dim deletedDuplicateCount As Long
    Dim unchangedCount As Long
    Dim unresolvedConflictCount As Long

    If db Is Nothing Then
        Exit Sub
    End If

    If Not modDbSchema.TableExists(db, tableName) Then
        Exit Sub
    End If

    If Not modDbSchema.FieldExists(db, tableName, languageFieldName) Then
        Exit Sub
    End If

    Set rs = db.OpenRecordset("SELECT * FROM [" & tableName & "];", dbOpenDynaset)

    Do While Not rs.EOF
        sourceLanguageCode = Trim$(modDaoHelper.NzString(rs.Fields(languageFieldName).Value))
        targetLanguageCode = NormalizeSeedLanguageCode(sourceLanguageCode)

        If LenB(sourceLanguageCode) > 0 Then
            If LenB(targetLanguageCode) > 0 Then
                If StrComp(sourceLanguageCode, targetLanguageCode, vbBinaryCompare) <> 0 Then
                    rs.Edit
                    rs.Fields(languageFieldName).Value = targetLanguageCode
                    rs.Update
                    updatedCount = updatedCount + 1
                Else
                    unchangedCount = unchangedCount + 1
                End If
            Else
                unchangedCount = unchangedCount + 1
            End If
        Else
            unchangedCount = unchangedCount + 1
        End If

        rs.MoveNext
    Loop

    modLoggingHandler.LogInfo MODULE_NAME & ".NormalizeLanguageCodesForTable", _
        "table='" & tableName & "'; updated_count=" & CStr(updatedCount) & "; deleted_duplicate_count=" & CStr(deletedDuplicateCount) & "; unchanged_count=" & CStr(unchangedCount) & "; unresolved_conflict_count=" & CStr(unresolvedConflictCount) & "."

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NormalizeLanguageCodesForTable", Err
    Resume CleanExit
End Sub

Private Sub NormalizeFwTranslationLanguageCodes(ByVal db As DAO.Database)
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim translationGroups As Object
    Dim groupRows As Collection
    Dim translationKey As Variant
    Dim rowData As Object
    Dim updatedCount As Long
    Dim deletedDuplicateCount As Long
    Dim unchangedCount As Long
    Dim unresolvedConflictCount As Long

    If db Is Nothing Then
        Exit Sub
    End If

    If Not modDbSchema.TableExists(db, "fw_translation") Then
        Exit Sub
    End If

    If Not modDbSchema.FieldExists(db, "fw_translation", "translation_id") Then
        Exit Sub
    End If

    If Not modDbSchema.FieldExists(db, "fw_translation", "translation_key") Then
        Exit Sub
    End If

    If Not modDbSchema.FieldExists(db, "fw_translation", "language_code") Then
        Exit Sub
    End If

    Set translationGroups = CreateObject("Scripting.Dictionary")
    translationGroups.CompareMode = vbTextCompare
    Set rs = db.OpenRecordset( _
        "SELECT translation_id, translation_key, language_code " & _
        "FROM fw_translation " & _
        "ORDER BY translation_key, translation_id;", _
        dbOpenSnapshot)

    Do While Not rs.EOF
        translationKey = modDaoHelper.NzString(rs.Fields("translation_key").Value)

        If Not translationGroups.Exists(CStr(translationKey)) Then
            Set groupRows = New Collection
            translationGroups.Add CStr(translationKey), groupRows
        End If

        Set rowData = CreateObject("Scripting.Dictionary")
        rowData.CompareMode = vbTextCompare
        rowData("translation_id") = modDaoHelper.NzLong(rs.Fields("translation_id").Value, 0)
        rowData("translation_key") = CStr(translationKey)
        rowData("language_code") = modDaoHelper.NzString(rs.Fields("language_code").Value)
        rowData("target_language_code") = MapFwTranslationLanguageCode(CStr(rowData("language_code")))
        rowData("priority") = 0&

        translationGroups(CStr(translationKey)).Add rowData
        rs.MoveNext
    Loop

    For Each translationKey In translationGroups.Keys
        Set groupRows = translationGroups(CStr(translationKey))
        NormalizeFwTranslationGroup db, groupRows, updatedCount, deletedDuplicateCount, unchangedCount, unresolvedConflictCount
    Next translationKey

    modLoggingHandler.LogInfo MODULE_NAME & ".NormalizeFwTranslationLanguageCodes", _
        "table='fw_translation'; updated_count=" & CStr(updatedCount) & "; deleted_duplicate_count=" & CStr(deletedDuplicateCount) & "; unchanged_count=" & CStr(unchangedCount) & "; unresolved_conflict_count=" & CStr(unresolvedConflictCount) & "."
    GoTo CleanExit

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NormalizeFwTranslationLanguageCodes", Err
    Resume CleanExit

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Sub

Private Sub NormalizeFwTranslationGroup( _
    ByVal db As DAO.Database, _
    ByVal groupRows As Collection, _
    ByRef updatedCount As Long, _
    ByRef deletedDuplicateCount As Long, _
    ByRef unchangedCount As Long, _
    ByRef unresolvedConflictCount As Long)
    On Error GoTo ErrorHandler

    Dim targetLanguageCode As Variant
    Dim candidates As Collection
    Dim keeperRow As Object
    Dim candidateRow As Object
    Dim canonicalRowCount As Long

    For Each targetLanguageCode In modFwTranslationRuntime.GetSupportedTranslationLanguages()
        Set candidates = CollectFwTranslationCandidates(groupRows, CStr(targetLanguageCode))

        If candidates.Count > 0 Then
            canonicalRowCount = CountFwTranslationCanonicalRows(candidates, CStr(targetLanguageCode))

            If canonicalRowCount > 1 Then
                unresolvedConflictCount = unresolvedConflictCount + 1
                modLoggingHandler.LogWarning MODULE_NAME & ".NormalizeFwTranslationGroup", _
                    "Unresolved canonical duplicate. translation_key='" & CStr(groupRows(1)("translation_key")) & "'; target_language='" & CStr(targetLanguageCode) & "'."
            Else
                Set keeperRow = SelectFwTranslationKeeper(candidates, CStr(targetLanguageCode))

                If keeperRow Is Nothing Then
                    unresolvedConflictCount = unresolvedConflictCount + 1
                Else
                    If StrComp(CStr(keeperRow("language_code")), CStr(targetLanguageCode), vbBinaryCompare) = 0 Then
                        unchangedCount = unchangedCount + 1
                    Else
                        UpdateFwTranslationLanguageCode db, CLng(keeperRow("translation_id")), CStr(targetLanguageCode)
                        keeperRow("language_code") = CStr(targetLanguageCode)
                        keeperRow("target_language_code") = CStr(targetLanguageCode)
                        updatedCount = updatedCount + 1
                    End If

                    For Each candidateRow In candidates
                        If CLng(candidateRow("translation_id")) <> CLng(keeperRow("translation_id")) Then
                            If StrComp(CStr(candidateRow("language_code")), CStr(targetLanguageCode), vbBinaryCompare) = 0 Then
                                unresolvedConflictCount = unresolvedConflictCount + 1
                                modLoggingHandler.LogWarning MODULE_NAME & ".NormalizeFwTranslationGroup", _
                                    "Skipped duplicate canonical row. translation_key='" & CStr(candidateRow("translation_key")) & "'; target_language='" & CStr(targetLanguageCode) & "'; translation_id=" & CStr(candidateRow("translation_id")) & "."
                            Else
                                DeleteFwTranslationRow db, CLng(candidateRow("translation_id"))
                                deletedDuplicateCount = deletedDuplicateCount + 1
                            End If
                        End If
                    Next candidateRow
                End If
            End If
        End If
    Next targetLanguageCode

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NormalizeFwTranslationGroup", Err
End Sub

Private Function CollectFwTranslationCandidates( _
    ByVal groupRows As Collection, _
    ByVal targetLanguageCode As String) As Collection

    Dim rowData As Object
    Dim candidates As Collection

    Set candidates = New Collection

    For Each rowData In groupRows
        If StrComp(CStr(rowData("target_language_code")), targetLanguageCode, vbTextCompare) = 0 Then
            rowData("priority") = GetFwTranslationLanguagePriority(CStr(rowData("language_code")), targetLanguageCode)
            candidates.Add rowData
        End If
    Next rowData

    Set CollectFwTranslationCandidates = candidates
End Function

Private Function CountFwTranslationCanonicalRows( _
    ByVal candidates As Collection, _
    ByVal targetLanguageCode As String) As Long

    Dim rowData As Object

    For Each rowData In candidates
        If StrComp(CStr(rowData("language_code")), targetLanguageCode, vbBinaryCompare) = 0 Then
            CountFwTranslationCanonicalRows = CountFwTranslationCanonicalRows + 1
        End If
    Next rowData
End Function

Private Function SelectFwTranslationKeeper( _
    ByVal candidates As Collection, _
    ByVal targetLanguageCode As String) As Object

    Dim rowData As Object
    Dim keeperRow As Object
    Dim currentPriority As Long
    Dim keeperPriority As Long

    For Each rowData In candidates
        currentPriority = CLng(rowData("priority"))

        If keeperRow Is Nothing Then
            Set keeperRow = rowData
        Else
            keeperPriority = CLng(keeperRow("priority"))

            If currentPriority < keeperPriority Then
                Set keeperRow = rowData
            ElseIf currentPriority = keeperPriority Then
                If CLng(rowData("translation_id")) < CLng(keeperRow("translation_id")) Then
                    Set keeperRow = rowData
                End If
            End If
        End If
    Next rowData

    Set SelectFwTranslationKeeper = keeperRow
End Function

Private Function MapFwTranslationLanguageCode(ByVal languageCode As String) As String
    Dim comparableLanguageCode As String

    comparableLanguageCode = UCase$(Replace(Trim$(modDaoHelper.NzString(languageCode)), "_", "-"))

    Select Case comparableLanguageCode
        Case "DE", "DE-CH", "DE-DE"
            MapFwTranslationLanguageCode = "de-CH"
        Case "EN", "EN-US"
            MapFwTranslationLanguageCode = "en-US"
        Case "FR", "FR-CH", "FR-FR"
            MapFwTranslationLanguageCode = "fr-CH"
        Case Else
            MapFwTranslationLanguageCode = Trim$(modDaoHelper.NzString(languageCode))
    End Select
End Function

Private Function GetFwTranslationLanguagePriority( _
    ByVal languageCode As String, _
    ByVal targetLanguageCode As String) As Long

    Dim comparableLanguageCode As String

    comparableLanguageCode = UCase$(Replace(Trim$(modDaoHelper.NzString(languageCode)), "_", "-"))

    Select Case targetLanguageCode
        Case "de-CH"
            Select Case comparableLanguageCode
                Case "DE-CH"
                    If StrComp(Trim$(modDaoHelper.NzString(languageCode)), "de-CH", vbBinaryCompare) = 0 Then
                        GetFwTranslationLanguagePriority = 1
                    Else
                        GetFwTranslationLanguagePriority = 2
                    End If
                Case "DE-DE"
                    GetFwTranslationLanguagePriority = 3
                Case "DE"
                    GetFwTranslationLanguagePriority = 4
                Case Else
                    GetFwTranslationLanguagePriority = 100
            End Select
        Case "en-US"
            Select Case comparableLanguageCode
                Case "EN-US"
                    If StrComp(Trim$(modDaoHelper.NzString(languageCode)), "en-US", vbBinaryCompare) = 0 Then
                        GetFwTranslationLanguagePriority = 1
                    Else
                        GetFwTranslationLanguagePriority = 2
                    End If
                Case "EN"
                    GetFwTranslationLanguagePriority = 3
                Case Else
                    GetFwTranslationLanguagePriority = 100
            End Select
        Case "fr-CH"
            Select Case comparableLanguageCode
                Case "FR-CH"
                    If StrComp(Trim$(modDaoHelper.NzString(languageCode)), "fr-CH", vbBinaryCompare) = 0 Then
                        GetFwTranslationLanguagePriority = 1
                    Else
                        GetFwTranslationLanguagePriority = 2
                    End If
                Case "FR-FR"
                    GetFwTranslationLanguagePriority = 3
                Case "FR"
                    GetFwTranslationLanguagePriority = 4
                Case Else
                    GetFwTranslationLanguagePriority = 100
            End Select
        Case Else
            GetFwTranslationLanguagePriority = 100
    End Select
End Function

Private Sub UpdateFwTranslationLanguageCode( _
    ByVal db As DAO.Database, _
    ByVal translationId As Long, _
    ByVal targetLanguageCode As String)

    db.Execute _
        "UPDATE fw_translation " & _
        "SET language_code = " & SqlText(targetLanguageCode) & " " & _
        "WHERE translation_id = " & CStr(translationId) & ";", _
        dbFailOnError
End Sub

Private Sub DeleteFwTranslationRow( _
    ByVal db As DAO.Database, _
    ByVal translationId As Long)

    db.Execute _
        "DELETE FROM fw_translation " & _
        "WHERE translation_id = " & CStr(translationId) & ";", _
        dbFailOnError
End Sub

Public Sub SeedTagHelp()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = currentDb

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


