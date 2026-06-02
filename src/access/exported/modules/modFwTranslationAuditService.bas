Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFwTranslationAuditService
' Purpose   : Builds deterministic translation audit work data for expected
'             navigation, reference, form-tag, and registry translation keys,
'             and classifies historical legacy keys separately from true orphans.
' Author    : Codex
' Version   : 0.3.0
'===============================================================================

Private Const MODULE_NAME As String = "modFwTranslationAuditService"

Private Const TABLE_AUDIT As String = "tmp_fw_translation_audit"
Private Const TABLE_TRANSLATION As String = "fw_translation"
Private Const TABLE_NAVIGATION As String = "fw_navigation"
Private Const TABLE_EXPECTED As String = "fw_translation_expected"

Private Const STATUS_OK As String = "OK"
Private Const STATUS_MISSING_ROW As String = "MISSING_ROW"
Private Const STATUS_EMPTY_VALUE As String = "EMPTY_VALUE"
Private Const STATUS_ORPHAN As String = "ORPHAN"
Private Const STATUS_LEGACY_KEY As String = "LEGACY_KEY"

Private Const SCOPE_FORM As String = "FORM"
Private Const SCOPE_MSG As String = "MSG"
Private Const SCOPE_STATUS As String = "STATUS"
Private Const SCOPE_COMMON As String = "COMMON"
Private Const SCOPE_NAV As String = "NAV"
Private Const SCOPE_REF As String = "REF"

Private Const SOURCE_FORM As String = "FORM"
Private Const SOURCE_REGISTRY As String = "REGISTRY"
Private Const SOURCE_NAVIGATION As String = "NAVIGATION"
Private Const SOURCE_REFERENCE As String = "REFERENCE"
Private Const SOURCE_TRANSLATION As String = "FW_TRANSLATION"

Private Const FIELD_TRANSLATION_KEY As String = "translation_key"
Private Const FIELD_LANGUAGE_CODE As String = "language_code"
Private Const FIELD_TRANSLATION_VALUE As String = "translation_value"
Private Const FIELD_MODULE_CODE As String = "module_code"
Private Const FIELD_IS_ACTIVE As String = "is_active"
Private Const FIELD_SORT_ORDER As String = "sort_order"
Private Const FIELD_CREATED_AT As String = "created_at"
Private Const FIELD_CREATED_BY As String = "created_by"
Private Const FIELD_UPDATED_AT As String = "updated_at"
Private Const FIELD_UPDATED_BY As String = "updated_by"

Private Const FIELD_NAVIGATION_ID As String = "navigation_id"
Private Const FIELD_OBJECT_NAME As String = "object_name"
Private Const FIELD_CAPTION_KEY As String = "caption_key"
Private Const FIELD_FALLBACK_CAPTION As String = "fallback_caption"
Private Const FIELD_SCOPE_CODE As String = "scope_code"
Private Const FIELD_SOURCE_OBJECT As String = "source_object"
Private Const FIELD_SOURCE_CONTROL As String = "source_control"
Private Const FIELD_SOURCE_TYPE As String = "source_type"
Private Const FIELD_AUDIT_STATUS As String = "audit_status"
Private Const FIELD_EXISTS_IN_TRANSLATION As String = "exists_in_fw_translation"
Private Const FIELD_FALLBACK_TEXT As String = "fallback_text"

Private Const TR_PREFIX As String = "TR:"

Public Sub EnsureTranslationAuditTable()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    If TableExists(db, TABLE_AUDIT) Then
        Exit Sub
    End If

    db.Execute _
        "CREATE TABLE " & TABLE_AUDIT & " (" & _
        "audit_id AUTOINCREMENT CONSTRAINT pk_tmp_fw_translation_audit PRIMARY KEY, " & _
        "scope_code TEXT(50), " & _
        "translation_key TEXT(255), " & _
        "language_code TEXT(20), " & _
        "audit_status TEXT(50), " & _
        "source_type TEXT(50), " & _
        "source_object TEXT(255), " & _
        "source_control TEXT(255), " & _
        "fallback_text LONGTEXT, " & _
        "translation_value LONGTEXT, " & _
        "exists_in_fw_translation YESNO, " & _
        "created_at DATETIME" & _
        ");", dbFailOnError

    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureTranslationAuditTable", _
        "Created audit work table '" & TABLE_AUDIT & "'."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureTranslationAuditTable", Err
End Sub

Public Sub EnsureTranslationExpectedTable()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    If TableExists(db, TABLE_EXPECTED) Then
        Exit Sub
    End If

    db.Execute _
        "CREATE TABLE " & TABLE_EXPECTED & " (" & _
        "expected_id AUTOINCREMENT CONSTRAINT pk_fw_translation_expected PRIMARY KEY, " & _
        "scope_code TEXT(50), " & _
        "translation_key TEXT(255), " & _
        "source_object TEXT(255), " & _
        "fallback_text LONGTEXT, " & _
        "is_active YESNO, " & _
        "created_at DATETIME, " & _
        "created_by TEXT(255), " & _
        "updated_at DATETIME, " & _
        "updated_by TEXT(255)" & _
        ");", dbFailOnError

    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureTranslationExpectedTable", _
        "Created registry table '" & TABLE_EXPECTED & "'."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureTranslationExpectedTable", Err
End Sub

Public Sub BuildTranslationAuditData()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim expectedKeys As Collection
    Dim expectedCount As Long
    Dim orphanCount As Long

    EnsureTranslationAuditTable
    EnsureTranslationExpectedTable
    EnsureRegistryExpectedKeysFromActiveTranslations

    Set db = CurrentDb
    db.Execute "DELETE FROM " & TABLE_AUDIT, dbFailOnError

    Set expectedKeys = New Collection

    expectedCount = expectedCount + CollectFormExpectedKeys(db, expectedKeys)
    expectedCount = expectedCount + CollectRegistryExpectedKeys(db, expectedKeys)
    expectedCount = expectedCount + CollectNavigationExpectedKeys(db, expectedKeys)
    expectedCount = expectedCount + CollectReferenceExpectedKeys(db, expectedKeys, "ref_unit", "", "unit_code")
    expectedCount = expectedCount + CollectReferenceExpectedKeys(db, expectedKeys, "ref_vat_code", "", "vat_code")
    expectedCount = expectedCount + CollectReferenceExpectedKeys(db, expectedKeys, "ref_article_type_code", "article_type_name", "article_type_code")
    expectedCount = expectedCount + CollectReferenceExpectedKeys(db, expectedKeys, "ref_address_type", "", "address_type_code")
    expectedCount = expectedCount + CollectReferenceExpectedKeys(db, expectedKeys, "ref_salutation", "", "salutation_code")
    expectedCount = expectedCount + CollectReferenceExpectedKeys(db, expectedKeys, "ref_addressing_mode", "", "addressing_mode_code")
    expectedCount = expectedCount + CollectReferenceExpectedKeys(db, expectedKeys, "ref_contact_type", "", "contact_type_code")

    orphanCount = AppendOrphanAuditRows(db, expectedKeys)

    modLoggingHandler.LogInfo MODULE_NAME & ".BuildTranslationAuditData", _
        "Built translation audit data. expected_keys=" & CStr(expectedCount) & _
        "; orphan_rows=" & CStr(orphanCount) & "; " & GetTranslationCoverageSummary()
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "BuildTranslationAuditData", Err
End Sub

Public Sub EnsureMissingTranslationRows()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim insertedCount As Long
    Dim translationKey As String
    Dim languageCode As String
    Dim scopeCode As String

    EnsureTranslationAuditTable

    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT translation_key, language_code, scope_code " & _
        "FROM " & TABLE_AUDIT & " " & _
        "WHERE audit_status = " & SqlText(STATUS_MISSING_ROW) & " " & _
        "ORDER BY translation_key, language_code;", dbOpenSnapshot)

    Do While Not rs.EOF
        translationKey = modDaoHelper.NzString(rs.Fields(FIELD_TRANSLATION_KEY).Value)
        languageCode = modDaoHelper.NzString(rs.Fields(FIELD_LANGUAGE_CODE).Value)
        scopeCode = modDaoHelper.NzString(rs.Fields(FIELD_SCOPE_CODE).Value)

        If Not TranslationRowExists(db, translationKey, languageCode) Then
            InsertMissingTranslationRow db, translationKey, languageCode, scopeCode
            insertedCount = insertedCount + 1
        End If

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureMissingTranslationRows", _
        "Inserted " & CStr(insertedCount) & " missing fw_translation row(s)."
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    modErrorHandler.HandleError MODULE_NAME, "EnsureMissingTranslationRows", Err
End Sub

Public Function GetTranslationCoverageSummary() As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlStatement As String
    Dim summaryText As String

    Set db = CurrentDb
    If Not TableExists(db, TABLE_AUDIT) Then
        Exit Function
    End If

    sqlStatement = "SELECT scope_code, language_code, audit_status, Count(*) AS row_count " & _
                   "FROM " & TABLE_AUDIT & " " & _
                   "GROUP BY scope_code, language_code, audit_status " & _
                   "ORDER BY scope_code, language_code, audit_status;"

    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)
    Do While Not rs.EOF
        If LenB(summaryText) > 0 Then
            summaryText = summaryText & " | "
        End If
        summaryText = summaryText & _
                      modDaoHelper.NzString(rs.Fields(FIELD_SCOPE_CODE).Value, "?") & "/" & _
                      modDaoHelper.NzString(rs.Fields(FIELD_LANGUAGE_CODE).Value, "?") & "/" & _
                      modDaoHelper.NzString(rs.Fields(FIELD_AUDIT_STATUS).Value, "?") & "=" & _
                      CStr(modDaoHelper.NzLong(rs.Fields("row_count").Value, 0))
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    GetTranslationCoverageSummary = summaryText
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    GetTranslationCoverageSummary = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "GetTranslationCoverageSummary", Err
End Function

Private Function CollectFormExpectedKeys(ByVal db As DAO.Database, ByRef expectedKeys As Collection) As Long
    On Error GoTo ErrorHandler

    Dim accessObject As Access.AccessObject
    Dim formName As String
    Dim wasAlreadyOpen As Boolean
    Dim openedByAudit As Boolean
    Dim formInstance As Access.Form
    Dim ctl As Control
    Dim translationKey As String

    For Each accessObject In CurrentProject.AllForms
        formName = accessObject.Name
        wasAlreadyOpen = IsProjectFormLoaded(formName)
        openedByAudit = OpenFormHiddenDesignIfNeeded(formName, wasAlreadyOpen)

        Set formInstance = ResolveScannableFormInstance(formName)
        If Not formInstance Is Nothing Then

            translationKey = NormalizeExpectedKey(modFwTranslationRuntime.GetTranslationKeyFromTag(GetTagTextSafely(formInstance.Tag)))
            If ResolveScopeCodeFromKey(translationKey) = SCOPE_FORM Then
                If AddExpectedKey(expectedKeys, translationKey) Then
                    AppendExpectedAuditRows db, SCOPE_FORM, translationKey, SOURCE_FORM, formName, vbNullString, NzCaptionTextSafely(formInstance, translationKey)
                    CollectFormExpectedKeys = CollectFormExpectedKeys + 1
                End If
            End If

            For Each ctl In formInstance.Controls
                translationKey = NormalizeExpectedKey(modFwTranslationRuntime.GetTranslationKeyFromTag(GetControlTagSafely(ctl)))
                If ResolveScopeCodeFromKey(translationKey) = SCOPE_FORM Then
                    If AddExpectedKey(expectedKeys, translationKey) Then
                        AppendExpectedAuditRows db, SCOPE_FORM, translationKey, SOURCE_FORM, formName, ctl.Name, GetControlFallbackText(ctl, translationKey)
                        CollectFormExpectedKeys = CollectFormExpectedKeys + 1
                    End If
                End If
            Next ctl
        End If

CleanForm:
        If openedByAudit And Not wasAlreadyOpen Then
            CloseFormNoSave formName
        End If
        Set formInstance = Nothing
    Next accessObject
    Exit Function

ErrorHandler:
    modLoggingHandler.LogWarning MODULE_NAME & ".CollectFormExpectedKeys", _
        "Skipped form '" & formName & "' during translation audit scan."
    Resume CleanForm
End Function

Private Function CollectRegistryExpectedKeys(ByVal db As DAO.Database, ByRef expectedKeys As Collection) As Long
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlStatement As String
    Dim translationKey As String
    Dim scopeCode As String
    Dim sourceObject As String
    Dim fallbackText As String

    If Not TableExists(db, TABLE_EXPECTED) Then
        Exit Function
    End If

    sqlStatement = "SELECT scope_code, translation_key, source_object, fallback_text " & _
                   "FROM " & TABLE_EXPECTED & " " & _
                   "WHERE Nz(is_active, True) = True " & _
                   "AND Len(Trim(Nz(translation_key, ''))) > 0;"

    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)
    Do While Not rs.EOF
        translationKey = NormalizeExpectedKey(modDaoHelper.NzString(rs.Fields(FIELD_TRANSLATION_KEY).Value))
        scopeCode = modDaoHelper.NzString(rs.Fields(FIELD_SCOPE_CODE).Value)
        If LenB(scopeCode) = 0 Then
            scopeCode = ResolveScopeCodeFromKey(translationKey)
        End If
        sourceObject = modDaoHelper.NzString(rs.Fields(FIELD_SOURCE_OBJECT).Value)
        fallbackText = modDaoHelper.NzString(rs.Fields(FIELD_FALLBACK_TEXT).Value, translationKey)

        If LenB(translationKey) > 0 And StrComp(scopeCode, "UNKNOWN", vbTextCompare) <> 0 Then
            If AddExpectedKey(expectedKeys, translationKey) Then
                AppendExpectedAuditRows db, scopeCode, translationKey, SOURCE_REGISTRY, sourceObject, vbNullString, fallbackText
                CollectRegistryExpectedKeys = CollectRegistryExpectedKeys + 1
            End If
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    modErrorHandler.HandleError MODULE_NAME, "CollectRegistryExpectedKeys", Err
End Function

Private Function CollectNavigationExpectedKeys(ByVal db As DAO.Database, ByRef expectedKeys As Collection) As Long
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlStatement As String
    Dim translationKey As String
    Dim fallbackText As String
    Dim sourceObject As String

    If Not TableExists(db, TABLE_NAVIGATION) Then
        Exit Function
    End If

    sqlStatement = "SELECT " & FIELD_NAVIGATION_ID & ", " & FIELD_OBJECT_NAME & ", " & _
                   FIELD_CAPTION_KEY & ", " & FIELD_FALLBACK_CAPTION & " " & _
                   "FROM " & TABLE_NAVIGATION & " " & _
                   "WHERE Len(Trim(Nz(" & FIELD_CAPTION_KEY & ", ''))) > 0;"

    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)
    Do While Not rs.EOF
        translationKey = NormalizeExpectedKey(modDaoHelper.NzString(rs.Fields(FIELD_CAPTION_KEY).Value))
        If LenB(translationKey) > 0 Then
            fallbackText = modDaoHelper.NzString(rs.Fields(FIELD_FALLBACK_CAPTION).Value, translationKey)
            sourceObject = modDaoHelper.NzString(rs.Fields(FIELD_OBJECT_NAME).Value)
            If LenB(sourceObject) = 0 Then
                sourceObject = CStr(modDaoHelper.NzLong(rs.Fields(FIELD_NAVIGATION_ID).Value, 0))
            End If

            If AddExpectedKey(expectedKeys, translationKey) Then
                AppendExpectedAuditRows db, SCOPE_NAV, translationKey, SOURCE_NAVIGATION, sourceObject, vbNullString, fallbackText
                CollectNavigationExpectedKeys = CollectNavigationExpectedKeys + 1
            End If
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    modErrorHandler.HandleError MODULE_NAME, "CollectNavigationExpectedKeys", Err
End Function

Private Function CollectReferenceExpectedKeys( _
    ByVal db As DAO.Database, _
    ByRef expectedKeys As Collection, _
    ByVal tableName As String, _
    ByVal preferredNameField As String, _
    ByVal fallbackCodeFieldName As String) As Long
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlStatement As String
    Dim translationKey As String
    Dim fallbackText As String
    Dim readableFieldName As String

    If Not TableExists(db, tableName) Then
        Exit Function
    End If

    sqlStatement = "SELECT * FROM " & tableName & ";"
    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)

    readableFieldName = ResolveReadableFieldName(rs, preferredNameField, fallbackCodeFieldName)

    Do While Not rs.EOF
        If modDaoHelper.RecordsetHasField(rs, FIELD_TRANSLATION_KEY) Then
            translationKey = NormalizeExpectedKey(modDaoHelper.NzString(rs.Fields(FIELD_TRANSLATION_KEY).Value))
            If LenB(translationKey) > 0 Then
                fallbackText = ResolveRecordsetFallbackText(rs, readableFieldName, translationKey)
                If AddExpectedKey(expectedKeys, translationKey) Then
                    AppendExpectedAuditRows db, SCOPE_REF, translationKey, SOURCE_REFERENCE, tableName, vbNullString, fallbackText
                    CollectReferenceExpectedKeys = CollectReferenceExpectedKeys + 1
                End If
            End If
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    modErrorHandler.HandleError MODULE_NAME, "CollectReferenceExpectedKeys", Err
End Function

Private Sub AppendExpectedAuditRows( _
    ByVal db As DAO.Database, _
    ByVal scopeCode As String, _
    ByVal translationKey As String, _
    ByVal sourceType As String, _
    ByVal sourceObject As String, _
    ByVal sourceControl As String, _
    ByVal fallbackText As String)

    Dim languages As Variant
    Dim languageCode As Variant
    Dim translationValue As Variant
    Dim existsInTranslation As Boolean
    Dim auditStatus As String

    languages = RequiredLanguages()

    For Each languageCode In languages
        existsInTranslation = TryGetTranslationValue(db, translationKey, CStr(languageCode), translationValue)

        If Not existsInTranslation Then
            auditStatus = STATUS_MISSING_ROW
        ElseIf LenB(Trim$(modDaoHelper.NzString(translationValue))) = 0 Then
            auditStatus = STATUS_EMPTY_VALUE
        Else
            auditStatus = STATUS_OK
        End If

        InsertAuditRow db, scopeCode, translationKey, CStr(languageCode), auditStatus, _
                       sourceType, sourceObject, sourceControl, fallbackText, translationValue, existsInTranslation
    Next languageCode
End Sub

Private Function AppendOrphanAuditRows(ByVal db As DAO.Database, ByRef expectedKeys As Collection) As Long
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlStatement As String
    Dim translationKey As String
    Dim languageCode As String
    Dim translationValue As Variant
    Dim sourceObject As String

    If Not TableExists(db, TABLE_TRANSLATION) Then
        Exit Function
    End If

    sqlStatement = "SELECT " & FIELD_TRANSLATION_KEY & ", " & FIELD_LANGUAGE_CODE & ", " & _
                   FIELD_TRANSLATION_VALUE & ", " & FIELD_MODULE_CODE & " " & _
                   "FROM " & TABLE_TRANSLATION & " " & _
                   "WHERE Len(Trim(Nz(" & FIELD_TRANSLATION_KEY & ", ''))) > 0;"

    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)
    Do While Not rs.EOF
        translationKey = NormalizeExpectedKey(modDaoHelper.NzString(rs.Fields(FIELD_TRANSLATION_KEY).Value))
        If LenB(translationKey) > 0 Then
            If Not CollectionContains(expectedKeys, BuildExpectedKeyToken(translationKey)) Then
                languageCode = modDaoHelper.NzString(rs.Fields(FIELD_LANGUAGE_CODE).Value)
                translationValue = rs.Fields(FIELD_TRANSLATION_VALUE).Value
                sourceObject = TABLE_TRANSLATION
                If modDaoHelper.RecordsetHasField(rs, FIELD_MODULE_CODE) Then
                    If LenB(Trim$(modDaoHelper.NzString(rs.Fields(FIELD_MODULE_CODE).Value))) > 0 Then
                        sourceObject = sourceObject & ":" & Trim$(modDaoHelper.NzString(rs.Fields(FIELD_MODULE_CODE).Value))
                    End If
                End If

                InsertAuditRow db, ResolveScopeCodeFromKey(translationKey), translationKey, languageCode, ResolveLegacyAuditStatus(translationKey), _
                               SOURCE_TRANSLATION, sourceObject, vbNullString, vbNullString, translationValue, True
                AppendOrphanAuditRows = AppendOrphanAuditRows + 1
            End If
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    modErrorHandler.HandleError MODULE_NAME, "AppendOrphanAuditRows", Err
End Function

Private Sub EnsureRegistryExpectedKeysFromActiveTranslations()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlStatement As String
    Dim translationKey As String
    Dim scopeCode As String

    EnsureTranslationExpectedTable

    Set db = CurrentDb
    If Not TableExists(db, TABLE_TRANSLATION) Then
        Exit Sub
    End If

    sqlStatement = "SELECT DISTINCT " & FIELD_TRANSLATION_KEY & " " & _
                   "FROM " & TABLE_TRANSLATION & " " & _
                   "WHERE Nz(" & FIELD_IS_ACTIVE & ", True) = True " & _
                   "AND (" & _
                   FIELD_TRANSLATION_KEY & " Like " & SqlText("MSG.*") & " " & _
                   "OR " & FIELD_TRANSLATION_KEY & " Like " & SqlText("MSG_*") & " " & _
                   "OR " & FIELD_TRANSLATION_KEY & " Like " & SqlText("STATUS.*") & " " & _
                   "OR " & FIELD_TRANSLATION_KEY & " Like " & SqlText("ERR_*") & " " & _
                   "OR " & FIELD_TRANSLATION_KEY & " Like " & SqlText("COMMON.*") & ");"

    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)
    Do While Not rs.EOF
        translationKey = NormalizeExpectedKey(modDaoHelper.NzString(rs.Fields(FIELD_TRANSLATION_KEY).Value))
        scopeCode = ResolveScopeCodeFromKey(translationKey)

        If LenB(translationKey) > 0 And _
           (StrComp(scopeCode, SCOPE_MSG, vbTextCompare) = 0 Or _
            StrComp(scopeCode, SCOPE_STATUS, vbTextCompare) = 0 Or _
            StrComp(scopeCode, SCOPE_COMMON, vbTextCompare) = 0) Then
            EnsureExpectedRegistryRow db, scopeCode, translationKey, TABLE_TRANSLATION, vbNullString
        End If
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    modErrorHandler.HandleError MODULE_NAME, "EnsureRegistryExpectedKeysFromActiveTranslations", Err
End Sub

Private Sub EnsureExpectedRegistryRow( _
    ByVal db As DAO.Database, _
    ByVal scopeCode As String, _
    ByVal translationKey As String, _
    ByVal sourceObject As String, _
    ByVal fallbackText As String)
    On Error GoTo ErrorHandler

    Dim criteria As String
    Dim sqlStatement As String
    Dim auditUser As String

    criteria = FIELD_TRANSLATION_KEY & " = " & SqlText(translationKey)
    auditUser = ResolveAuditUser()

    If DCount("*", TABLE_EXPECTED, criteria) > 0 Then
        Exit Sub
    End If

    sqlStatement = "INSERT INTO " & TABLE_EXPECTED & " (" & _
                   "scope_code, translation_key, source_object, fallback_text, is_active, " & _
                   "created_at, created_by, updated_at, updated_by) VALUES (" & _
                   SqlNullableText(scopeCode) & ", " & _
                   SqlText(translationKey) & ", " & _
                   SqlNullableText(sourceObject) & ", " & _
                   SqlNullableText(fallbackText) & ", True, Now(), " & _
                   SqlText(auditUser) & ", Now(), " & SqlText(auditUser) & ");"

    db.Execute sqlStatement, dbFailOnError
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureExpectedRegistryRow", Err
End Sub

Private Function TryGetTranslationValue( _
    ByVal db As DAO.Database, _
    ByVal translationKey As String, _
    ByVal languageCode As String, _
    ByRef translationValue As Variant) As Boolean
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlStatement As String

    If Not TableExists(db, TABLE_TRANSLATION) Then
        Exit Function
    End If

    sqlStatement = "SELECT TOP 1 " & FIELD_TRANSLATION_VALUE & " " & _
                   "FROM " & TABLE_TRANSLATION & " " & _
                   "WHERE " & FIELD_TRANSLATION_KEY & " = " & SqlText(translationKey) & " " & _
                   "AND " & FIELD_LANGUAGE_CODE & " = " & SqlText(languageCode) & ";"

    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then
        translationValue = rs.Fields(FIELD_TRANSLATION_VALUE).Value
        TryGetTranslationValue = True
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    modErrorHandler.HandleError MODULE_NAME, "TryGetTranslationValue", Err
End Function

Private Function TranslationRowExists(ByVal db As DAO.Database, ByVal translationKey As String, ByVal languageCode As String) As Boolean
    On Error GoTo ErrorHandler

    If Not TableExists(db, TABLE_TRANSLATION) Then
        Exit Function
    End If

    TranslationRowExists = (DCount("*", TABLE_TRANSLATION, _
        FIELD_TRANSLATION_KEY & " = " & SqlText(translationKey) & _
        " AND " & FIELD_LANGUAGE_CODE & " = " & SqlText(languageCode)) > 0)
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "TranslationRowExists", Err
End Function

Private Sub InsertMissingTranslationRow( _
    ByVal db As DAO.Database, _
    ByVal translationKey As String, _
    ByVal languageCode As String, _
    ByVal scopeCode As String)
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset

    Set rs = db.OpenRecordset(TABLE_TRANSLATION, dbOpenDynaset)
    rs.AddNew
    SetRecordsetFieldValue rs, FIELD_LANGUAGE_CODE, languageCode
    SetRecordsetFieldValue rs, FIELD_TRANSLATION_KEY, translationKey
    SetRecordsetFieldValue rs, FIELD_TRANSLATION_VALUE, Null
    SetRecordsetFieldValue rs, FIELD_IS_ACTIVE, True
    SetRecordsetFieldValue rs, FIELD_MODULE_CODE, scopeCode
    SetRecordsetFieldValue rs, FIELD_SORT_ORDER, 0
    SetRecordsetFieldValue rs, FIELD_CREATED_AT, Now()
    SetRecordsetFieldValue rs, FIELD_CREATED_BY, ResolveAuditUser()
    SetRecordsetFieldValue rs, FIELD_UPDATED_AT, Now()
    SetRecordsetFieldValue rs, FIELD_UPDATED_BY, ResolveAuditUser()
    rs.Update

    rs.Close
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.CancelUpdate
        rs.Close
    End If
    Set rs = Nothing
    modErrorHandler.HandleError MODULE_NAME, "InsertMissingTranslationRow", Err
End Sub

Private Sub InsertAuditRow( _
    ByVal db As DAO.Database, _
    ByVal scopeCode As String, _
    ByVal translationKey As String, _
    ByVal languageCode As String, _
    ByVal auditStatus As String, _
    ByVal sourceType As String, _
    ByVal sourceObject As String, _
    ByVal sourceControl As String, _
    ByVal fallbackText As String, _
    ByVal translationValue As Variant, _
    ByVal existsInTranslation As Boolean)

    Dim sqlStatement As String

    sqlStatement = "INSERT INTO " & TABLE_AUDIT & " (" & _
                   "scope_code, translation_key, language_code, audit_status, source_type, " & _
                   "source_object, source_control, fallback_text, translation_value, exists_in_fw_translation, created_at) " & _
                   "VALUES (" & _
                   SqlNullableText(scopeCode) & ", " & _
                   SqlText(translationKey) & ", " & _
                   SqlNullableText(languageCode) & ", " & _
                   SqlText(auditStatus) & ", " & _
                   SqlNullableText(sourceType) & ", " & _
                   SqlNullableText(sourceObject) & ", " & _
                   SqlNullableText(sourceControl) & ", " & _
                   SqlNullableText(fallbackText) & ", " & _
                   SqlLongText(translationValue) & ", " & _
                   IIf(existsInTranslation, "True", "False") & ", " & _
                   "Now());"

    db.Execute sqlStatement, dbFailOnError
End Sub

Private Function ResolveReadableFieldName( _
    ByVal rs As DAO.Recordset, _
    ByVal preferredNameField As String, _
    ByVal fallbackCodeFieldName As String) As String

    If LenB(preferredNameField) > 0 Then
        If modDaoHelper.RecordsetHasField(rs, preferredNameField) Then
            ResolveReadableFieldName = preferredNameField
            Exit Function
        End If
    End If

    If LenB(fallbackCodeFieldName) > 0 Then
        If modDaoHelper.RecordsetHasField(rs, fallbackCodeFieldName) Then
            ResolveReadableFieldName = fallbackCodeFieldName
        End If
    End If
End Function

Private Function ResolveRecordsetFallbackText( _
    ByVal rs As DAO.Recordset, _
    ByVal fieldName As String, _
    ByVal defaultValue As String) As String

    If LenB(fieldName) > 0 Then
        If modDaoHelper.RecordsetHasField(rs, fieldName) Then
            ResolveRecordsetFallbackText = Trim$(modDaoHelper.NzString(rs.Fields(fieldName).Value, defaultValue))
        End If
    End If

    If LenB(ResolveRecordsetFallbackText) = 0 Then
        ResolveRecordsetFallbackText = defaultValue
    End If
End Function

Private Function NormalizeExpectedKey(ByVal translationKey As String) As String
    translationKey = Trim$(modDaoHelper.NzString(translationKey))

    If LenB(translationKey) = 0 Then
        Exit Function
    End If

    If UCase$(Left$(translationKey, Len(TR_PREFIX))) = TR_PREFIX Then
        modLoggingHandler.LogWarning MODULE_NAME & ".NormalizeExpectedKey", _
            "Ignored invalid translation key with TR: prefix: " & translationKey
        Exit Function
    End If

    NormalizeExpectedKey = translationKey
End Function

Private Function RequiredLanguages() As Variant
    RequiredLanguages = Array("DE-CH", "EN-US", "FR-FR")
End Function

Private Function ResolveScopeCodeFromKey(ByVal translationKey As String) As String
    Dim normalizedKey As String

    normalizedKey = UCase$(Trim$(translationKey))

    If Left$(normalizedKey, 4) = "NAV." Then
        ResolveScopeCodeFromKey = SCOPE_NAV
    ElseIf Left$(normalizedKey, 5) = "FORM." Then
        ResolveScopeCodeFromKey = SCOPE_FORM
    ElseIf Left$(normalizedKey, 7) = "REPORT." Then
        ResolveScopeCodeFromKey = "REPORT"
    ElseIf Left$(normalizedKey, 4) = "MSG." Then
        ResolveScopeCodeFromKey = SCOPE_MSG
    ElseIf Left$(normalizedKey, 4) = "MSG_" Then
        ResolveScopeCodeFromKey = SCOPE_MSG
    ElseIf Left$(normalizedKey, 4) = "ERR_" Then
        ResolveScopeCodeFromKey = SCOPE_MSG
    ElseIf Left$(normalizedKey, 7) = "STATUS." Then
        ResolveScopeCodeFromKey = SCOPE_STATUS
    ElseIf Left$(normalizedKey, 7) = "COMMON." Then
        ResolveScopeCodeFromKey = SCOPE_COMMON
    ElseIf normalizedKey = "APP_TITLE" Or _
           normalizedKey = "CUSTOMER" Or _
           normalizedKey = "DOCUMENT" Or _
           normalizedKey = "TOTAL" Then
        ResolveScopeCodeFromKey = SCOPE_COMMON
    ElseIf Left$(normalizedKey, 9) = "DOCUMENT." Then
        ResolveScopeCodeFromKey = "DOCUMENT"
    ElseIf Left$(normalizedKey, 4) = "REF." Or _
           Left$(normalizedKey, 13) = "ADDRESS_TYPE." Or _
           Left$(normalizedKey, 11) = "SALUTATION." Or _
           Left$(normalizedKey, 13) = "CONTACT_TYPE." Or _
           Left$(normalizedKey, 5) = "UNIT." Or _
           Left$(normalizedKey, 4) = "VAT." Then
        ResolveScopeCodeFromKey = SCOPE_REF
    Else
        ResolveScopeCodeFromKey = "UNKNOWN"
    End If
End Function

Private Function ResolveLegacyAuditStatus(ByVal translationKey As String) As String
    If IsLegacyTranslationKey(translationKey) Then
        ResolveLegacyAuditStatus = STATUS_LEGACY_KEY
    Else
        ResolveLegacyAuditStatus = STATUS_ORPHAN
    End If
End Function

Private Function IsLegacyTranslationKey(ByVal translationKey As String) As Boolean
    Dim normalizedKey As String

    normalizedKey = UCase$(Trim$(translationKey))
    If LenB(normalizedKey) = 0 Then
        Exit Function
    End If

    If Left$(normalizedKey, 4) = "MSG_" Then
        IsLegacyTranslationKey = True
    ElseIf Left$(normalizedKey, 4) = "ERR_" Then
        IsLegacyTranslationKey = True
    ElseIf normalizedKey = "APP_TITLE" Or _
           normalizedKey = "CUSTOMER" Or _
           normalizedKey = "DOCUMENT" Or _
           normalizedKey = "TOTAL" Then
        IsLegacyTranslationKey = True
    End If
End Function

Private Function BuildExpectedKeyToken(ByVal translationKey As String) As String
    BuildExpectedKeyToken = UCase$(Trim$(translationKey))
End Function

Private Function AddExpectedKey(ByRef expectedKeys As Collection, ByVal translationKey As String) As Boolean
    Dim token As String

    token = BuildExpectedKeyToken(translationKey)
    If LenB(token) = 0 Then
        Exit Function
    End If

    If Not CollectionContains(expectedKeys, token) Then
        expectedKeys.Add translationKey, token
        AddExpectedKey = True
    End If
End Function

Private Function CollectionContains(ByVal items As Collection, ByVal itemKey As String) As Boolean
    On Error GoTo NotFound

    Dim value As Variant

    value = items.Item(itemKey)
    CollectionContains = True
    Exit Function

NotFound:
    CollectionContains = False
End Function

Private Sub SetRecordsetFieldValue(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal fieldValue As Variant)
    If modDaoHelper.RecordsetHasField(rs, fieldName) Then
        rs.Fields(fieldName).Value = fieldValue
    End If
End Sub

Private Function ResolveAuditUser() As String
    On Error Resume Next

    ResolveAuditUser = Trim$(modDaoHelper.NzString(modSessionContext.currentUserId))
    If LenB(ResolveAuditUser) = 0 Then
        ResolveAuditUser = Trim$(modDaoHelper.NzString(Environ$("Username")))
    End If
    If LenB(ResolveAuditUser) = 0 Then
        ResolveAuditUser = "SYSTEM"
    End If
End Function

Private Function SqlLongText(ByVal fieldValue As Variant) As String
    If IsNull(fieldValue) Or IsEmpty(fieldValue) Then
        SqlLongText = "Null"
    Else
        SqlLongText = SqlNullableText(modDaoHelper.NzString(fieldValue))
    End If
End Function

Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tableDefinition As DAO.TableDef

    For Each tableDefinition In db.TableDefs
        If StrComp(tableDefinition.Name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next tableDefinition
    Exit Function

ErrorHandler:
    TableExists = False
End Function

Private Function OpenFormHiddenDesignIfNeeded(ByVal formName As String, Optional ByVal wasAlreadyOpen As Boolean = False) As Boolean
    On Error GoTo ErrorHandler

    If Not wasAlreadyOpen And Not IsFormLoadedAnywhere(formName) Then
        DoCmd.OpenForm formName, acDesign, , , , acHidden
        OpenFormHiddenDesignIfNeeded = True
    End If
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "OpenFormHiddenDesignIfNeeded", Err
End Function

Private Sub CloseFormNoSave(ByVal formName As String)
    On Error Resume Next
    If IsProjectFormLoaded(formName) Then
        DoCmd.Close acForm, formName, acSaveNo
    End If
End Sub

Private Function ResolveScannableFormInstance(ByVal formName As String) As Access.Form
    On Error GoTo SafeExit

    If TryResolveLoadedFormInstance(formName, ResolveScannableFormInstance) Then
        Exit Function
    End If

    If IsProjectFormLoaded(formName) Then
        Set ResolveScannableFormInstance = Forms(formName)
    End If

SafeExit:
End Function

Private Function IsProjectFormLoaded(ByVal formName As String) As Boolean
    On Error GoTo SafeExit

    IsProjectFormLoaded = CurrentProject.AllForms(formName).IsLoaded
    If Not IsProjectFormLoaded Then
        IsProjectFormLoaded = IsFormLoadedAnywhere(formName)
    End If

SafeExit:
End Function

Private Function IsFormLoadedAnywhere(ByVal formName As String) As Boolean
    On Error GoTo SafeExit

    Dim openForm As Access.Form
    Dim resolvedForm As Access.Form

    For Each openForm In Forms
        Set resolvedForm = Nothing
        If TryFindLoadedFormInstance(openForm, formName, resolvedForm) Then
            IsFormLoadedAnywhere = True
            Exit Function
        End If
    Next openForm

SafeExit:
    Set openForm = Nothing
    Set resolvedForm = Nothing
End Function

Private Function TryResolveLoadedFormInstance(ByVal formName As String, ByRef resolvedForm As Access.Form) As Boolean
    On Error GoTo SafeExit

    Dim openForm As Access.Form

    For Each openForm In Forms
        If TryFindLoadedFormInstance(openForm, formName, resolvedForm) Then
            TryResolveLoadedFormInstance = True
            Exit Function
        End If
    Next openForm

SafeExit:
    Set openForm = Nothing
End Function

Private Function TryFindLoadedFormInstance( _
    ByVal currentForm As Access.Form, _
    ByVal targetFormName As String, _
    ByRef resolvedForm As Access.Form) As Boolean
    On Error GoTo SafeExit

    Dim ctl As Control
    Dim childForm As Access.Form

    If currentForm Is Nothing Then
        Exit Function
    End If

    If StrComp(currentForm.Name, targetFormName, vbTextCompare) = 0 Then
        Set resolvedForm = currentForm
        TryFindLoadedFormInstance = True
        Exit Function
    End If

    For Each ctl In currentForm.Controls
        If ctl.ControlType = acSubform Then
            Set childForm = Nothing
            On Error Resume Next
            Set childForm = ctl.Form
            On Error GoTo SafeExit

            If Not childForm Is Nothing Then
                If TryFindLoadedFormInstance(childForm, targetFormName, resolvedForm) Then
                    TryFindLoadedFormInstance = True
                    Exit Function
                End If
            End If
        End If
    Next ctl

SafeExit:
    Set childForm = Nothing
End Function

Private Function IsFormOpenByName(ByVal formName As String) As Boolean
    On Error GoTo SafeExit

    Dim openForm As Access.Form

    For Each openForm In Forms
        If StrComp(openForm.Name, formName, vbTextCompare) = 0 Then
            IsFormOpenByName = True
            Exit Function
        End If
    Next openForm

SafeExit:
    Set openForm = Nothing
End Function

Private Function GetTagTextSafely(ByVal tagValue As Variant) As String
    GetTagTextSafely = modDaoHelper.NzString(tagValue)
End Function

Private Function GetControlTagSafely(ByVal ctl As Control) As String
    On Error GoTo SafeExit
    GetControlTagSafely = modDaoHelper.NzString(ctl.Tag)
SafeExit:
End Function

Private Function NzCaptionTextSafely(ByVal formInstance As Access.Form, ByVal defaultValue As String) As String
    On Error GoTo SafeExit
    NzCaptionTextSafely = modDaoHelper.NzString(formInstance.Caption, defaultValue)
SafeExit:
    If LenB(NzCaptionTextSafely) = 0 Then
        NzCaptionTextSafely = defaultValue
    End If
End Function

Private Function GetControlFallbackText(ByVal ctl As Control, ByVal defaultValue As String) As String
    On Error GoTo SafeExit

    Select Case ctl.ControlType
        Case acLabel, acCommandButton, acCheckBox, acOptionButton, acToggleButton
            GetControlFallbackText = modDaoHelper.NzString(ctl.Caption, defaultValue)
        Case Else
            GetControlFallbackText = defaultValue
    End Select

SafeExit:
    If LenB(GetControlFallbackText) = 0 Then
        GetControlFallbackText = defaultValue
    End If
End Function
