Attribute VB_Name = "modMigrationPaymentTerms"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modMigrationPaymentTerms
' Purpose   : Consolidates ten_payment_term to one canonical business row per
'             payment term and enforces the stage-2 schema without legacy text fields.
' Author    : Codex
' Version   : 0.2.0
'===============================================================================

Private Const MODULE_NAME As String = "modMigrationPaymentTerms"

Private Const TABLE_DOC_DOCUMENT As String = "doc_document"
Private Const TABLE_FW_TRANSLATION As String = "fw_translation"
Private Const TABLE_ORD_ORDER As String = "ord_order"
Private Const TABLE_TEN_PAYMENT_TERM As String = "ten_payment_term"
Private Const TABLE_TMP_ORDER As String = "tmp_order"

Private Const FIELD_PAYMENT_TERM_ID As String = "payment_term_id"
Private Const FIELD_PAYMENT_TERM_CODE As String = "payment_term_code"
Private Const FIELD_PAYMENT_TERM_TYPE_CODE As String = "payment_term_type_code"
Private Const FIELD_PAYMENT_TERMS_TEXT As String = "payment_terms_text"
Private Const FIELD_LANGUAGE_CODE As String = "language_code"
Private Const FIELD_DAYS_NET As String = "days_net"
Private Const FIELD_DISCOUNT_DAYS As String = "discount_days"
Private Const FIELD_DISCOUNT_PERCENT As String = "discount_percent"
Private Const FIELD_IS_DEFAULT As String = "is_default"
Private Const FIELD_IS_ACTIVE As String = "is_active"
Private Const FIELD_SORT_ORDER As String = "sort_order"
Private Const FIELD_CREATED_AT As String = "created_at"
Private Const FIELD_CREATED_BY As String = "created_by"
Private Const FIELD_UPDATED_AT As String = "updated_at"
Private Const FIELD_UPDATED_BY As String = "updated_by"
Private Const FIELD_TRANSLATION_KEY As String = "translation_key"
Private Const FIELD_TRANSLATION_VALUE As String = "translation_value"
Private Const FIELD_MODULE_CODE As String = "module_code"

Private Const INDEX_UQ_PAYMENT_TERM_CODE As String = "ux_ten_payment_term_code"
Private Const LEGACY_INDEX_UQ_PAYMENT_TERM_CODE_LANGUAGE As String = "ux_ten_payment_term_code_language"
Private Const INDEX_IS_DEFAULT As String = "ix_ten_payment_term_is_default"
Private Const INDEX_IS_ACTIVE As String = "ix_ten_payment_term_is_active"

Private Const MODULE_CODE_REFERENCE As String = "REF"
Private Const DEFAULT_CREATED_BY As String = "SYSTEM"
Private Const DEFAULT_UPDATED_BY As String = "SYSTEM"

Private Const CODE_PREPAYMENT As String = "PREPAYMENT"
Private Const CODE_NET_30 As String = "NET_30"
Private Const CODE_NET30_LEGACY As String = "NET30"
Private Const CODE_CASH_DISCOUNT As String = "CASH_DISCOUNT_10_2_NET_30"
Private Const CODE_CASH_DISCOUNT_LEGACY As String = "2D10N30"
Private Const CODE_SPLIT_50_50 As String = "SPLIT_50_ORDER_50_DELIVERY"
Private Const CODE_MILESTONE_50_25_25 As String = "MILESTONE_50_25_25"

Private mInsertedTranslationCount As Long
Private mSkippedTranslationCount As Long
Private mConflictTranslationCount As Long
Private mUpdatedReferenceCount As Long
Private mDeletedDuplicateCount As Long
Private mInsertedCanonicalRowCount As Long

Public Function ApplyPaymentTermsMigration() As Boolean
    On Error GoTo ErrorHandler

    Dim frontendDb As DAO.Database
    Dim backendDb As DAO.Database
    Dim backendPath As String
    Dim workspace As DAO.Workspace
    Dim transactionStarted As Boolean
    Dim tenPaymentTermLinked As Boolean

    ApplyPaymentTermsMigration = False
    ResetMigrationCounters

    Set frontendDb = CurrentDb
    backendPath = ResolveBusinessBackendPath(frontendDb)
    tenPaymentTermLinked = IsLinkedAccessTable(frontendDb, TABLE_TEN_PAYMENT_TERM)

    If tenPaymentTermLinked And LenB(backendPath) = 0 Then
        Err.Raise vbObjectError + 6520, MODULE_NAME & ".ApplyPaymentTermsMigration", _
            "Linked table '" & TABLE_TEN_PAYMENT_TERM & "' has no resolvable backend path."
    End If

    If LenB(backendPath) > 0 And StrComp(NormalizePath(frontendDb.Name), NormalizePath(backendPath), vbTextCompare) <> 0 Then
        Set backendDb = DBEngine.OpenDatabase(backendPath)
    Else
        Set backendDb = frontendDb
        backendPath = frontendDb.Name
    End If

    LogMigrationExecutionPath frontendDb, backendDb, backendPath, tenPaymentTermLinked

    EnsureDocDocumentFields backendDb
    EnsureTenPaymentTermTable backendDb
    EnsureTenPaymentTermStage2Schema backendDb

    Set workspace = DBEngine.Workspaces(0)
    workspace.BeginTrans
    transactionStarted = True

    MigrateLegacyPaymentTermCodeReferences backendDb
    EnsureCanonicalPaymentTermTranslationSeeds frontendDb
    ConsolidateTenPaymentTerms backendDb

    workspace.CommitTrans
    transactionStarted = False

    DBEngine.Idle dbFreeLocks

    EnsureTenPaymentTermIndexes backendDb
    RemoveTenPaymentTermLegacyFields backendDb
    VerifyTenPaymentTermStage2Cleanup backendDb

    EnsureLinkedBackendTable frontendDb, backendPath, TABLE_TEN_PAYMENT_TERM

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyPaymentTermsMigration", _
        "Payment-term migration completed. backend_path=" & backendPath & _
        "; inserted_translations=" & CStr(mInsertedTranslationCount) & _
        "; skipped_translations=" & CStr(mSkippedTranslationCount) & _
        "; conflicting_translations=" & CStr(mConflictTranslationCount) & _
        "; updated_references=" & CStr(mUpdatedReferenceCount) & _
        "; deleted_duplicates=" & CStr(mDeletedDuplicateCount) & _
        "; inserted_canonical_rows=" & CStr(mInsertedCanonicalRowCount) & "."

    ApplyPaymentTermsMigration = True
    GoTo CleanExit

ErrorHandler:
    On Error Resume Next
    If transactionStarted Then
        workspace.Rollback
    End If
    On Error GoTo 0
    modErrorHandler.HandleError MODULE_NAME, "ApplyPaymentTermsMigration", Err

CleanExit:
    On Error Resume Next
    If Not backendDb Is Nothing Then
        If StrComp(NormalizePath(backendDb.Name), NormalizePath(frontendDb.Name), vbTextCompare) <> 0 Then
            backendDb.Close
        End If
    End If
    Set backendDb = Nothing
    Set frontendDb = Nothing
End Function

Private Sub VerifyTenPaymentTermStage2Cleanup(ByVal db As DAO.Database)
    On Error GoTo ErrorHandler

    If db Is Nothing Then
        Err.Raise vbObjectError + 6521, MODULE_NAME & ".VerifyTenPaymentTermStage2Cleanup", _
            "Schema verification database is not available."
    End If

    If modDbSchema.FieldExists(db, TABLE_TEN_PAYMENT_TERM, "language_code") Then
        Err.Raise vbObjectError + 6522, MODULE_NAME & ".VerifyTenPaymentTermStage2Cleanup", _
            "Legacy field ten_payment_term.language_code still exists after cleanup."
    End If

    If modDbSchema.FieldExists(db, TABLE_TEN_PAYMENT_TERM, "title") Then
        Err.Raise vbObjectError + 6523, MODULE_NAME & ".VerifyTenPaymentTermStage2Cleanup", _
            "Legacy field ten_payment_term.title still exists after cleanup."
    End If

    If modDbSchema.FieldExists(db, TABLE_TEN_PAYMENT_TERM, "terms_text") Then
        Err.Raise vbObjectError + 6524, MODULE_NAME & ".VerifyTenPaymentTermStage2Cleanup", _
            "Legacy field ten_payment_term.terms_text still exists after cleanup."
    End If

    If Not IndexExists(db, TABLE_TEN_PAYMENT_TERM, INDEX_UQ_PAYMENT_TERM_CODE) Then
        Err.Raise vbObjectError + 6525, MODULE_NAME & ".VerifyTenPaymentTermStage2Cleanup", _
            "Required unique index ux_ten_payment_term_code is missing after cleanup."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "VerifyTenPaymentTermStage2Cleanup", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub EnsureDocDocumentFields(ByVal db As DAO.Database)
    On Error GoTo ErrorHandler

    If db Is Nothing Then
        Exit Sub
    End If

    If Not modDbSchema.TableExists(db, TABLE_DOC_DOCUMENT) Then
        Exit Sub
    End If

    If Not modDbSchema.FieldExists(db, TABLE_DOC_DOCUMENT, FIELD_LANGUAGE_CODE) Then
        ExecSql db, "ALTER TABLE [" & TABLE_DOC_DOCUMENT & "] ADD COLUMN [" & FIELD_LANGUAGE_CODE & "] TEXT(10);"
    End If

    If Not modDbSchema.FieldExists(db, TABLE_DOC_DOCUMENT, FIELD_PAYMENT_TERM_CODE) Then
        ExecSql db, "ALTER TABLE [" & TABLE_DOC_DOCUMENT & "] ADD COLUMN [" & FIELD_PAYMENT_TERM_CODE & "] TEXT(50);"
    End If

    If Not modDbSchema.FieldExists(db, TABLE_DOC_DOCUMENT, FIELD_PAYMENT_TERMS_TEXT) Then
        ExecSql db, "ALTER TABLE [" & TABLE_DOC_DOCUMENT & "] ADD COLUMN [" & FIELD_PAYMENT_TERMS_TEXT & "] LONGTEXT;"
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureDocDocumentFields", Err
End Sub

Private Sub EnsureTenPaymentTermTable(ByVal db As DAO.Database)
    On Error GoTo ErrorHandler

    Dim sqlText As String

    If db Is Nothing Then
        Exit Sub
    End If

    If modDbSchema.TableExists(db, TABLE_TEN_PAYMENT_TERM) Then
        Exit Sub
    End If

    sqlText = "CREATE TABLE [" & TABLE_TEN_PAYMENT_TERM & "] (" & _
              "[" & FIELD_PAYMENT_TERM_ID & "] COUNTER CONSTRAINT [pk_ten_payment_term] PRIMARY KEY, " & _
              "[" & FIELD_PAYMENT_TERM_CODE & "] TEXT(50) NOT NULL, " & _
              "[" & FIELD_PAYMENT_TERM_TYPE_CODE & "] TEXT(50), " & _
              "[" & FIELD_DAYS_NET & "] LONG, " & _
              "[" & FIELD_DISCOUNT_DAYS & "] LONG, " & _
              "[" & FIELD_DISCOUNT_PERCENT & "] DOUBLE, " & _
              "[" & FIELD_IS_DEFAULT & "] YESNO, " & _
              "[" & FIELD_IS_ACTIVE & "] YESNO, " & _
              "[" & FIELD_SORT_ORDER & "] LONG, " & _
              "[" & FIELD_CREATED_AT & "] DATETIME, " & _
              "[" & FIELD_CREATED_BY & "] TEXT(50), " & _
              "[" & FIELD_UPDATED_AT & "] DATETIME, " & _
              "[" & FIELD_UPDATED_BY & "] TEXT(50));"

    ExecSql db, sqlText
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureTenPaymentTermTable", Err
End Sub

Private Sub EnsureTenPaymentTermStage2Schema(ByVal db As DAO.Database)
    On Error GoTo ErrorHandler

    If db Is Nothing Then
        Exit Sub
    End If

    EnsureTextField db, FIELD_PAYMENT_TERM_CODE, 50
    EnsureTextField db, FIELD_PAYMENT_TERM_TYPE_CODE, 50
    EnsureLongField db, FIELD_DAYS_NET
    EnsureLongField db, FIELD_DISCOUNT_DAYS
    EnsureDoubleField db, FIELD_DISCOUNT_PERCENT
    EnsureYesNoField db, FIELD_IS_DEFAULT
    EnsureYesNoField db, FIELD_IS_ACTIVE
    EnsureLongField db, FIELD_SORT_ORDER
    EnsureDateField db, FIELD_CREATED_AT
    EnsureTextField db, FIELD_CREATED_BY, 50
    EnsureDateField db, FIELD_UPDATED_AT
    EnsureTextField db, FIELD_UPDATED_BY, 50
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureTenPaymentTermStage2Schema", Err
End Sub

Private Sub MigrateLegacyPaymentTermCodeReferences(ByVal db As DAO.Database)
    UpdatePaymentTermCodeReferenceTable db, TABLE_ORD_ORDER, FIELD_PAYMENT_TERM_CODE
    UpdatePaymentTermCodeReferenceTable db, TABLE_TMP_ORDER, FIELD_PAYMENT_TERM_CODE
    UpdatePaymentTermCodeReferenceTable db, TABLE_DOC_DOCUMENT, FIELD_PAYMENT_TERM_CODE
End Sub

Private Sub UpdatePaymentTermCodeReferenceTable( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal fieldName As String)
    On Error GoTo ErrorHandler

    If db Is Nothing Then
        Exit Sub
    End If

    If Not modDbSchema.TableExists(db, tableName) Then
        Exit Sub
    End If

    If Not modDbSchema.FieldExists(db, tableName, fieldName) Then
        Exit Sub
    End If

    mUpdatedReferenceCount = mUpdatedReferenceCount + _
        ExecuteReferenceCodeUpdate(db, tableName, fieldName, CODE_NET30_LEGACY, CODE_NET_30)
    mUpdatedReferenceCount = mUpdatedReferenceCount + _
        ExecuteReferenceCodeUpdate(db, tableName, fieldName, CODE_CASH_DISCOUNT_LEGACY, CODE_CASH_DISCOUNT)
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "UpdatePaymentTermCodeReferenceTable", Err
End Sub

Private Function ExecuteReferenceCodeUpdate( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal fieldName As String, _
    ByVal oldCode As String, _
    ByVal newCode As String) As Long
    On Error GoTo ErrorHandler

    Dim sqlStatement As String

    sqlStatement = "UPDATE [" & tableName & "] " & _
                   "SET [" & fieldName & "]=" & SqlText(newCode) & " " & _
                   "WHERE UCase(Trim(Nz([" & fieldName & "], '')))=" & SqlText(UCase$(oldCode)) & ";"

    db.Execute sqlStatement, dbFailOnError
    ExecuteReferenceCodeUpdate = db.RecordsAffected
    Exit Function

ErrorHandler:
    ExecuteReferenceCodeUpdate = 0
    modErrorHandler.HandleError MODULE_NAME, "ExecuteReferenceCodeUpdate", Err
End Function

Private Sub EnsureCanonicalPaymentTermTranslationSeeds(ByVal translationDb As DAO.Database)
    If translationDb Is Nothing Then
        Exit Sub
    End If

    EnsurePaymentTermTranslation translationDb, BuildPaymentTermTranslationKey(CODE_PREPAYMENT, "TITLE"), "de-CH", "Vorkasse", 101
    EnsurePaymentTermTranslation translationDb, BuildPaymentTermTranslationKey(CODE_PREPAYMENT, "TERMS"), "de-CH", "Zahlbar im Voraus.", 102
    EnsurePaymentTermTranslation translationDb, BuildPaymentTermTranslationKey(CODE_PREPAYMENT, "TITLE"), "en-US", "Prepayment", 103
    EnsurePaymentTermTranslation translationDb, BuildPaymentTermTranslationKey(CODE_PREPAYMENT, "TERMS"), "en-US", "Payable in advance.", 104

    EnsurePaymentTermTranslation translationDb, BuildPaymentTermTranslationKey(CODE_NET_30, "TITLE"), "de-CH", "30 Tage netto", 201
    EnsurePaymentTermTranslation translationDb, BuildPaymentTermTranslationKey(CODE_NET_30, "TERMS"), "de-CH", "Zahlbar innert 30 Tagen netto.", 202
    EnsurePaymentTermTranslation translationDb, BuildPaymentTermTranslationKey(CODE_NET_30, "TITLE"), "en-US", "Net 30 days", 203
    EnsurePaymentTermTranslation translationDb, BuildPaymentTermTranslationKey(CODE_NET_30, "TERMS"), "en-US", "Payable within 30 days net.", 204

    EnsurePaymentTermTranslation translationDb, BuildPaymentTermTranslationKey(CODE_CASH_DISCOUNT, "TITLE"), "de-CH", "10 Tage -2% Skonto, 30 Tage netto", 301
    EnsurePaymentTermTranslation translationDb, BuildPaymentTermTranslationKey(CODE_CASH_DISCOUNT, "TERMS"), "de-CH", "2% Skonto bei Zahlung innert 10 Tagen, ansonsten zahlbar innert 30 Tagen netto.", 302

    EnsurePaymentTermTranslation translationDb, BuildPaymentTermTranslationKey(CODE_SPLIT_50_50, "TITLE"), "de-CH", "50% bei Auftragserteilung, 50% bei Lieferung", 401
    EnsurePaymentTermTranslation translationDb, BuildPaymentTermTranslationKey(CODE_SPLIT_50_50, "TERMS"), "de-CH", "50% zahlbar bei Auftragserteilung, 50% zahlbar bei Lieferung.", 402

    EnsurePaymentTermTranslation translationDb, BuildPaymentTermTranslationKey(CODE_MILESTONE_50_25_25, "TITLE"), "de-CH", "50% Anzahlung, 25% bei Launch, 25% nach 30 Tagen", 501
    EnsurePaymentTermTranslation translationDb, BuildPaymentTermTranslationKey(CODE_MILESTONE_50_25_25, "TERMS"), "de-CH", "50% Anzahlung bei Auftragserteilung, 25% bei Launch, 25% 30 Tage nach Launch.", 502
End Sub

Private Sub EnsurePaymentTermTranslation( _
    ByVal db As DAO.Database, _
    ByVal translationKey As String, _
    ByVal languageCode As String, _
    ByVal translationValue As String, _
    ByVal sortOrder As Long)
    On Error GoTo ErrorHandler

    Dim normalizedLanguageCode As String
    Dim existingValue As String

    normalizedLanguageCode = NormalizeSupportedLanguageCode(languageCode)
    If LenB(normalizedLanguageCode) = 0 Then
        Exit Sub
    End If

    existingValue = GetExistingTranslationValue(db, translationKey, normalizedLanguageCode)
    If LenB(existingValue) > 0 Then
        If StrComp(Trim$(existingValue), Trim$(translationValue), vbTextCompare) <> 0 Then
            mConflictTranslationCount = mConflictTranslationCount + 1
            modLoggingHandler.LogWarning MODULE_NAME & ".EnsurePaymentTermTranslation", _
                "Skipped conflicting translation for key='" & translationKey & "'; language_code='" & normalizedLanguageCode & "'."
        Else
            mSkippedTranslationCount = mSkippedTranslationCount + 1
        End If
        Exit Sub
    End If

    If modFwSetup.TranslationSeedExists(db, normalizedLanguageCode, translationKey) Then
        mSkippedTranslationCount = mSkippedTranslationCount + 1
        Exit Sub
    End If

    modFwSetup.EnsureTranslationSeed db, normalizedLanguageCode, translationKey, translationValue, MODULE_CODE_REFERENCE, sortOrder
    mInsertedTranslationCount = mInsertedTranslationCount + 1
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsurePaymentTermTranslation", Err
End Sub

Private Sub ConsolidateTenPaymentTerms(ByVal db As DAO.Database)
    On Error GoTo ErrorHandler

    ConsolidatePaymentTermFamily db, CODE_PREPAYMENT
    ConsolidatePaymentTermFamily db, CODE_NET_30
    ConsolidatePaymentTermFamily db, CODE_CASH_DISCOUNT
    ConsolidatePaymentTermFamily db, CODE_SPLIT_50_50
    ConsolidatePaymentTermFamily db, CODE_MILESTONE_50_25_25
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConsolidateTenPaymentTerms", Err
End Sub

Private Sub ConsolidatePaymentTermFamily(ByVal db As DAO.Database, ByVal canonicalCode As String)
    On Error GoTo ErrorHandler

    Dim survivorId As Long
    Dim isActiveValue As Boolean

    survivorId = ResolvePaymentTermSurvivorId(db, canonicalCode)
    isActiveValue = ResolveCanonicalIsActive(db, canonicalCode)

    If survivorId <= 0 Then
        survivorId = InsertCanonicalPaymentTermRow(db, canonicalCode, isActiveValue)
    End If

    NormalizeSurvivingPaymentTermRow db, survivorId, canonicalCode, isActiveValue
    DeleteNonSurvivingFamilyRows db, canonicalCode, survivorId
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConsolidatePaymentTermFamily", Err
End Sub

Private Function ResolvePaymentTermSurvivorId(ByVal db As DAO.Database, ByVal canonicalCode As String) As Long
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim rowCode As String
    Dim currentScore As Long
    Dim bestScore As Long
    Dim rowId As Long

    Set rs = db.OpenRecordset( _
        "SELECT * FROM [" & TABLE_TEN_PAYMENT_TERM & "] ORDER BY [" & FIELD_PAYMENT_TERM_ID & "];", _
        dbOpenSnapshot)

    Do While Not rs.EOF
        rowCode = Trim$(modDaoHelper.NzString(rs.Fields(FIELD_PAYMENT_TERM_CODE).Value, vbNullString))
        If StrComp(ResolveCanonicalPaymentTermCode(rowCode), canonicalCode, vbTextCompare) = 0 Then
            currentScore = CalculatePaymentTermSurvivorScore(rs, canonicalCode)
            rowId = modDaoHelper.NzLong(rs.Fields(FIELD_PAYMENT_TERM_ID).Value, 0)
            If currentScore > bestScore Or (currentScore = bestScore And (ResolvePaymentTermSurvivorId = 0 Or rowId < ResolvePaymentTermSurvivorId)) Then
                bestScore = currentScore
                ResolvePaymentTermSurvivorId = rowId
            End If
        End If
        rs.MoveNext
    Loop

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ResolvePaymentTermSurvivorId", Err
    Resume CleanExit
End Function

Private Function CalculatePaymentTermSurvivorScore(ByVal rs As DAO.Recordset, ByVal canonicalCode As String) As Long
    Dim rowCode As String

    rowCode = Trim$(modDaoHelper.NzString(rs.Fields(FIELD_PAYMENT_TERM_CODE).Value, vbNullString))

    If StrComp(rowCode, canonicalCode, vbTextCompare) = 0 Then
        CalculatePaymentTermSurvivorScore = CalculatePaymentTermSurvivorScore + 100
    ElseIf LenB(ResolveCanonicalPaymentTermCode(rowCode)) > 0 Then
        CalculatePaymentTermSurvivorScore = CalculatePaymentTermSurvivorScore + 40
    End If

    If modDaoHelper.NzBoolean(rs.Fields(FIELD_IS_DEFAULT).Value, False) Then
        CalculatePaymentTermSurvivorScore = CalculatePaymentTermSurvivorScore + 30
    End If

    If modDaoHelper.NzBoolean(rs.Fields(FIELD_IS_ACTIVE).Value, False) Then
        CalculatePaymentTermSurvivorScore = CalculatePaymentTermSurvivorScore + 10
    End If
End Function

Private Function ResolveCanonicalIsActive(ByVal db As DAO.Database, ByVal canonicalCode As String) As Boolean
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim rowCode As String

    ResolveCanonicalIsActive = True

    Set rs = db.OpenRecordset( _
        "SELECT [" & FIELD_PAYMENT_TERM_CODE & "], [" & FIELD_IS_ACTIVE & "] FROM [" & TABLE_TEN_PAYMENT_TERM & "];", _
        dbOpenSnapshot)

    Do While Not rs.EOF
        rowCode = Trim$(modDaoHelper.NzString(rs.Fields(FIELD_PAYMENT_TERM_CODE).Value, vbNullString))
        If StrComp(ResolveCanonicalPaymentTermCode(rowCode), canonicalCode, vbTextCompare) = 0 Then
            If modDaoHelper.NzBoolean(rs.Fields(FIELD_IS_ACTIVE).Value, False) Then
                ResolveCanonicalIsActive = True
                GoTo CleanExit
            Else
                ResolveCanonicalIsActive = False
            End If
        End If
        rs.MoveNext
    Loop

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    ResolveCanonicalIsActive = True
    modErrorHandler.HandleError MODULE_NAME, "ResolveCanonicalIsActive", Err
    Resume CleanExit
End Function

Private Function InsertCanonicalPaymentTermRow( _
    ByVal db As DAO.Database, _
    ByVal canonicalCode As String, _
    ByVal isActiveValue As Boolean) As Long
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset

    Set rs = db.OpenRecordset(TABLE_TEN_PAYMENT_TERM, dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs.Fields(FIELD_PAYMENT_TERM_CODE).Value = canonicalCode
    If modDaoHelper.RecordsetHasField(rs, FIELD_PAYMENT_TERM_TYPE_CODE) Then
        rs.Fields(FIELD_PAYMENT_TERM_TYPE_CODE).Value = ResolvePaymentTermTypeCode(canonicalCode)
    End If
    If modDaoHelper.RecordsetHasField(rs, FIELD_IS_DEFAULT) Then
        rs.Fields(FIELD_IS_DEFAULT).Value = (StrComp(canonicalCode, CODE_NET_30, vbTextCompare) = 0)
    End If
    If modDaoHelper.RecordsetHasField(rs, FIELD_IS_ACTIVE) Then
        rs.Fields(FIELD_IS_ACTIVE).Value = isActiveValue
    End If
    If modDaoHelper.RecordsetHasField(rs, FIELD_SORT_ORDER) Then
        rs.Fields(FIELD_SORT_ORDER).Value = ResolvePaymentTermSortOrder(canonicalCode)
    End If
    ApplyCanonicalBusinessDefaults rs, canonicalCode, True
    SetCreatedAuditFields rs
    SetUpdatedAuditFields rs
    rs.Update

    rs.Bookmark = rs.LastModified
    InsertCanonicalPaymentTermRow = modDaoHelper.NzLong(rs.Fields(FIELD_PAYMENT_TERM_ID).Value, 0)
    mInsertedCanonicalRowCount = mInsertedCanonicalRowCount + 1

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "InsertCanonicalPaymentTermRow", Err
    Resume CleanExit
End Function

Private Sub NormalizeSurvivingPaymentTermRow( _
    ByVal db As DAO.Database, _
    ByVal survivorId As Long, _
    ByVal canonicalCode As String, _
    ByVal isActiveValue As Boolean)
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset

    Set rs = db.OpenRecordset( _
        "SELECT * FROM [" & TABLE_TEN_PAYMENT_TERM & "] WHERE [" & FIELD_PAYMENT_TERM_ID & "]=" & CStr(survivorId) & ";", _
        dbOpenDynaset)

    If rs.BOF And rs.EOF Then
        GoTo CleanExit
    End If

    rs.Edit
    rs.Fields(FIELD_PAYMENT_TERM_CODE).Value = canonicalCode
    If modDaoHelper.RecordsetHasField(rs, FIELD_PAYMENT_TERM_TYPE_CODE) Then
        rs.Fields(FIELD_PAYMENT_TERM_TYPE_CODE).Value = ResolvePaymentTermTypeCode(canonicalCode)
    End If
    If modDaoHelper.RecordsetHasField(rs, FIELD_IS_DEFAULT) Then
        rs.Fields(FIELD_IS_DEFAULT).Value = (StrComp(canonicalCode, CODE_NET_30, vbTextCompare) = 0)
    End If
    If modDaoHelper.RecordsetHasField(rs, FIELD_IS_ACTIVE) Then
        rs.Fields(FIELD_IS_ACTIVE).Value = isActiveValue
    End If
    If modDaoHelper.RecordsetHasField(rs, FIELD_SORT_ORDER) Then
        rs.Fields(FIELD_SORT_ORDER).Value = ResolvePaymentTermSortOrder(canonicalCode)
    End If
    ApplyCanonicalBusinessDefaults rs, canonicalCode, False
    EnsureCreatedAuditFields rs
    SetUpdatedAuditFields rs
    rs.Update

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NormalizeSurvivingPaymentTermRow", Err
    Resume CleanExit
End Sub

Private Sub DeleteNonSurvivingFamilyRows( _
    ByVal db As DAO.Database, _
    ByVal canonicalCode As String, _
    ByVal survivorId As Long)
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim rowCode As String
    Dim rowId As Long

    Set rs = db.OpenRecordset( _
        "SELECT [" & FIELD_PAYMENT_TERM_ID & "], [" & FIELD_PAYMENT_TERM_CODE & "] FROM [" & TABLE_TEN_PAYMENT_TERM & "] ORDER BY [" & FIELD_PAYMENT_TERM_ID & "];", _
        dbOpenDynaset)

    Do While Not rs.EOF
        rowCode = Trim$(modDaoHelper.NzString(rs.Fields(FIELD_PAYMENT_TERM_CODE).Value, vbNullString))
        rowId = modDaoHelper.NzLong(rs.Fields(FIELD_PAYMENT_TERM_ID).Value, 0)

        If rowId <> survivorId Then
            If StrComp(ResolveCanonicalPaymentTermCode(rowCode), canonicalCode, vbTextCompare) = 0 Then
                rs.Delete
                mDeletedDuplicateCount = mDeletedDuplicateCount + 1
            End If
        End If
        rs.MoveNext
    Loop

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "DeleteNonSurvivingFamilyRows", Err
    Resume CleanExit
End Sub

Private Sub ApplyCanonicalBusinessDefaults( _
    ByVal rs As DAO.Recordset, _
    ByVal canonicalCode As String, _
    ByVal forceDefaults As Boolean)

    Select Case UCase$(canonicalCode)
        Case CODE_NET_30
            SetNumericFieldIfMissing rs, FIELD_DAYS_NET, 30, forceDefaults
            SetNumericFieldIfMissing rs, FIELD_DISCOUNT_DAYS, Null, forceDefaults
            SetNumericFieldIfMissing rs, FIELD_DISCOUNT_PERCENT, Null, forceDefaults

        Case CODE_CASH_DISCOUNT
            SetNumericFieldIfMissing rs, FIELD_DAYS_NET, 30, forceDefaults
            SetNumericFieldIfMissing rs, FIELD_DISCOUNT_DAYS, 10, forceDefaults
            SetNumericFieldIfMissing rs, FIELD_DISCOUNT_PERCENT, 2, forceDefaults

        Case Else
            If forceDefaults Then
                SetNumericFieldIfMissing rs, FIELD_DAYS_NET, Null, True
                SetNumericFieldIfMissing rs, FIELD_DISCOUNT_DAYS, Null, True
                SetNumericFieldIfMissing rs, FIELD_DISCOUNT_PERCENT, Null, True
            End If
    End Select
End Sub

Private Sub SetNumericFieldIfMissing( _
    ByVal rs As DAO.Recordset, _
    ByVal fieldName As String, _
    ByVal fieldValue As Variant, _
    ByVal forceValue As Boolean)

    If Not modDaoHelper.RecordsetHasField(rs, fieldName) Then
        Exit Sub
    End If

    If forceValue Then
        rs.Fields(fieldName).Value = fieldValue
    ElseIf IsNull(rs.Fields(fieldName).Value) Or LenB(Trim$(modDaoHelper.NzString(rs.Fields(fieldName).Value, vbNullString))) = 0 Then
        rs.Fields(fieldName).Value = fieldValue
    End If
End Sub

Private Sub EnsureCreatedAuditFields(ByVal rs As DAO.Recordset)
    If Not modDaoHelper.RecordsetHasField(rs, FIELD_CREATED_AT) Then
        Exit Sub
    End If

    If IsNull(rs.Fields(FIELD_CREATED_AT).Value) Then
        SetCreatedAuditFields rs
    ElseIf modDaoHelper.RecordsetHasField(rs, FIELD_CREATED_BY) Then
        If LenB(Trim$(modDaoHelper.NzString(rs.Fields(FIELD_CREATED_BY).Value, vbNullString))) = 0 Then
            rs.Fields(FIELD_CREATED_BY).Value = ResolveAuditUser()
        End If
    End If
End Sub

Private Sub EnsureTenPaymentTermIndexes(ByVal db As DAO.Database)
    On Error GoTo ErrorHandler

    DropIndexIfExists db, TABLE_TEN_PAYMENT_TERM, LEGACY_INDEX_UQ_PAYMENT_TERM_CODE_LANGUAGE
    ExecuteCreateIndexIfMissing db, TABLE_TEN_PAYMENT_TERM, INDEX_UQ_PAYMENT_TERM_CODE, _
        "CREATE UNIQUE INDEX [" & INDEX_UQ_PAYMENT_TERM_CODE & "] ON [" & TABLE_TEN_PAYMENT_TERM & "] ([" & FIELD_PAYMENT_TERM_CODE & "]);"
    ExecuteCreateIndexIfMissing db, TABLE_TEN_PAYMENT_TERM, INDEX_IS_DEFAULT, _
        "CREATE INDEX [" & INDEX_IS_DEFAULT & "] ON [" & TABLE_TEN_PAYMENT_TERM & "] ([" & FIELD_IS_DEFAULT & "]);"
    ExecuteCreateIndexIfMissing db, TABLE_TEN_PAYMENT_TERM, INDEX_IS_ACTIVE, _
        "CREATE INDEX [" & INDEX_IS_ACTIVE & "] ON [" & TABLE_TEN_PAYMENT_TERM & "] ([" & FIELD_IS_ACTIVE & "]);"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureTenPaymentTermIndexes", Err
End Sub

Private Sub RemoveTenPaymentTermLegacyFields(ByVal db As DAO.Database)
    On Error GoTo ErrorHandler

    If db Is Nothing Then
        Exit Sub
    End If

    DropIndexIfExists db, TABLE_TEN_PAYMENT_TERM, LEGACY_INDEX_UQ_PAYMENT_TERM_CODE_LANGUAGE
    DropIndexesForField db, TABLE_TEN_PAYMENT_TERM, "language_code"
    modBasicModuleSchema.DropFieldIfExists db, TABLE_TEN_PAYMENT_TERM, "language_code"
    modBasicModuleSchema.DropFieldIfExists db, TABLE_TEN_PAYMENT_TERM, "title"
    modBasicModuleSchema.DropFieldIfExists db, TABLE_TEN_PAYMENT_TERM, "terms_text"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RemoveTenPaymentTermLegacyFields", Err
End Sub

Private Function ResolveBusinessBackendPath(ByVal db As DAO.Database) As String
    If db Is Nothing Then
        Exit Function
    End If

    ResolveBusinessBackendPath = GetBackendPathForLinkedTable(db, TABLE_TEN_PAYMENT_TERM)
    If LenB(ResolveBusinessBackendPath) = 0 Then
        ResolveBusinessBackendPath = GetBackendPathForLinkedTable(db, TABLE_DOC_DOCUMENT)
    End If
    If LenB(ResolveBusinessBackendPath) = 0 Then
        ResolveBusinessBackendPath = GetBackendPathForLinkedTable(db, TABLE_ORD_ORDER)
    End If
End Function

Private Sub LogMigrationExecutionPath( _
    ByVal frontendDb As DAO.Database, _
    ByVal backendDb As DAO.Database, _
    ByVal backendPath As String, _
    ByVal tenPaymentTermLinked As Boolean)

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyPaymentTermsMigration", _
        "frontend_db=" & ResolveDatabaseName(frontendDb) & _
        "; table=" & TABLE_TEN_PAYMENT_TERM & _
        "; table_linked=" & CStr(tenPaymentTermLinked) & _
        "; resolved_backend_path=" & backendPath & _
        "; ddl_target_db=" & ResolveDatabaseName(backendDb) & _
        "; alter_table_runs_on_backend=" & CStr(StrComp(NormalizePath(ResolveDatabaseName(backendDb)), NormalizePath(backendPath), vbTextCompare) = 0)
End Sub

Private Function ResolveCanonicalPaymentTermCode(ByVal paymentTermCode As String) As String
    paymentTermCode = UCase$(Trim$(paymentTermCode))

    Select Case paymentTermCode
        Case CODE_PREPAYMENT
            ResolveCanonicalPaymentTermCode = CODE_PREPAYMENT
        Case CODE_NET_30, CODE_NET30_LEGACY
            ResolveCanonicalPaymentTermCode = CODE_NET_30
        Case CODE_CASH_DISCOUNT, CODE_CASH_DISCOUNT_LEGACY
            ResolveCanonicalPaymentTermCode = CODE_CASH_DISCOUNT
        Case CODE_SPLIT_50_50
            ResolveCanonicalPaymentTermCode = CODE_SPLIT_50_50
        Case CODE_MILESTONE_50_25_25
            ResolveCanonicalPaymentTermCode = CODE_MILESTONE_50_25_25
    End Select
End Function

Private Function ResolvePaymentTermTypeCode(ByVal canonicalCode As String) As String
    Select Case UCase$(Trim$(canonicalCode))
        Case CODE_PREPAYMENT
            ResolvePaymentTermTypeCode = "PREPAYMENT"
        Case CODE_NET_30
            ResolvePaymentTermTypeCode = "NET"
        Case CODE_CASH_DISCOUNT
            ResolvePaymentTermTypeCode = "CASH_DISCOUNT"
        Case CODE_SPLIT_50_50
            ResolvePaymentTermTypeCode = "INSTALLMENT"
        Case CODE_MILESTONE_50_25_25
            ResolvePaymentTermTypeCode = "MILESTONE"
    End Select
End Function

Private Function ResolvePaymentTermSortOrder(ByVal canonicalCode As String) As Long
    Select Case UCase$(Trim$(canonicalCode))
        Case CODE_PREPAYMENT
            ResolvePaymentTermSortOrder = 10
        Case CODE_NET_30
            ResolvePaymentTermSortOrder = 20
        Case CODE_CASH_DISCOUNT
            ResolvePaymentTermSortOrder = 30
        Case CODE_SPLIT_50_50
            ResolvePaymentTermSortOrder = 40
        Case CODE_MILESTONE_50_25_25
            ResolvePaymentTermSortOrder = 50
    End Select
End Function

Private Function BuildPaymentTermTranslationKey(ByVal paymentTermCode As String, ByVal suffixText As String) As String
    BuildPaymentTermTranslationKey = "PAYMENT_TERM." & UCase$(Trim$(paymentTermCode)) & "." & UCase$(Trim$(suffixText))
End Function

Private Function NormalizeSupportedLanguageCode(ByVal languageCode As String) As String
    languageCode = Trim$(modFwTranslationRuntime.NormalizeProjectLanguageCode(languageCode))
    If modFwTranslationRuntime.IsSupportedTranslationLanguage(languageCode) Then
        NormalizeSupportedLanguageCode = languageCode
    End If
End Function

Private Function GetExistingTranslationValue( _
    ByVal db As DAO.Database, _
    ByVal translationKey As String, _
    ByVal languageCode As String) As String
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlStatement As String

    sqlStatement = "SELECT TOP 1 [" & FIELD_TRANSLATION_VALUE & "] " & _
                   "FROM [" & TABLE_FW_TRANSLATION & "] " & _
                   "WHERE [" & FIELD_TRANSLATION_KEY & "]=" & SqlText(translationKey) & " " & _
                   "AND [" & FIELD_LANGUAGE_CODE & "]=" & SqlText(languageCode) & ";"

    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then
        GetExistingTranslationValue = Trim$(modDaoHelper.NzString(rs.Fields(FIELD_TRANSLATION_VALUE).Value, vbNullString))
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetExistingTranslationValue", Err
    Resume CleanExit
End Function

Private Function GetBackendPathForLinkedTable(ByVal db As DAO.Database, ByVal tableName As String) As String
    On Error GoTo ErrorHandler

    Dim connectText As String
    Dim markerPosition As Long
    Const DATABASE_MARKER As String = ";DATABASE="

    If db Is Nothing Then
        Exit Function
    End If

    If Not modDbSchema.TableExists(db, tableName) Then
        Exit Function
    End If

    connectText = Trim$(modDaoHelper.NzString(db.TableDefs(tableName).Connect, vbNullString))
    markerPosition = InStr(1, connectText, DATABASE_MARKER, vbTextCompare)
    If markerPosition <= 0 Then
        Exit Function
    End If

    GetBackendPathForLinkedTable = Trim$(Mid$(connectText, markerPosition + Len(DATABASE_MARKER)))
    Exit Function

ErrorHandler:
    GetBackendPathForLinkedTable = vbNullString
End Function

Private Function IsLinkedAccessTable(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim connectText As String
    Const DATABASE_MARKER As String = ";DATABASE="

    If db Is Nothing Then
        Exit Function
    End If

    If Not modDbSchema.TableExists(db, tableName) Then
        Exit Function
    End If

    connectText = Trim$(modDaoHelper.NzString(db.TableDefs(tableName).Connect, vbNullString))
    IsLinkedAccessTable = (InStr(1, connectText, DATABASE_MARKER, vbTextCompare) > 0)
    Exit Function

ErrorHandler:
    IsLinkedAccessTable = False
End Function

Private Sub EnsureLinkedBackendTable( _
    ByVal frontendDb As DAO.Database, _
    ByVal backendPath As String, _
    ByVal tableName As String)
    On Error GoTo ErrorHandler

    Dim tdf As DAO.TableDef
    Dim currentBackendPath As String

    If frontendDb Is Nothing Then
        Exit Sub
    End If

    If LenB(Trim$(backendPath)) = 0 Then
        Exit Sub
    End If

    If StrComp(NormalizePath(frontendDb.Name), NormalizePath(backendPath), vbTextCompare) = 0 Then
        Exit Sub
    End If

    currentBackendPath = GetBackendPathForLinkedTable(frontendDb, tableName)
    If LenB(currentBackendPath) > 0 Then
        If StrComp(NormalizePath(currentBackendPath), NormalizePath(backendPath), vbTextCompare) = 0 Then
            frontendDb.TableDefs(tableName).RefreshLink
            Exit Sub
        End If
    End If

    If modDbSchema.TableExists(frontendDb, tableName) Then
        If LenB(Trim$(modDaoHelper.NzString(frontendDb.TableDefs(tableName).Connect, vbNullString))) = 0 Then
            Err.Raise vbObjectError + 2500, MODULE_NAME & ".EnsureLinkedBackendTable", _
                "Local table exists and cannot be replaced automatically: " & tableName
        End If

        frontendDb.TableDefs.Delete tableName
        frontendDb.TableDefs.Refresh
    End If

    Set tdf = frontendDb.CreateTableDef(tableName)
    tdf.Connect = ";DATABASE=" & backendPath
    tdf.SourceTableName = tableName
    frontendDb.TableDefs.Append tdf
    frontendDb.TableDefs.Refresh
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureLinkedBackendTable", Err
End Sub

Private Sub DropIndexIfExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal indexName As String)
    On Error GoTo ErrorHandler

    If db Is Nothing Then
        Exit Sub
    End If

    If Not modDbSchema.TableExists(db, tableName) Then
        Exit Sub
    End If

    If Not IndexExists(db, tableName, indexName) Then
        Exit Sub
    End If

    db.TableDefs(tableName).Indexes.Delete indexName
    db.TableDefs(tableName).Indexes.Refresh
    Exit Sub

ErrorHandler:
    modLoggingHandler.LogWarning MODULE_NAME & ".DropIndexIfExists", _
        "Could not drop index '" & indexName & "' from table '" & tableName & "'."
End Sub

Private Sub DropIndexesForField(ByVal db As DAO.Database, ByVal tableName As String, ByVal fieldName As String)
    On Error GoTo ErrorHandler

    Dim tdf As DAO.TableDef
    Dim indexNames As Collection
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim indexName As Variant

    If db Is Nothing Then
        Exit Sub
    End If

    If Not modDbSchema.TableExists(db, tableName) Then
        Exit Sub
    End If

    Set tdf = db.TableDefs(tableName)
    Set indexNames = New Collection

    For Each idx In tdf.Indexes
        If Not idx.Primary Then
            For Each fld In idx.Fields
                If StrComp(Trim$(modDaoHelper.NzString(fld.Name, vbNullString)), fieldName, vbTextCompare) = 0 Then
                    indexNames.Add idx.Name
                    Exit For
                End If
            Next fld
        End If
    Next idx

    For Each indexName In indexNames
        tdf.Indexes.Delete CStr(indexName)
        modLoggingHandler.LogInfo MODULE_NAME & ".DropIndexesForField", _
            "Dropped index '" & CStr(indexName) & "' for field '" & fieldName & "' on table '" & tableName & "'."
    Next indexName
    tdf.Indexes.Refresh
    Exit Sub

ErrorHandler:
    modLoggingHandler.LogWarning MODULE_NAME & ".DropIndexesForField", _
        "Could not drop all indexes for field '" & fieldName & "' on table '" & tableName & "' (" & Err.Number & " - " & Err.Description & ")."
End Sub

Private Function IndexExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal indexName As String) As Boolean
    On Error GoTo SafeExit

    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index

    If db Is Nothing Then
        Exit Function
    End If

    If Not modDbSchema.TableExists(db, tableName) Then
        Exit Function
    End If

    Set tdf = db.TableDefs(tableName)
    For Each idx In tdf.Indexes
        If StrComp(idx.Name, indexName, vbTextCompare) = 0 Then
            IndexExists = True
            Exit Function
        End If
    Next idx

SafeExit:
    Set idx = Nothing
    Set tdf = Nothing
End Function

Private Sub ExecuteCreateIndexIfMissing( _
    ByVal db As DAO.Database, _
    ByVal tableName As String, _
    ByVal indexName As String, _
    ByVal sqlText As String)
    On Error GoTo ErrorHandler

    If IndexExists(db, tableName, indexName) Then
        Exit Sub
    End If

    ExecSql db, sqlText
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ExecuteCreateIndexIfMissing", Err
End Sub

Private Sub EnsureTextField(ByVal db As DAO.Database, ByVal fieldName As String, ByVal fieldSize As Long)
    If Not modDbSchema.FieldExists(db, TABLE_TEN_PAYMENT_TERM, fieldName) Then
        ExecSql db, "ALTER TABLE [" & TABLE_TEN_PAYMENT_TERM & "] ADD COLUMN [" & fieldName & "] TEXT(" & CStr(fieldSize) & ");"
    End If
End Sub

Private Sub EnsureLongTextField(ByVal db As DAO.Database, ByVal fieldName As String)
    If Not modDbSchema.FieldExists(db, TABLE_TEN_PAYMENT_TERM, fieldName) Then
        ExecSql db, "ALTER TABLE [" & TABLE_TEN_PAYMENT_TERM & "] ADD COLUMN [" & fieldName & "] LONGTEXT;"
    End If
End Sub

Private Sub EnsureLongField(ByVal db As DAO.Database, ByVal fieldName As String)
    If Not modDbSchema.FieldExists(db, TABLE_TEN_PAYMENT_TERM, fieldName) Then
        ExecSql db, "ALTER TABLE [" & TABLE_TEN_PAYMENT_TERM & "] ADD COLUMN [" & fieldName & "] LONG;"
    End If
End Sub

Private Sub EnsureDoubleField(ByVal db As DAO.Database, ByVal fieldName As String)
    If Not modDbSchema.FieldExists(db, TABLE_TEN_PAYMENT_TERM, fieldName) Then
        ExecSql db, "ALTER TABLE [" & TABLE_TEN_PAYMENT_TERM & "] ADD COLUMN [" & fieldName & "] DOUBLE;"
    End If
End Sub

Private Sub EnsureYesNoField(ByVal db As DAO.Database, ByVal fieldName As String)
    If Not modDbSchema.FieldExists(db, TABLE_TEN_PAYMENT_TERM, fieldName) Then
        ExecSql db, "ALTER TABLE [" & TABLE_TEN_PAYMENT_TERM & "] ADD COLUMN [" & fieldName & "] YESNO;"
    End If
End Sub

Private Sub EnsureDateField(ByVal db As DAO.Database, ByVal fieldName As String)
    If Not modDbSchema.FieldExists(db, TABLE_TEN_PAYMENT_TERM, fieldName) Then
        ExecSql db, "ALTER TABLE [" & TABLE_TEN_PAYMENT_TERM & "] ADD COLUMN [" & fieldName & "] DATETIME;"
    End If
End Sub

Private Sub ExecSql(ByVal db As DAO.Database, ByVal sqlText As String)
    db.Execute sqlText, dbFailOnError
End Sub

Private Function SqlText(ByVal valueText As String) As String
    SqlText = "'" & Replace(Trim$(valueText), "'", "''") & "'"
End Function

Private Function NormalizePath(ByVal pathText As String) As String
    NormalizePath = LCase$(Trim$(Replace(pathText, "/", "\")))
End Function

Private Function ResolveDatabaseName(ByVal db As DAO.Database) As String
    If db Is Nothing Then
        ResolveDatabaseName = "<nothing>"
    Else
        ResolveDatabaseName = Trim$(modDaoHelper.NzString(db.Name, "<unnamed>"))
    End If
End Function

Private Sub ResetMigrationCounters()
    mInsertedTranslationCount = 0
    mSkippedTranslationCount = 0
    mConflictTranslationCount = 0
    mUpdatedReferenceCount = 0
    mDeletedDuplicateCount = 0
    mInsertedCanonicalRowCount = 0
End Sub

Private Function ResolveAuditUser() As String
    ResolveAuditUser = Trim$(modDaoHelper.NzString(modSessionContext.currentUserId, vbNullString))
    If LenB(ResolveAuditUser) = 0 Then
        ResolveAuditUser = Trim$(modDaoHelper.NzString(modSessionContext.CurrentUserName, vbNullString))
    End If
    If LenB(ResolveAuditUser) = 0 Then
        ResolveAuditUser = DEFAULT_UPDATED_BY
    End If
End Function

Private Sub SetCreatedAuditFields(ByVal rs As DAO.Recordset)
    If modDaoHelper.RecordsetHasField(rs, FIELD_CREATED_AT) Then
        rs.Fields(FIELD_CREATED_AT).Value = Now()
    End If
    If modDaoHelper.RecordsetHasField(rs, FIELD_CREATED_BY) Then
        rs.Fields(FIELD_CREATED_BY).Value = ResolveAuditUser()
    End If
End Sub

Private Sub SetUpdatedAuditFields(ByVal rs As DAO.Recordset)
    If modDaoHelper.RecordsetHasField(rs, FIELD_UPDATED_AT) Then
        rs.Fields(FIELD_UPDATED_AT).Value = Now()
    End If
    If modDaoHelper.RecordsetHasField(rs, FIELD_UPDATED_BY) Then
        rs.Fields(FIELD_UPDATED_BY).Value = ResolveAuditUser()
    End If
End Sub
