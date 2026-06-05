Attribute VB_Name = "modMigrationPaymentTerms"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modMigrationPaymentTerms
' Purpose   : Applies schema and seed data migration for document language and
'             payment terms support.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modMigrationPaymentTerms"

Private Const TABLE_DOC_DOCUMENT As String = "doc_document"
Private Const TABLE_TEN_PAYMENT_TERM As String = "ten_payment_term"

Private Const FIELD_LANGUAGE_CODE As String = "language_code"
Private Const FIELD_PAYMENT_TERM_CODE As String = "payment_term_code"
Private Const FIELD_PAYMENT_TERMS_TEXT As String = "payment_terms_text"

Private Const FIELD_PAYMENT_TERM_ID As String = "payment_term_id"
Private Const FIELD_TITLE As String = "title"
Private Const FIELD_TERMS_TEXT As String = "terms_text"
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

Private Const INDEX_UQ_PAYMENT_TERM_CODE_LANGUAGE As String = "UQ_ten_payment_term_code_language"
Private Const INDEX_IS_DEFAULT As String = "IX_ten_payment_term_is_default"
Private Const INDEX_IS_ACTIVE As String = "IX_ten_payment_term_is_active"

Private Const DEFAULT_CREATED_BY As String = "migration"
Private Const DEFAULT_UPDATED_BY As String = "migration"

Public Sub ApplyPaymentTermsMigration()
    On Error GoTo ErrorHandler

    Dim frontendDb As DAO.Database
    Dim beDb As DAO.Database
    Dim backendPath As String

    Set frontendDb = CurrentDb
    backendPath = GetBackendPathForLinkedTable(frontendDb, TABLE_DOC_DOCUMENT)
    Set beDb = DBEngine.OpenDatabase(backendPath)

    EnsureDocDocumentFields beDb
    EnsureTenPaymentTermTable beDb
    EnsureTenPaymentTermIndexes beDb
    SeedDefaultPaymentTerms beDb
    EnsureLinkedBackendTable frontendDb, backendPath, TABLE_TEN_PAYMENT_TERM

    MsgBox "Payment terms migration completed. Backend: " & backendPath & _
           ". Linked table: " & TABLE_TEN_PAYMENT_TERM, vbInformation, MODULE_NAME
    GoTo CleanExit

ErrorHandler:
    MsgBox "Payment terms migration failed:" & vbCrLf & _
           Err.Number & " - " & Err.description & vbCrLf & vbCrLf & _
           "Backend: " & backendPath, vbCritical, MODULE_NAME

CleanExit:
    On Error Resume Next
    If Not beDb Is Nothing Then
        beDb.Close
        Set beDb = Nothing
    End If
    Set frontendDb = Nothing
    On Error GoTo 0
End Sub

Private Sub EnsureDocDocumentFields(ByVal db As DAO.Database)
    On Error GoTo ErrorHandler

    If db Is Nothing Then
        Exit Sub
    End If

    If Not TableExists(db, TABLE_DOC_DOCUMENT) Then
        Debug.Print MODULE_NAME & ".EnsureDocDocumentFields: Table '" & TABLE_DOC_DOCUMENT & "' not found."
        Exit Sub
    End If

    If Not FieldExists(db, TABLE_DOC_DOCUMENT, FIELD_LANGUAGE_CODE) Then
        ExecSql db, "ALTER TABLE " & TABLE_DOC_DOCUMENT & " ADD COLUMN " & FIELD_LANGUAGE_CODE & " TEXT(10)"
    End If

    If Not FieldExists(db, TABLE_DOC_DOCUMENT, FIELD_PAYMENT_TERM_CODE) Then
        ExecSql db, "ALTER TABLE " & TABLE_DOC_DOCUMENT & " ADD COLUMN " & FIELD_PAYMENT_TERM_CODE & " TEXT(50)"
    End If

    If Not FieldExists(db, TABLE_DOC_DOCUMENT, FIELD_PAYMENT_TERMS_TEXT) Then
        ExecSql db, "ALTER TABLE " & TABLE_DOC_DOCUMENT & " ADD COLUMN " & FIELD_PAYMENT_TERMS_TEXT & " LONGTEXT"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print MODULE_NAME & ".EnsureDocDocumentFields: " & Err.Number & " - " & Err.description
    Err.Clear
End Sub

Private Sub EnsureTenPaymentTermTable(ByVal db As DAO.Database)
    On Error GoTo ErrorHandler

    Dim sql As String

    If db Is Nothing Then
        Exit Sub
    End If

    If TableExists(db, TABLE_TEN_PAYMENT_TERM) Then
        Exit Sub
    End If

    sql = "CREATE TABLE " & TABLE_TEN_PAYMENT_TERM & " ("
    sql = sql & FIELD_PAYMENT_TERM_ID & " COUNTER CONSTRAINT PK_" & TABLE_TEN_PAYMENT_TERM & " PRIMARY KEY, "
    sql = sql & FIELD_PAYMENT_TERM_CODE & " TEXT(50) NOT NULL, "
    sql = sql & FIELD_LANGUAGE_CODE & " TEXT(10) NOT NULL, "
    sql = sql & FIELD_TITLE & " TEXT(100), "
    sql = sql & FIELD_TERMS_TEXT & " LONGTEXT, "
    sql = sql & FIELD_DAYS_NET & " INTEGER, "
    sql = sql & FIELD_DISCOUNT_DAYS & " INTEGER, "
    sql = sql & FIELD_DISCOUNT_PERCENT & " DOUBLE, "
    sql = sql & FIELD_IS_DEFAULT & " YESNO, "
    sql = sql & FIELD_IS_ACTIVE & " YESNO, "
    sql = sql & FIELD_SORT_ORDER & " INTEGER, "
    sql = sql & FIELD_CREATED_AT & " DATETIME, "
    sql = sql & FIELD_CREATED_BY & " TEXT(50), "
    sql = sql & FIELD_UPDATED_AT & " DATETIME, "
    sql = sql & FIELD_UPDATED_BY & " TEXT(50)"
    sql = sql & ")"

    ExecSql db, sql
    Exit Sub

ErrorHandler:
    Debug.Print MODULE_NAME & ".EnsureTenPaymentTermTable: " & Err.Number & " - " & Err.description
    Err.Clear
End Sub

Private Sub EnsureTenPaymentTermIndexes(ByVal db As DAO.Database)
    On Error GoTo ErrorHandler

    If db Is Nothing Then
        Exit Sub
    End If

    If Not TableExists(db, TABLE_TEN_PAYMENT_TERM) Then
        Exit Sub
    End If

    If Not IndexExists(db, TABLE_TEN_PAYMENT_TERM, INDEX_UQ_PAYMENT_TERM_CODE_LANGUAGE) Then
        ExecSql db, "CREATE UNIQUE INDEX " & INDEX_UQ_PAYMENT_TERM_CODE_LANGUAGE & _
                    " ON " & TABLE_TEN_PAYMENT_TERM & " (" & FIELD_PAYMENT_TERM_CODE & ", " & FIELD_LANGUAGE_CODE & ")"
    End If

    If Not IndexExists(db, TABLE_TEN_PAYMENT_TERM, INDEX_IS_DEFAULT) Then
        ExecSql db, "CREATE INDEX " & INDEX_IS_DEFAULT & _
                    " ON " & TABLE_TEN_PAYMENT_TERM & " (" & FIELD_IS_DEFAULT & ")"
    End If

    If Not IndexExists(db, TABLE_TEN_PAYMENT_TERM, INDEX_IS_ACTIVE) Then
        ExecSql db, "CREATE INDEX " & INDEX_IS_ACTIVE & _
                    " ON " & TABLE_TEN_PAYMENT_TERM & " (" & FIELD_IS_ACTIVE & ")"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print MODULE_NAME & ".EnsureTenPaymentTermIndexes: " & Err.Number & " - " & Err.description
    Err.Clear
End Sub

Private Sub SeedDefaultPaymentTerms(ByVal db As DAO.Database)
    On Error GoTo ErrorHandler

    If db Is Nothing Then
        Exit Sub
    End If

    If Not TableExists(db, TABLE_TEN_PAYMENT_TERM) Then
        Exit Sub
    End If

    EnsurePaymentTerm db, "PREPAYMENT", "de-CH", "Vorkasse", _
        "Zahlbar im Voraus.", Null, Null, Null, False, True, 10

    EnsurePaymentTerm db, "NET_30", "de-CH", "30 Tage netto", _
        "Zahlbar innert 30 Tagen netto.", 30, Null, Null, True, True, 20

    EnsurePaymentTerm db, "CASH_DISCOUNT_10_2_NET_30", "de-CH", "10 Tage -2% Skonto, 30 Tage netto", _
        "2% Skonto bei Zahlung innert 10 Tagen, ansonsten zahlbar innert 30 Tagen netto.", 30, 10, 2, False, True, 30

    EnsurePaymentTerm db, "SPLIT_50_ORDER_50_DELIVERY", "de-CH", "50% bei Auftragserteilung, 50% bei Lieferung", _
        "50% zahlbar bei Auftragserteilung, 50% zahlbar bei Lieferung.", Null, Null, Null, False, True, 40

    EnsurePaymentTerm db, "MILESTONE_50_25_25", "de-CH", "50% Anzahlung, 25% bei Launch, 25% nach 30 Tagen", _
        "50% Anzahlung bei Auftragserteilung, 25% bei Launch, 25% 30 Tage nach Launch.", Null, Null, Null, False, True, 50

    EnsurePaymentTerm db, "PREPAYMENT", "en", "Prepayment", _
        "Payable in advance.", Null, Null, Null, False, True, 110

    EnsurePaymentTerm db, "NET_30", "en", "Net 30 days", _
        "Payable within 30 days net.", 30, Null, Null, False, True, 120

    Exit Sub

ErrorHandler:
    Debug.Print MODULE_NAME & ".SeedDefaultPaymentTerms: " & Err.Number & " - " & Err.description
    Err.Clear
End Sub

Private Sub EnsurePaymentTerm( _
    ByVal db As DAO.Database, _
    ByVal PaymentTermCode As String, _
    ByVal LanguageCode As String, _
    ByVal Title As String, _
    ByVal TermsText As String, _
    ByVal DaysNet As Variant, _
    ByVal DiscountDays As Variant, _
    ByVal DiscountPercent As Variant, _
    ByVal IsDefault As Boolean, _
    ByVal isActive As Boolean, _
    ByVal sortOrder As Variant)
    On Error GoTo ErrorHandler

    Dim sql As String

    If PaymentTermExists(db, PaymentTermCode, LanguageCode) Then
        Exit Sub
    End If

    sql = "INSERT INTO " & TABLE_TEN_PAYMENT_TERM & " ("
    sql = sql & FIELD_PAYMENT_TERM_CODE & ", "
    sql = sql & FIELD_LANGUAGE_CODE & ", "
    sql = sql & FIELD_TITLE & ", "
    sql = sql & FIELD_TERMS_TEXT & ", "
    sql = sql & FIELD_DAYS_NET & ", "
    sql = sql & FIELD_DISCOUNT_DAYS & ", "
    sql = sql & FIELD_DISCOUNT_PERCENT & ", "
    sql = sql & FIELD_IS_DEFAULT & ", "
    sql = sql & FIELD_IS_ACTIVE & ", "
    sql = sql & FIELD_SORT_ORDER & ", "
    sql = sql & FIELD_CREATED_AT & ", "
    sql = sql & FIELD_CREATED_BY & ", "
    sql = sql & FIELD_UPDATED_AT & ", "
    sql = sql & FIELD_UPDATED_BY & ") VALUES ("
    sql = sql & SqlText(PaymentTermCode) & ", "
    sql = sql & SqlText(LanguageCode) & ", "
    sql = sql & SqlText(Title) & ", "
    sql = sql & SqlText(TermsText) & ", "
    sql = sql & SqlNumber(DaysNet) & ", "
    sql = sql & SqlNumber(DiscountDays) & ", "
    sql = sql & SqlNumber(DiscountPercent) & ", "
    sql = sql & SqlBool(IsDefault) & ", "
    sql = sql & SqlBool(isActive) & ", "
    sql = sql & SqlNumber(sortOrder) & ", "
    sql = sql & SqlDate(Now()) & ", "
    sql = sql & SqlText(DEFAULT_CREATED_BY) & ", "
    sql = sql & SqlDate(Now()) & ", "
    sql = sql & SqlText(DEFAULT_UPDATED_BY) & ")"

    ExecSql db, sql
    Exit Sub

ErrorHandler:
    Debug.Print MODULE_NAME & ".EnsurePaymentTerm: " & Err.Number & " - " & Err.description
    Err.Clear
End Sub

Private Function PaymentTermExists( _
    ByVal db As DAO.Database, _
    ByVal PaymentTermCode As String, _
    ByVal LanguageCode As String) As Boolean
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sql As String

    PaymentTermExists = False

    If db Is Nothing Then
        Exit Function
    End If

    If Not TableExists(db, TABLE_TEN_PAYMENT_TERM) Then
        Exit Function
    End If

    sql = "SELECT " & FIELD_PAYMENT_TERM_ID & _
          " FROM " & TABLE_TEN_PAYMENT_TERM & _
          " WHERE " & FIELD_PAYMENT_TERM_CODE & " = " & SqlText(PaymentTermCode) & _
          " AND " & FIELD_LANGUAGE_CODE & " = " & SqlText(LanguageCode)

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    PaymentTermExists = Not (rs.BOF And rs.EOF)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    On Error GoTo 0
    Exit Function

ErrorHandler:
    PaymentTermExists = False
    Debug.Print MODULE_NAME & ".PaymentTermExists: " & Err.Number & " - " & Err.description
    Resume CleanExit
End Function

Private Function GetBackendPathForLinkedTable(ByVal db As DAO.Database, ByVal linkedTableName As String) As String
    On Error GoTo ErrorHandler

    Dim connectText As String
    Dim databaseMarker As String
    Dim markerPosition As Long

    databaseMarker = ";DATABASE="

    If db Is Nothing Then
        Err.Raise vbObjectError + 2000, MODULE_NAME & ".GetBackendPathForLinkedTable", _
            "Database reference is not available."
    End If

    If Not TableExists(db, linkedTableName) Then
        Err.Raise vbObjectError + 2001, MODULE_NAME & ".GetBackendPathForLinkedTable", _
            "Linked table '" & linkedTableName & "' was not found."
    End If

    connectText = Nz(db.TableDefs(linkedTableName).Connect, vbNullString)
    If LenB(Trim$(connectText)) = 0 Then
        Err.Raise vbObjectError + 2002, MODULE_NAME & ".GetBackendPathForLinkedTable", _
            "Table '" & linkedTableName & "' is not linked."
    End If

    markerPosition = InStr(1, connectText, databaseMarker, vbTextCompare)
    If markerPosition <= 0 Then
        Err.Raise vbObjectError + 2003, MODULE_NAME & ".GetBackendPathForLinkedTable", _
            "Could not determine backend path for linked table '" & linkedTableName & "'."
    End If

    GetBackendPathForLinkedTable = Trim$(Mid$(connectText, markerPosition + Len(databaseMarker)))
    If LenB(GetBackendPathForLinkedTable) = 0 Then
        Err.Raise vbObjectError + 2004, MODULE_NAME & ".GetBackendPathForLinkedTable", _
            "Backend path for linked table '" & linkedTableName & "' is empty."
    End If

    Exit Function

ErrorHandler:
    If Err.Number >= vbObjectError + 2000 And Err.Number <= vbObjectError + 2004 Then
        Err.Raise Err.Number, Err.Source, Err.description
    End If

    Err.Raise vbObjectError + 2005, MODULE_NAME & ".GetBackendPathForLinkedTable", _
        "Failed to resolve backend path for linked table '" & linkedTableName & "': " & Err.description
End Function

Private Sub EnsureLinkedBackendTable( _
    ByVal frontendDb As DAO.Database, _
    ByVal backendPath As String, _
    ByVal tableName As String)
    On Error GoTo ErrorHandler

    Dim tdf As DAO.tableDef
    Dim existingTdf As DAO.tableDef
    Dim existingConnect As String

    If frontendDb Is Nothing Then
        Exit Sub
    End If

    If LenB(Trim$(backendPath)) = 0 Then
        Err.Raise vbObjectError + 2010, MODULE_NAME & ".EnsureLinkedBackendTable", _
            "Backend path is empty for linked table '" & tableName & "'."
    End If

    If TableExists(frontendDb, tableName) Then
        Set existingTdf = frontendDb.TableDefs(tableName)
        existingConnect = Trim$(Nz(existingTdf.Connect, vbNullString))

        If LenB(existingConnect) = 0 Then
            Err.Raise vbObjectError + 2011, MODULE_NAME & ".EnsureLinkedBackendTable", _
                "Local table exists in frontend and cannot be replaced automatically: " & tableName
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
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

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

Private Function IndexExists(ByVal db As DAO.Database, ByVal tableName As String, ByVal indexName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tdf As DAO.tableDef
    Dim idx As DAO.index

    If db Is Nothing Then
        Exit Function
    End If

    If Not TableExists(db, tableName) Then
        Exit Function
    End If

    Set tdf = db.TableDefs(tableName)

    For Each idx In tdf.Indexes
        If StrComp(idx.Name, indexName, vbTextCompare) = 0 Then
            IndexExists = True
            Exit Function
        End If
    Next idx

    Exit Function

ErrorHandler:
    IndexExists = False
End Function

Private Sub ExecSql(ByVal db As DAO.Database, ByVal sql As String)
    Debug.Print sql
    db.Execute sql, dbFailOnError
End Sub

Private Function SqlDate(ByVal v As Variant) As String
    Dim dt As Date

    If IsNull(v) Then
        SqlDate = "NULL"
    Else
        dt = CDate(v)
        SqlDate = "#" & Year(dt) & "-" & Right$("0" & Month(dt), 2) & "-" & Right$("0" & Day(dt), 2) & _
                  " " & Right$("0" & Hour(dt), 2) & ":" & Right$("0" & Minute(dt), 2) & ":" & Right$("0" & Second(dt), 2) & "#"
    End If
End Function

Private Function SqlBool(ByVal v As Variant) As String
    If IsNull(v) Then
        SqlBool = "NULL"
    ElseIf CBool(v) Then
        SqlBool = "True"
    Else
        SqlBool = "False"
    End If
End Function

Private Function SqlNumber(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        SqlNumber = "NULL"
    Else
        SqlNumber = Replace(Trim$(Str$(CDbl(v))), ",", ".")
    End If
End Function
