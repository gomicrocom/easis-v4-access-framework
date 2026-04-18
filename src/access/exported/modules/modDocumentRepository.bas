Option Compare Database
Option Explicit

'===============================================================================
' Module    : modDocumentRepository
' Purpose   : DAO persistence helpers for document headers and document lines.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modDocumentRepository"

Private Const TABLE_DOC_DOCUMENT As String = "doc_document"
Private Const TABLE_DOC_DOCUMENT_POSITION As String = "doc_document_position"

' doc_document
Private Const FIELD_DOCUMENT_ID As String = "document_id"
Private Const FIELD_DOCUMENT_TYPE_CODE As String = "document_type_code"
Private Const FIELD_DOCUMENT_STATUS_CODE As String = "document_status_code"
Private Const FIELD_DOCUMENT_NO As String = "document_no"
Private Const FIELD_DOCUMENT_DATE As String = "document_date"
Private Const FIELD_CUSTOMER_NAME As String = "customer_name"
Private Const FIELD_CURRENCY_CODE As String = "currency_code"
Private Const FIELD_VAT_MODE As String = "vat_mode"
Private Const FIELD_VAT_RATE As String = "vat_rate"
Private Const FIELD_TOTAL_NET As String = "total_net"
Private Const FIELD_TOTAL_VAT As String = "total_vat"
Private Const FIELD_TOTAL_GROSS As String = "total_gross"
Private Const FIELD_REMARKS As String = "remarks"
Private Const FIELD_CREATED_AT As String = "created_at"
Private Const FIELD_CREATED_BY As String = "created_by"

' doc_document_position
Private Const FIELD_DOCUMENT_POSITION_ID As String = "document_position_id"
Private Const FIELD_LINE_NO As String = "line_no"
Private Const FIELD_DESCRIPTION As String = "description"
Private Const FIELD_QUANTITY As String = "quantity"
Private Const FIELD_UNIT_CODE As String = "unit_code"
Private Const FIELD_UNIT_PRICE As String = "unit_price"
Private Const FIELD_LINE_TOTAL_NET As String = "line_total_net"
Private Const FIELD_LINE_TOTAL_VAT As String = "line_total_vat"
Private Const FIELD_LINE_TOTAL_GROSS As String = "line_total_gross"

Private Const DEFAULT_DOCUMENT_STATUS As String = "DRAFT"

Public Function CreateDocumentHeader( _
    ByVal DocumentTypeCode As String, _
    Optional ByVal DocumentDate As Date = 0, _
    Optional ByVal CustomerName As String = "", _
    Optional ByVal TotalNet As Currency = 0, _
    Optional ByVal TotalVat As Currency = 0, _
    Optional ByVal TotalGross As Currency = 0, _
    Optional ByVal Remarks As String = "" _
) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim effectiveDate As Date
    Dim documentType As String

    CreateDocumentHeader = 0

    documentType = UCase$(Trim$(DocumentTypeCode))
    If Not modDocumentService.ValidateDocumentType(documentType) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CreateDocumentHeader", _
            "Document header creation skipped because document type is not supported."
        Exit Function
    End If

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CreateDocumentHeader", _
            "Document header creation skipped because backend configuration is not valid."
        Exit Function
    End If

    If Not TableExists(TABLE_DOC_DOCUMENT) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CreateDocumentHeader", _
            "Table '" & TABLE_DOC_DOCUMENT & "' is not available yet."
        Exit Function
    End If

    effectiveDate = IIf(DocumentDate = 0, Date, DocumentDate)

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset(TABLE_DOC_DOCUMENT, dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    SetRecordsetValue rs, FIELD_DOCUMENT_TYPE_CODE, documentType
    SetRecordsetValue rs, FIELD_DOCUMENT_STATUS_CODE, DEFAULT_DOCUMENT_STATUS
    SetRecordsetValue rs, FIELD_DOCUMENT_NO, vbNullString
    SetRecordsetValue rs, FIELD_DOCUMENT_DATE, effectiveDate
    SetRecordsetValue rs, FIELD_CUSTOMER_NAME, Trim$(CustomerName)
    SetRecordsetValue rs, FIELD_CURRENCY_CODE, ResolveCurrencyCode()
    SetRecordsetValue rs, FIELD_VAT_MODE, modVatHandler.GetVatMode()
    SetRecordsetValue rs, FIELD_VAT_RATE, modVatHandler.GetVatRate()
    SetRecordsetValue rs, FIELD_TOTAL_NET, TotalNet
    SetRecordsetValue rs, FIELD_TOTAL_VAT, TotalVat
    SetRecordsetValue rs, FIELD_TOTAL_GROSS, TotalGross
    SetRecordsetValue rs, FIELD_REMARKS, Trim$(Remarks)
    SetRecordsetValue rs, FIELD_CREATED_AT, Now()
    SetRecordsetValue rs, FIELD_CREATED_BY, ResolveCreatedBy()
    rs.Update

    rs.Bookmark = rs.LastModified
    If modDaoHelper.RecordsetHasField(rs, FIELD_DOCUMENT_ID) Then
        CreateDocumentHeader = modDaoHelper.NzLong(rs.Fields(FIELD_DOCUMENT_ID).Value, 0)
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".CreateDocumentHeader", _
        "Document header created. DocumentId=" & CStr(CreateDocumentHeader) & _
        ", Type='" & documentType & "'."

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    CreateDocumentHeader = 0
    modErrorHandler.HandleError MODULE_NAME, "CreateDocumentHeader", Err
    Resume CleanExit
End Function

Public Function DeleteDocumentPositions(ByVal DocumentId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim sqlText As String

    DeleteDocumentPositions = False

    If DocumentId <= 0 Then
        Exit Function
    End If

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".DeleteDocumentPositions", _
            "Document position delete skipped because backend configuration is not valid."
        Exit Function
    End If

    If Not TableExists(TABLE_DOC_DOCUMENT_POSITION) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".DeleteDocumentPositions", _
            "Table '" & TABLE_DOC_DOCUMENT_POSITION & "' is not available yet."
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    sqlText = "DELETE FROM [" & TABLE_DOC_DOCUMENT_POSITION & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    db.Execute sqlText, dbFailOnError

    modLoggingHandler.LogInfo MODULE_NAME & ".DeleteDocumentPositions", _
        "Deleted document positions for DocumentId=" & CStr(DocumentId) & "."

    DeleteDocumentPositions = True

CleanExit:
    On Error Resume Next
    Set db = Nothing
    Exit Function

ErrorHandler:
    DeleteDocumentPositions = False
    modErrorHandler.HandleError MODULE_NAME, "DeleteDocumentPositions", Err
    Resume CleanExit
End Function

Public Function CreateDocumentPosition( _
    ByVal DocumentId As Long, _
    ByVal LineNo As Long, _
    ByVal Description As String, _
    ByVal Quantity As Double, _
    ByVal UnitPrice As Currency, _
    Optional ByVal UnitCode As String = "", _
    Optional ByVal VatRate As Double = -1, _
    Optional ByVal VatMode As String = "" _
) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim effectiveVatRate As Double
    Dim effectiveVatMode As String
    Dim lineNet As Currency
    Dim lineVat As Currency
    Dim lineGross As Currency

    CreateDocumentPosition = False

    If DocumentId <= 0 Then
        Exit Function
    End If

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CreateDocumentPosition", _
            "Document position creation skipped because backend configuration is not valid."
        Exit Function
    End If

    If Not TableExists(TABLE_DOC_DOCUMENT_POSITION) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CreateDocumentPosition", _
            "Table '" & TABLE_DOC_DOCUMENT_POSITION & "' is not available yet."
        Exit Function
    End If

    If VatRate < 0 Then
        effectiveVatRate = modVatHandler.GetVatRate()
    Else
        effectiveVatRate = VatRate
    End If

    If LenB(Trim$(VatMode)) = 0 Then
        effectiveVatMode = modVatHandler.GetVatMode()
    Else
        effectiveVatMode = modVatHandler.NormalizeVatMode(VatMode)
    End If

    lineNet = modDocumentService.CalculateDocumentLineNet(Quantity, UnitPrice)
    lineVat = modDocumentService.CalculateDocumentLineVat(Quantity, UnitPrice, effectiveVatRate, effectiveVatMode)
    lineGross = modDocumentService.CalculateDocumentLineGross(Quantity, UnitPrice, effectiveVatRate, effectiveVatMode)

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset(TABLE_DOC_DOCUMENT_POSITION, dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    SetRecordsetValue rs, FIELD_DOCUMENT_ID, DocumentId
    SetRecordsetValue rs, FIELD_LINE_NO, LineNo
    SetRecordsetValue rs, FIELD_DESCRIPTION, Trim$(Description)
    SetRecordsetValue rs, FIELD_QUANTITY, Quantity
    SetRecordsetValue rs, FIELD_UNIT_CODE, Trim$(UnitCode)
    SetRecordsetValue rs, FIELD_UNIT_PRICE, UnitPrice
    SetRecordsetValue rs, FIELD_VAT_RATE, effectiveVatRate
    SetRecordsetValue rs, FIELD_LINE_TOTAL_NET, lineNet
    SetRecordsetValue rs, FIELD_LINE_TOTAL_VAT, lineVat
    SetRecordsetValue rs, FIELD_LINE_TOTAL_GROSS, lineGross
    rs.Update

    modLoggingHandler.LogInfo MODULE_NAME & ".CreateDocumentPosition", _
        "Document position created for DocumentId=" & CStr(DocumentId) & _
        ", LineNo=" & CStr(LineNo) & "."

    CreateDocumentPosition = True

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    CreateDocumentPosition = False
    modErrorHandler.HandleError MODULE_NAME, "CreateDocumentPosition", Err
    Resume CleanExit
End Function

Public Function UpdateDocumentTotals(ByVal DocumentId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsHeader As DAO.Recordset
    Dim rsPositions As DAO.Recordset
    Dim NetSum As Currency
    Dim VatSum As Currency
    Dim GrossSum As Currency
    Dim sqlText As String

    UpdateDocumentTotals = False

    If DocumentId <= 0 Then
        Exit Function
    End If

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".UpdateDocumentTotals", _
            "Document total update skipped because backend configuration is not valid."
        Exit Function
    End If

    If Not TableExists(TABLE_DOC_DOCUMENT) Or Not TableExists(TABLE_DOC_DOCUMENT_POSITION) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".UpdateDocumentTotals", _
            "Document total update skipped because required document tables are not available yet."
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()

    sqlText = "SELECT * FROM [" & TABLE_DOC_DOCUMENT & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsHeader = db.OpenRecordset(sqlText, dbOpenDynaset)

    If rsHeader.BOF And rsHeader.EOF Then
        modLoggingHandler.LogWarning MODULE_NAME & ".UpdateDocumentTotals", _
            "Document total update skipped because DocumentId=" & CStr(DocumentId) & " does not exist."
        GoTo CleanExit
    End If

    sqlText = "SELECT [" & FIELD_LINE_TOTAL_NET & "], [" & FIELD_LINE_TOTAL_VAT & "], [" & FIELD_LINE_TOTAL_GROSS & "] " & _
              "FROM [" & TABLE_DOC_DOCUMENT_POSITION & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsPositions = db.OpenRecordset(sqlText, dbOpenSnapshot)

    If Not (rsPositions.BOF And rsPositions.EOF) Then
        rsPositions.MoveFirst
        Do Until rsPositions.EOF
            If modDaoHelper.RecordsetHasField(rsPositions, FIELD_LINE_TOTAL_NET) Then
                NetSum = NetSum + CCur(modDaoHelper.NzString(rsPositions.Fields(FIELD_LINE_TOTAL_NET).Value, "0"))
            End If

            If modDaoHelper.RecordsetHasField(rsPositions, FIELD_LINE_TOTAL_VAT) Then
                VatSum = VatSum + CCur(modDaoHelper.NzString(rsPositions.Fields(FIELD_LINE_TOTAL_VAT).Value, "0"))
            End If

            If modDaoHelper.RecordsetHasField(rsPositions, FIELD_LINE_TOTAL_GROSS) Then
                GrossSum = GrossSum + CCur(modDaoHelper.NzString(rsPositions.Fields(FIELD_LINE_TOTAL_GROSS).Value, "0"))
            End If

            rsPositions.MoveNext
        Loop
    End If

    If Not modDocumentService.CalculateDocumentTotalsFromPositions(NetSum, VatSum, GrossSum) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".UpdateDocumentTotals", _
            "Document total update skipped because calculated totals were not accepted."
        GoTo CleanExit
    End If

    rsHeader.Edit
    SetRecordsetValue rsHeader, FIELD_TOTAL_NET, NetSum
    SetRecordsetValue rsHeader, FIELD_TOTAL_VAT, VatSum
    SetRecordsetValue rsHeader, FIELD_TOTAL_GROSS, GrossSum
    rsHeader.Update

    modLoggingHandler.LogInfo MODULE_NAME & ".UpdateDocumentTotals", _
        "Document totals updated for DocumentId=" & CStr(DocumentId) & _
        " (Net=" & CStr(NetSum) & ", Vat=" & CStr(VatSum) & ", Gross=" & CStr(GrossSum) & ")."

    UpdateDocumentTotals = True

CleanExit:
    On Error Resume Next
    If Not rsPositions Is Nothing Then rsPositions.Close
    If Not rsHeader Is Nothing Then rsHeader.Close
    Set rsPositions = Nothing
    Set rsHeader = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    UpdateDocumentTotals = False
    modLoggingHandler.LogError MODULE_NAME & ".UpdateDocumentTotals", _
        "Failed to update document totals for DocumentId=" & CStr(DocumentId) & ".", Err.Number
    modErrorHandler.HandleError MODULE_NAME, "UpdateDocumentTotals", Err
    Resume CleanExit
End Function

Public Function DocumentExists(ByVal DocumentId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsHeader As DAO.Recordset
    Dim sqlText As String

    If DocumentId <= 0 Then
        Exit Function
    End If

    If Not modDb.ValidateBackendConfiguration() Then
        Exit Function
    End If

    If Not TableExists(TABLE_DOC_DOCUMENT) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".DocumentExists", _
            "Table '" & TABLE_DOC_DOCUMENT & "' is not available yet."
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    sqlText = "SELECT * FROM [" & TABLE_DOC_DOCUMENT & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsHeader = db.OpenRecordset(sqlText, dbOpenSnapshot)

    DocumentExists = Not (rsHeader.BOF And rsHeader.EOF)

CleanExit:
    On Error Resume Next
    If Not rsHeader Is Nothing Then rsHeader.Close
    Set rsHeader = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    DocumentExists = False
    modErrorHandler.HandleError MODULE_NAME, "DocumentExists", Err
    Resume CleanExit
End Function

Public Function GetDocumentStatus(ByVal DocumentId As Long, Optional ByVal DefaultValue As String = "") As String
    On Error GoTo ErrorHandler

    GetDocumentStatus = ResolveDocumentFieldValue(DocumentId, FIELD_DOCUMENT_STATUS_CODE, DefaultValue)
    Exit Function

ErrorHandler:
    GetDocumentStatus = DefaultValue
    modErrorHandler.HandleError MODULE_NAME, "GetDocumentStatus", Err
End Function

Public Function GetDocumentNumber(ByVal DocumentId As Long, Optional ByVal DefaultValue As String = "") As String
    On Error GoTo ErrorHandler

    GetDocumentNumber = ResolveDocumentFieldValue(DocumentId, FIELD_DOCUMENT_NO, DefaultValue)
    Exit Function

ErrorHandler:
    GetDocumentNumber = DefaultValue
    modErrorHandler.HandleError MODULE_NAME, "GetDocumentNumber", Err
End Function

Public Function AssignDocumentNumber(ByVal DocumentId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsHeader As DAO.Recordset
    Dim sqlText As String
    Dim documentTypeCode As String
    Dim documentNo As String
    Dim documentDate As Date
    Dim nextDocumentNo As String

    AssignDocumentNumber = False

    If DocumentId <= 0 Then
        Exit Function
    End If

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".AssignDocumentNumber", _
            "Document number assignment skipped because backend configuration is not valid."
        Exit Function
    End If

    If Not TableExists(TABLE_DOC_DOCUMENT) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".AssignDocumentNumber", _
            "Table '" & TABLE_DOC_DOCUMENT & "' is not available yet."
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    sqlText = "SELECT * FROM [" & TABLE_DOC_DOCUMENT & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsHeader = db.OpenRecordset(sqlText, dbOpenDynaset)

    If rsHeader.BOF And rsHeader.EOF Then
        modLoggingHandler.LogWarning MODULE_NAME & ".AssignDocumentNumber", _
            "Document number assignment skipped because DocumentId=" & CStr(DocumentId) & " does not exist."
        GoTo CleanExit
    End If

    documentTypeCode = UCase$(Trim$(modDaoHelper.NzString(rsHeader.Fields(FIELD_DOCUMENT_TYPE_CODE).Value)))
    documentNo = Trim$(modDaoHelper.NzString(rsHeader.Fields(FIELD_DOCUMENT_NO).Value))

    If modDaoHelper.RecordsetHasField(rsHeader, FIELD_DOCUMENT_DATE) Then
        documentDate = rsHeader.Fields(FIELD_DOCUMENT_DATE).Value
    Else
        documentDate = Date
    End If

    If LenB(documentNo) > 0 Then
        modLoggingHandler.LogInfo MODULE_NAME & ".AssignDocumentNumber", _
            "DocumentId=" & CStr(DocumentId) & " already has document number '" & documentNo & "'."
        AssignDocumentNumber = True
        GoTo CleanExit
    End If

    nextDocumentNo = modNumberingHandler.GetNextDocumentNumber(documentTypeCode, documentDate)
    If LenB(Trim$(nextDocumentNo)) = 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".AssignDocumentNumber", _
            "Document number assignment failed because no next number could be generated for DocumentId=" & CStr(DocumentId) & "."
        GoTo CleanExit
    End If

    rsHeader.Edit
    SetRecordsetValue rsHeader, FIELD_DOCUMENT_NO, nextDocumentNo
    rsHeader.Update

    modLoggingHandler.LogInfo MODULE_NAME & ".AssignDocumentNumber", _
        "Assigned document number '" & nextDocumentNo & "' to DocumentId=" & CStr(DocumentId) & "."

    AssignDocumentNumber = True

CleanExit:
    On Error Resume Next
    If Not rsHeader Is Nothing Then rsHeader.Close
    Set rsHeader = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    AssignDocumentNumber = False
    modLoggingHandler.LogError MODULE_NAME & ".AssignDocumentNumber", _
        "Failed to assign document number for DocumentId=" & CStr(DocumentId) & ".", Err.Number
    modErrorHandler.HandleError MODULE_NAME, "AssignDocumentNumber", Err
    Resume CleanExit
End Function

Public Function SetDocumentStatus(ByVal DocumentId As Long, ByVal StatusCode As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsHeader As DAO.Recordset
    Dim sqlText As String

    SetDocumentStatus = False

    If DocumentId <= 0 Then
        Exit Function
    End If

    If Not modDb.ValidateBackendConfiguration() Then
        Exit Function
    End If

    If Not TableExists(TABLE_DOC_DOCUMENT) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".SetDocumentStatus", _
            "Table '" & TABLE_DOC_DOCUMENT & "' is not available yet."
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    sqlText = "SELECT * FROM [" & TABLE_DOC_DOCUMENT & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsHeader = db.OpenRecordset(sqlText, dbOpenDynaset)

    If rsHeader.BOF And rsHeader.EOF Then
        GoTo CleanExit
    End If

    rsHeader.Edit
    SetRecordsetValue rsHeader, FIELD_DOCUMENT_STATUS_CODE, UCase$(Trim$(StatusCode))
    rsHeader.Update

    SetDocumentStatus = True

CleanExit:
    On Error Resume Next
    If Not rsHeader Is Nothing Then rsHeader.Close
    Set rsHeader = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    SetDocumentStatus = False
    modErrorHandler.HandleError MODULE_NAME, "SetDocumentStatus", Err
    Resume CleanExit
End Function

Public Function CountDocumentPositions(ByVal DocumentId As Long) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsPositions As DAO.Recordset
    Dim sqlText As String

    If DocumentId <= 0 Then
        Exit Function
    End If

    If Not modDb.ValidateBackendConfiguration() Then
        Exit Function
    End If

    If Not TableExists(TABLE_DOC_DOCUMENT_POSITION) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CountDocumentPositions", _
            "Table '" & TABLE_DOC_DOCUMENT_POSITION & "' is not available yet."
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    sqlText = "SELECT * FROM [" & TABLE_DOC_DOCUMENT_POSITION & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsPositions = db.OpenRecordset(sqlText, dbOpenSnapshot)

    If Not (rsPositions.BOF And rsPositions.EOF) Then
        rsPositions.MoveLast
        CountDocumentPositions = rsPositions.RecordCount
    End If

CleanExit:
    On Error Resume Next
    If Not rsPositions Is Nothing Then rsPositions.Close
    Set rsPositions = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    CountDocumentPositions = 0
    modErrorHandler.HandleError MODULE_NAME, "CountDocumentPositions", Err
    Resume CleanExit
End Function

Private Function TableExists(ByVal TableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = modDb.GetCurrentDatabase()

    For Each tdf In db.TableDefs
        If UCase$(Trim$(tdf.Name)) = UCase$(Trim$(TableName)) Then
            TableExists = True
            Exit For
        End If
    Next tdf

CleanExit:
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    TableExists = False
    modErrorHandler.HandleError MODULE_NAME, "TableExists", Err
    Resume CleanExit
End Function

Private Sub SetRecordsetValue(ByVal rs As DAO.Recordset, ByVal FieldName As String, ByVal FieldValue As Variant)
    If modDaoHelper.RecordsetHasField(rs, FieldName) Then
        rs.Fields(FieldName).Value = FieldValue
    End If
End Sub

Private Function ResolveCurrencyCode() As String
    ResolveCurrencyCode = modTenantRepository.GetTenantParameter("CURRENCY_CODE", "CHF")
End Function

Private Function ResolveCreatedBy() As String
    If IsSessionInitialized() Then
        ResolveCreatedBy = currentUserId
    Else
        ResolveCreatedBy = "SYSTEM"
    End If
End Function

Private Function ResolveDocumentFieldValue(ByVal DocumentId As Long, ByVal FieldName As String, ByVal DefaultValue As String) As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rsHeader As DAO.Recordset
    Dim sqlText As String

    ResolveDocumentFieldValue = DefaultValue

    If DocumentId <= 0 Then
        Exit Function
    End If

    If Not modDb.ValidateBackendConfiguration() Then
        Exit Function
    End If

    If Not TableExists(TABLE_DOC_DOCUMENT) Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    sqlText = "SELECT * FROM [" & TABLE_DOC_DOCUMENT & "] WHERE [" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId) & ";"
    Set rsHeader = db.OpenRecordset(sqlText, dbOpenSnapshot)

    If rsHeader.BOF And rsHeader.EOF Then
        GoTo CleanExit
    End If

    If modDaoHelper.RecordsetHasField(rsHeader, FieldName) Then
        ResolveDocumentFieldValue = modDaoHelper.NzString(rsHeader.Fields(FieldName).Value, DefaultValue)
    End If

CleanExit:
    On Error Resume Next
    If Not rsHeader Is Nothing Then rsHeader.Close
    Set rsHeader = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    ResolveDocumentFieldValue = DefaultValue
    modErrorHandler.HandleError MODULE_NAME, "ResolveDocumentFieldValue", Err
    Resume CleanExit
End Function
