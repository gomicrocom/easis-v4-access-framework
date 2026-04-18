Attribute VB_Name = "modNumberingHandler"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modNumberingHandler
' Purpose   : Document numbering helpers based on tenant number ranges.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modNumberingHandler"
Private Const TABLE_TEN_NUMBERRANGE As String = "ten_numberrange"

Private Const FIELD_DOCUMENT_TYPE_CODE As String = "document_type_code"
Private Const FIELD_FISCAL_YEAR As String = "fiscal_year"
Private Const FIELD_PREFIX As String = "prefix"
Private Const FIELD_FORMAT_MASK As String = "format_mask"
Private Const FIELD_IS_ACTIVE As String = "is_active"

Public Function GetNextDocumentNumber(ByVal DocumentTypeCode As String, Optional ByVal DocumentDate As Date = 0) As String
    On Error GoTo ErrorHandler

    Dim fiscalYear As Long
    Dim nextValue As Long
    Dim prefix As String
    Dim formatMask As String
    Dim normalizedType As String

    normalizedType = UCase$(Trim$(DocumentTypeCode))
    If LenB(normalizedType) = 0 Then
        Exit Function
    End If

    fiscalYear = ResolveFiscalYear(DocumentDate)
    nextValue = modNumberRangeRepository.IncrementNumberValue(normalizedType, fiscalYear)

    If nextValue <= 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".GetNextDocumentNumber", _
            "Next document number could not be generated for DocumentType='" & normalizedType & "', FiscalYear=" & CStr(fiscalYear) & "."
        Exit Function
    End If

    prefix = ResolvePrefix(normalizedType, fiscalYear)
    formatMask = ResolveFormatMask(normalizedType, fiscalYear)

    GetNextDocumentNumber = BuildFormattedDocumentNumber(prefix, fiscalYear, nextValue, formatMask)
    Exit Function

ErrorHandler:
    GetNextDocumentNumber = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "GetNextDocumentNumber", Err
End Function

Public Function BuildFormattedDocumentNumber(ByVal Prefix As String, ByVal FiscalYear As Long, ByVal NumberValue As Long, Optional ByVal FormatMask As String = "") As String
    On Error GoTo ErrorHandler

    Dim effectiveMask As String
    Dim paddedValue As String

    paddedValue = Format$(NumberValue, "000000")
    effectiveMask = Trim$(FormatMask)

    If LenB(effectiveMask) = 0 Then
        BuildFormattedDocumentNumber = UCase$(Trim$(Prefix)) & "-" & CStr(FiscalYear) & "-" & paddedValue
        Exit Function
    End If

    effectiveMask = Replace(effectiveMask, "{PREFIX}", UCase$(Trim$(Prefix)))
    effectiveMask = Replace(effectiveMask, "{YEAR}", CStr(FiscalYear))
    effectiveMask = Replace(effectiveMask, "{NUMBER}", paddedValue)
    effectiveMask = Replace(effectiveMask, "PREFIX", UCase$(Trim$(Prefix)))
    effectiveMask = Replace(effectiveMask, "YEAR", CStr(FiscalYear))
    effectiveMask = Replace(effectiveMask, "000001", paddedValue)

    BuildFormattedDocumentNumber = effectiveMask
    Exit Function

ErrorHandler:
    BuildFormattedDocumentNumber = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "BuildFormattedDocumentNumber", Err
End Function

Private Function ResolveFiscalYear(ByVal DocumentDate As Date) As Long
    If DocumentDate = 0 Then
        ResolveFiscalYear = Year(Date)
    Else
        ResolveFiscalYear = Year(DocumentDate)
    End If
End Function

Private Function ResolvePrefix(ByVal DocumentTypeCode As String, ByVal FiscalYear As Long) As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ResolvePrefix = modTenantRepository.GetTenantParameter("NR_PREFIX_" & UCase$(Trim$(DocumentTypeCode)), UCase$(Trim$(DocumentTypeCode)))

    If Not modNumberRangeRepository.NumberRangeExists(DocumentTypeCode, FiscalYear) Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_TEN_NUMBERRANGE & "];", dbOpenSnapshot)

    ResolvePrefix = ResolveTextField(rs, DocumentTypeCode, FiscalYear, FIELD_PREFIX, ResolvePrefix)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ResolvePrefix", Err
    Resume CleanExit
End Function

Private Function ResolveFormatMask(ByVal DocumentTypeCode As String, ByVal FiscalYear As Long) As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ResolveFormatMask = modTenantRepository.GetTenantParameter("NR_FORMATMASK_" & UCase$(Trim$(DocumentTypeCode)), vbNullString)

    If Not modNumberRangeRepository.NumberRangeExists(DocumentTypeCode, FiscalYear) Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_TEN_NUMBERRANGE & "];", dbOpenSnapshot)

    ResolveFormatMask = ResolveTextField(rs, DocumentTypeCode, FiscalYear, FIELD_FORMAT_MASK, ResolveFormatMask)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ResolveFormatMask", Err
    Resume CleanExit
End Function

Private Function ResolveTextField(ByVal rs As DAO.Recordset, ByVal DocumentTypeCode As String, ByVal FiscalYear As Long, ByVal FieldName As String, ByVal DefaultValue As String) As String
    On Error GoTo ErrorHandler

    Dim targetType As String

    targetType = UCase$(Trim$(DocumentTypeCode))
    ResolveTextField = DefaultValue

    If rs.BOF And rs.EOF Then
        Exit Function
    End If

    If Not modDaoHelper.RecordsetHasField(rs, FIELD_DOCUMENT_TYPE_CODE) _
        Or Not modDaoHelper.RecordsetHasField(rs, FIELD_FISCAL_YEAR) _
        Or Not modDaoHelper.RecordsetHasField(rs, FieldName) Then
        Exit Function
    End If

    rs.MoveFirst
    Do Until rs.EOF
        If UCase$(Trim$(modDaoHelper.NzString(rs.Fields(FIELD_DOCUMENT_TYPE_CODE).Value))) = targetType _
            And modDaoHelper.NzLong(rs.Fields(FIELD_FISCAL_YEAR).Value, 0) = FiscalYear Then

            If modDaoHelper.RecordsetHasField(rs, FIELD_IS_ACTIVE) Then
                If Not modDaoHelper.NzBoolean(rs.Fields(FIELD_IS_ACTIVE).Value, False) Then
                    rs.MoveNext
                    GoTo ContinueLoop
                End If
            End If

            ResolveTextField = modDaoHelper.NzString(rs.Fields(FieldName).Value, DefaultValue)
            Exit Function
        End If

ContinueLoop:
        rs.MoveNext
    Loop
    Exit Function

ErrorHandler:
    ResolveTextField = DefaultValue
    modErrorHandler.HandleError MODULE_NAME, "ResolveTextField", Err
End Function
