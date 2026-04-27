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

Public Function GetNextDocumentNumber(ByVal documentTypeCode As String, Optional ByVal documentDate As Date = 0) As String
    On Error GoTo ErrorHandler

    Dim FiscalYear As Long
    Dim nextValue As Long
    Dim Prefix As String
    Dim FormatMask As String
    Dim normalizedType As String

    normalizedType = UCase$(Trim$(documentTypeCode))
    If LenB(normalizedType) = 0 Then
        Exit Function
    End If

    FiscalYear = ResolveFiscalYear(documentDate)
    nextValue = modNumberRangeRepository.IncrementNumberValue(normalizedType, FiscalYear)

    If nextValue <= 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".GetNextDocumentNumber", _
            "Next document number could not be generated for DocumentType='" & normalizedType & "', FiscalYear=" & CStr(FiscalYear) & "."
        Exit Function
    End If

    Prefix = ResolvePrefix(normalizedType, FiscalYear)
    FormatMask = ResolveFormatMask(normalizedType, FiscalYear)

    GetNextDocumentNumber = BuildFormattedDocumentNumber(Prefix, FiscalYear, nextValue, FormatMask)
    Exit Function

ErrorHandler:
    GetNextDocumentNumber = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "GetNextDocumentNumber", Err
End Function

Public Function BuildFormattedDocumentNumber( _
    ByVal Prefix As String, _
    ByVal FiscalYear As Long, _
    ByVal NumberValue As Long, _
    Optional ByVal FormatMask As String = "" _
) As String
    On Error GoTo ErrorHandler

    Dim effectiveMask As String
    Dim resultText As String
    Dim numberStartPos As Long
    Dim numberEndPos As Long
    Dim numberToken As String
    Dim numberPattern As String
    Dim formattedNumber As String

    effectiveMask = Trim$(FormatMask)
    If LenB(effectiveMask) = 0 Then
        effectiveMask = "{PREFIX}-{YEAR}-{NUMBER:000001}"
    End If

    resultText = effectiveMask
    resultText = Replace(resultText, "{PREFIX}", Trim$(Prefix))
    resultText = Replace(resultText, "{YEAR}", CStr(FiscalYear))

    numberStartPos = InStr(1, resultText, "{NUMBER:", vbTextCompare)

    If numberStartPos > 0 Then
        numberEndPos = InStr(numberStartPos, resultText, "}", vbTextCompare)

        If numberEndPos > numberStartPos Then
            numberToken = Mid$(resultText, numberStartPos, numberEndPos - numberStartPos + 1)
            numberPattern = Mid$(resultText, numberStartPos + 8, numberEndPos - numberStartPos - 8)

            If LenB(numberPattern) = 0 Then
                numberPattern = "000001"
            End If

            formattedNumber = Format$(NumberValue, numberPattern)
            resultText = Replace(resultText, numberToken, formattedNumber)
        End If
    Else
        resultText = resultText & "-" & Format$(NumberValue, "000001")
    End If

    BuildFormattedDocumentNumber = resultText
    Exit Function

ErrorHandler:
    BuildFormattedDocumentNumber = Trim$(Prefix) & "-" & CStr(FiscalYear) & "-" & Format$(NumberValue, "000001")
    modErrorHandler.HandleError "modNumberingHandler", "BuildFormattedDocumentNumber", Err
End Function

Private Function ResolveFiscalYear(ByVal documentDate As Date) As Long
    If documentDate = 0 Then
        ResolveFiscalYear = Year(Date)
    Else
        ResolveFiscalYear = Year(documentDate)
    End If
End Function

Private Function ResolvePrefix(ByVal documentTypeCode As String, ByVal FiscalYear As Long) As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ResolvePrefix = modTenantRepository.GetTenantParameter("NR_PREFIX_" & UCase$(Trim$(documentTypeCode)), UCase$(Trim$(documentTypeCode)))

    If Not modNumberRangeRepository.NumberRangeExists(documentTypeCode, FiscalYear) Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_TEN_NUMBERRANGE & "];", dbOpenSnapshot)

    ResolvePrefix = ResolveTextField(rs, documentTypeCode, FiscalYear, FIELD_PREFIX, ResolvePrefix)

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

Private Function ResolveFormatMask(ByVal documentTypeCode As String, ByVal FiscalYear As Long) As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    ResolveFormatMask = modTenantRepository.GetTenantParameter("NR_FORMATMASK_" & UCase$(Trim$(documentTypeCode)), vbNullString)

    If Not modNumberRangeRepository.NumberRangeExists(documentTypeCode, FiscalYear) Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_TEN_NUMBERRANGE & "];", dbOpenSnapshot)

    ResolveFormatMask = ResolveTextField(rs, documentTypeCode, FiscalYear, FIELD_FORMAT_MASK, ResolveFormatMask)

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

Private Function ResolveTextField(ByVal rs As DAO.Recordset, ByVal documentTypeCode As String, ByVal FiscalYear As Long, ByVal fieldName As String, ByVal DefaultValue As String) As String
    On Error GoTo ErrorHandler

    Dim targetType As String

    targetType = UCase$(Trim$(documentTypeCode))
    ResolveTextField = DefaultValue

    If rs.BOF And rs.EOF Then
        Exit Function
    End If

    If Not modDaoHelper.RecordsetHasField(rs, FIELD_DOCUMENT_TYPE_CODE) _
        Or Not modDaoHelper.RecordsetHasField(rs, FIELD_FISCAL_YEAR) _
        Or Not modDaoHelper.RecordsetHasField(rs, fieldName) Then
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

            ResolveTextField = modDaoHelper.NzString(rs.Fields(fieldName).Value, DefaultValue)
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
