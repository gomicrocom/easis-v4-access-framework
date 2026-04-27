Attribute VB_Name = "modOutputPathService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modOutputPathService
' Purpose   : Resolves configured document output paths and PDF target locations.
' Author    : Codex
' Version   : 0.1.3
'===============================================================================

Private Const MODULE_NAME As String = "modOutputPathService"

Private Const TABLE_TEN_PARAMETER As String = "ten_parameter"
Private Const TENANT_PARAMETER_BASIC_DOC_PATH As String = "BASIC_DOC_PATH"

Private Const DEFAULT_PATH_SEGMENT As String = "Unknown"
Private Const RESERVED_FILE_PREFIX As String = "File-"
Private Const PDF_EXTENSION As String = ".pdf"
Private Const DOCUMENT_FILE_PREFIX As String = "Document-"

Public Function GetDocumentRootPath(Optional ByVal IniPath As String = "") As String
    On Error GoTo ErrorHandler

    GetDocumentRootPath = GetTenantDocumentRootPath()
    Exit Function

ErrorHandler:
    GetDocumentRootPath = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "GetDocumentRootPath", Err
End Function

Public Function GetTenantDocumentRootPath() As String
    On Error GoTo ErrorHandler

    Dim rootPath As String

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".GetTenantDocumentRootPath", _
            "Document root path could not be resolved because backend configuration is not valid."
        Exit Function
    End If

    If Not TableExists(TABLE_TEN_PARAMETER) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".GetTenantDocumentRootPath", _
            "Table '" & TABLE_TEN_PARAMETER & "' is not available."
        Exit Function
    End If

    If Not modTenantRepository.HasTenantParameter(TENANT_PARAMETER_BASIC_DOC_PATH) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".GetTenantDocumentRootPath", _
            "Tenant parameter '" & TENANT_PARAMETER_BASIC_DOC_PATH & "' was not found."
        Exit Function
    End If

    rootPath = Trim$(modTenantRepository.GetTenantParameter(TENANT_PARAMETER_BASIC_DOC_PATH, vbNullString))

    If LenB(rootPath) = 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".GetTenantDocumentRootPath", _
            "Tenant parameter '" & TENANT_PARAMETER_BASIC_DOC_PATH & "' is empty."
        Exit Function
    End If

    GetTenantDocumentRootPath = rootPath
    Exit Function

ErrorHandler:
    GetTenantDocumentRootPath = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "GetTenantDocumentRootPath", Err
End Function

Public Function BuildDocumentPdfPath(ByVal DocumentId As Long) As String
    On Error GoTo ErrorHandler

    Dim rootPath As String
    Dim customerName As String
    Dim documentNo As String
    Dim customerSegment As String
    Dim documentSegment As String

    If DocumentId <= 0 Then
        Exit Function
    End If

    If Not modDocumentRepository.DocumentExists(DocumentId) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".BuildDocumentPdfPath", _
            "PDF path could not be built because DocumentId=" & CStr(DocumentId) & " does not exist."
        Exit Function
    End If

    rootPath = GetDocumentRootPath()
    If LenB(rootPath) = 0 Then
        Exit Function
    End If

    customerName = modDocumentRepository.GetDocumentCustomerName(DocumentId, DEFAULT_PATH_SEGMENT)
    documentNo = modDocumentRepository.GetDocumentNumber(DocumentId, DOCUMENT_FILE_PREFIX & CStr(DocumentId))

    If LenB(Trim$(customerName)) = 0 Then
        customerName = DEFAULT_PATH_SEGMENT
    End If

    If LenB(Trim$(documentNo)) = 0 Then
        documentNo = DOCUMENT_FILE_PREFIX & CStr(DocumentId)
    End If

    customerSegment = SanitizePathSegment(customerName)
    documentSegment = SanitizePathSegment(documentNo)

    BuildDocumentPdfPath = EnsureTrailingBackslash(rootPath) & _
                           customerSegment & "\" & _
                           documentSegment & PDF_EXTENSION
    Exit Function

ErrorHandler:
    BuildDocumentPdfPath = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "BuildDocumentPdfPath", Err
End Function

Public Function EnsureDirectoryExists(ByVal FolderPath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim normalizedPath As String
    Dim pathParts() As String
    Dim currentPath As String
    Dim index As Long
    Dim startIndex As Long

    normalizedPath = Trim$(FolderPath)
    If LenB(normalizedPath) = 0 Then
        Exit Function
    End If

    If Right$(normalizedPath, 1) = "\" Then
        normalizedPath = Left$(normalizedPath, Len(normalizedPath) - 1)
    End If

    If LenB(Dir$(normalizedPath, vbDirectory)) > 0 Then
        EnsureDirectoryExists = True
        Exit Function
    End If

    If Left$(normalizedPath, 2) = "\\" Then
        pathParts = Split(Mid$(normalizedPath, 3), "\")
        If UBound(pathParts) < 1 Then
            Exit Function
        End If

        currentPath = "\\" & pathParts(0) & "\" & pathParts(1)
        startIndex = 2
    Else
        pathParts = Split(normalizedPath, "\")
        If UBound(pathParts) < 0 Then
            Exit Function
        End If

        currentPath = pathParts(0)
        startIndex = 1
    End If

    For index = startIndex To UBound(pathParts)
        If LenB(pathParts(index)) > 0 Then
            currentPath = EnsureTrailingBackslash(currentPath) & pathParts(index)
            If LenB(Dir$(currentPath, vbDirectory)) = 0 Then
                MkDir currentPath
            End If
        End If
    Next index

    EnsureDirectoryExists = (LenB(Dir$(normalizedPath, vbDirectory)) > 0)
    Exit Function

ErrorHandler:
    EnsureDirectoryExists = False
    modErrorHandler.HandleError MODULE_NAME, "EnsureDirectoryExists", Err
End Function

Public Function SanitizePathSegment(ByVal Value As String) As String
    Dim resultText As String

    resultText = Trim$(Value)

    resultText = Replace(resultText, "\", "-")
    resultText = Replace(resultText, "/", "-")
    resultText = Replace(resultText, ":", "-")
    resultText = Replace(resultText, "*", "-")
    resultText = Replace(resultText, "?", "")
    resultText = Replace(resultText, """", "")
    resultText = Replace(resultText, "<", "(")
    resultText = Replace(resultText, ">", ")")
    resultText = Replace(resultText, "|", "-")

    resultText = TrimTrailingDotsAndSpaces(resultText)
    If LenB(resultText) = 0 Then
        resultText = DEFAULT_PATH_SEGMENT
    ElseIf IsReservedWindowsDeviceName(resultText) Then
        resultText = RESERVED_FILE_PREFIX & resultText
    End If

    SanitizePathSegment = resultText
End Function

Public Function EnsureTrailingBackslash(ByVal FolderPath As String) As String
    Dim normalizedPath As String

    normalizedPath = Trim$(FolderPath)
    If LenB(normalizedPath) = 0 Then
        Exit Function
    End If

    If Right$(normalizedPath, 1) = "\" Then
        EnsureTrailingBackslash = normalizedPath
    Else
        EnsureTrailingBackslash = normalizedPath & "\"
    End If
End Function

Private Function TrimTrailingDotsAndSpaces(ByVal Value As String) As String
    Dim resultText As String

    resultText = Trim$(Value)

    Do While LenB(resultText) > 0 And (Right$(resultText, 1) = "." Or Right$(resultText, 1) = " ")
        resultText = Left$(resultText, Len(resultText) - 1)
    Loop

    TrimTrailingDotsAndSpaces = Trim$(resultText)
End Function

Private Function IsReservedWindowsDeviceName(ByVal Value As String) As Boolean
    Dim normalizedValue As String
    Dim prefixText As String
    Dim suffixText As String

    normalizedValue = UCase$(Trim$(Value))

    Select Case normalizedValue
        Case "CON", "PRN", "AUX", "NUL"
            IsReservedWindowsDeviceName = True
            Exit Function
    End Select

    If Len(normalizedValue) = 4 Then
        prefixText = Left$(normalizedValue, 3)
        suffixText = Right$(normalizedValue, 1)

        If (prefixText = "COM" Or prefixText = "LPT") And suffixText >= "1" And suffixText <= "9" Then
            IsReservedWindowsDeviceName = True
        End If
    End If
End Function

Private Function TableExists(ByVal TableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tdf As DAO.tableDef

    If LenB(Trim$(TableName)) = 0 Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Exit Function
    End If

    For Each tdf In db.TableDefs
        If UCase$(Trim$(tdf.Name)) = UCase$(Trim$(TableName)) Then
            TableExists = True
            Exit Function
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
