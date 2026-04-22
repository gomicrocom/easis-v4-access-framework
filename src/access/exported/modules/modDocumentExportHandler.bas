Option Compare Database
Option Explicit

'===============================================================================
' Module    : modDocumentExportHandler
' Purpose   : Exports documents via Access reports to PDF files.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modDocumentExportHandler"
Private Const REPORT_DOCUMENT As String = "rpt_document"
Private Const DEFAULT_OUTPUT_PATH As String = "C:\Easis\Output\"
Private Const FIELD_DOCUMENT_ID As String = "document_id"

Public Function ExportDocumentToPdf(ByVal DocumentId As Long, Optional ByVal OutputPath As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim documentNo As String
    Dim filePath As String
    Dim whereCondition As String
    Dim targetFolder As String

    ExportDocumentToPdf = False

    If DocumentId <= 0 Then
        Exit Function
    End If

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExportDocumentToPdf", _
            "Export skipped because backend configuration is not valid."
        Exit Function
    End If

    If Not modDocumentRepository.DocumentExists(DocumentId) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExportDocumentToPdf", _
            "Export skipped because DocumentId=" & CStr(DocumentId) & " does not exist."
        Exit Function
    End If

    documentNo = Trim$(modDocumentRepository.GetDocumentNumber(DocumentId, vbNullString))
    If LenB(documentNo) = 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExportDocumentToPdf", _
            "Export skipped because DocumentId=" & CStr(DocumentId) & " has no document number."
        Exit Function
    End If

    If LenB(Trim$(OutputPath)) = 0 Then
        targetFolder = DEFAULT_OUTPUT_PATH
    Else
        targetFolder = Trim$(OutputPath)
    End If

    targetFolder = EnsureTrailingBackslash(targetFolder)
    EnsureFolderExists targetFolder
    filePath = targetFolder & SanitizeFileName(documentNo) & ".pdf"

    whereCondition = "[" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId)

    DoCmd.OpenReport REPORT_DOCUMENT, acViewPreview, , whereCondition, acHidden
    DoCmd.OutputTo acOutputReport, REPORT_DOCUMENT, acFormatPDF, filePath, False
    DoCmd.Close acReport, REPORT_DOCUMENT, acSaveNo

    modLoggingHandler.LogInfo MODULE_NAME & ".ExportDocumentToPdf", _
        "DocumentId=" & CStr(DocumentId) & " exported to '" & filePath & "'."

    ExportDocumentToPdf = True
    Exit Function

ErrorHandler:
    On Error Resume Next
    DoCmd.Close acReport, REPORT_DOCUMENT, acSaveNo
    ExportDocumentToPdf = False
    modErrorHandler.HandleError MODULE_NAME, "ExportDocumentToPdf", Err
End Function

Private Function EnsureTrailingBackslash(ByVal FolderPath As String) As String
    Dim normalizedPath As String

    normalizedPath = Trim$(FolderPath)

    If LenB(normalizedPath) = 0 Then
        EnsureTrailingBackslash = "\"
    ElseIf Right$(normalizedPath, 1) = "\" Then
        EnsureTrailingBackslash = normalizedPath
    Else
        EnsureTrailingBackslash = normalizedPath & "\"
    End If
End Function

Private Sub EnsureFolderExists(ByVal FolderPath As String)
    On Error GoTo ErrorHandler

    Dim normalizedPath As String
    Dim pathParts() As String
    Dim currentPath As String
    Dim index As Long

    normalizedPath = Trim$(FolderPath)
    If LenB(normalizedPath) = 0 Then
        Exit Sub
    End If

    If Right$(normalizedPath, 1) = "\" Then
        normalizedPath = Left$(normalizedPath, Len(normalizedPath) - 1)
    End If

    If LenB(Dir$(normalizedPath, vbDirectory)) > 0 Then
        Exit Sub
    End If

    pathParts = Split(normalizedPath, "\")
    If UBound(pathParts) < 0 Then
        Exit Sub
    End If

    currentPath = pathParts(0)
    If Right$(currentPath, 1) <> "\" Then
        currentPath = currentPath & "\"
    End If

    For index = 1 To UBound(pathParts)
        currentPath = currentPath & pathParts(index)
        If LenB(Dir$(currentPath, vbDirectory)) = 0 Then
            MkDir currentPath
        End If
        currentPath = currentPath & "\"
    Next index
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureFolderExists", Err
End Sub

Private Function SanitizeFileName(ByVal FileNameText As String) As String
    Dim resultText As String

    resultText = Trim$(FileNameText)

    resultText = Replace(resultText, "\", "-")
    resultText = Replace(resultText, "/", "-")
    resultText = Replace(resultText, ":", "-")
    resultText = Replace(resultText, "*", "-")
    resultText = Replace(resultText, "?", "")
    resultText = Replace(resultText, """", "")
    resultText = Replace(resultText, "<", "(")
    resultText = Replace(resultText, ">", ")")
    resultText = Replace(resultText, "|", "-")

    SanitizeFileName = resultText
End Function