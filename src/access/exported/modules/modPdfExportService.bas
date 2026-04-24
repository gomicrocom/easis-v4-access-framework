Attribute VB_Name = "modPdfExportService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modPdfExportService
' Purpose   : Exports framework documents to PDF using configured output paths.
' Author    : Codex
' Version   : 0.1.1
'===============================================================================

Private Const MODULE_NAME As String = "modPdfExportService"

Private Const REPORT_DOCUMENT As String = "rpt_document"
Private Const FIELD_DOCUMENT_ID As String = "document_id"

Public Function ExportDocumentToPdf(ByVal DocumentId As Long, Optional ByRef ExportedFilePath As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim targetPath As String

    ExportDocumentToPdf = False
    ExportedFilePath = vbNullString

    If DocumentId <= 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExportDocumentToPdf", _
            "PDF export skipped because DocumentId is invalid."
        Exit Function
    End If

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExportDocumentToPdf", _
            "PDF export skipped because backend configuration is not valid."
        Exit Function
    End If

    If Not modDocumentRepository.DocumentExists(DocumentId) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExportDocumentToPdf", _
            "PDF export skipped because DocumentId=" & CStr(DocumentId) & " does not exist."
        Exit Function
    End If

    If Not ReportExists(REPORT_DOCUMENT) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExportDocumentToPdf", _
            "PDF export skipped because report '" & REPORT_DOCUMENT & "' is not available."
        Exit Function
    End If

    targetPath = modOutputPathService.BuildDocumentPdfPath(DocumentId)
    If LenB(targetPath) = 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExportDocumentToPdf", _
            "PDF export skipped because target path could not be resolved for DocumentId=" & CStr(DocumentId) & "."
        Exit Function
    End If

    ExportDocumentToPdf = ExportDocumentToPdfAtPath(DocumentId, targetPath, ExportedFilePath)
    Exit Function

ErrorHandler:
    ExportedFilePath = vbNullString
    ExportDocumentToPdf = False
    modErrorHandler.HandleError MODULE_NAME, "ExportDocumentToPdf", Err
End Function

Public Function ExportDocumentToPdfAtPath(ByVal DocumentId As Long, ByVal TargetPdfPath As String, Optional ByRef ExportedFilePath As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim targetPath As String
    Dim targetFolder As String
    Dim whereCondition As String

    ExportDocumentToPdfAtPath = False
    ExportedFilePath = vbNullString

    targetPath = Trim$(TargetPdfPath)

    If DocumentId <= 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExportDocumentToPdfAtPath", _
            "PDF export skipped because DocumentId is invalid."
        Exit Function
    End If

    If LenB(targetPath) = 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExportDocumentToPdfAtPath", _
            "PDF export skipped because target PDF path is empty."
        Exit Function
    End If

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExportDocumentToPdfAtPath", _
            "PDF export skipped because backend configuration is not valid."
        Exit Function
    End If

    If Not modDocumentRepository.DocumentExists(DocumentId) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExportDocumentToPdfAtPath", _
            "PDF export skipped because DocumentId=" & CStr(DocumentId) & " does not exist."
        Exit Function
    End If

    If Not ReportExists(REPORT_DOCUMENT) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExportDocumentToPdfAtPath", _
            "PDF export skipped because report '" & REPORT_DOCUMENT & "' is not available."
        Exit Function
    End If

    targetFolder = GetFolderFromFilePath(targetPath)
    If Not modOutputPathService.EnsureDirectoryExists(targetFolder) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExportDocumentToPdfAtPath", _
            "PDF export skipped because target folder could not be prepared: '" & targetFolder & "'."
        Exit Function
    End If

    whereCondition = "[" & FIELD_DOCUMENT_ID & "]=" & CStr(DocumentId)

    DoCmd.OpenReport REPORT_DOCUMENT, acViewPreview, , whereCondition, acHidden
    DoCmd.OutputTo acOutputReport, REPORT_DOCUMENT, acFormatPDF, TargetPdfPath, False
    DoCmd.Close acReport, REPORT_DOCUMENT, acSaveNo

    ExportedFilePath = targetPath
    ExportDocumentToPdfAtPath = True

    modLoggingHandler.LogInfo MODULE_NAME & ".ExportDocumentToPdfAtPath", _
        "DocumentId=" & CStr(DocumentId) & " exported to PDF '" & targetPath & "'."
    Exit Function

ErrorHandler:
    ExportDocumentToPdfAtPath = False

    On Error Resume Next
    DoCmd.Close acReport, REPORT_DOCUMENT, acSaveNo
    On Error GoTo 0

    modLoggingHandler.LogError _
        MODULE_NAME & ".ExportDocumentToPdfAtPath", _
        "PDF export failed. DocumentId=" & CStr(DocumentId) & _
        "; TargetPdfPath=" & TargetPdfPath & _
        "; Err.Number=" & CStr(Err.Number) & _
        "; Err.Description=" & Err.Description, _
        Err.Number

    modErrorHandler.HandleError MODULE_NAME, "ExportDocumentToPdfAtPath", Err
End Function

Public Function ReportExists(ByVal ReportName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim reportObject As AccessObject
    Dim normalizedReportName As String

    normalizedReportName = UCase$(Trim$(ReportName))
    If LenB(normalizedReportName) = 0 Then
        Exit Function
    End If

    For Each reportObject In CurrentProject.AllReports
        If UCase$(Trim$(reportObject.Name)) = normalizedReportName Then
            ReportExists = True
            Exit Function
        End If
    Next reportObject
    Exit Function

ErrorHandler:
    ReportExists = False
    modErrorHandler.HandleError MODULE_NAME, "ReportExists", Err
End Function

Private Function GetFolderFromFilePath(ByVal FilePath As String) As String
    On Error GoTo ErrorHandler

    Dim separatorPosition As Long

    separatorPosition = InStrRev(Trim$(FilePath), "\")
    If separatorPosition > 0 Then
        GetFolderFromFilePath = Left$(Trim$(FilePath), separatorPosition - 1)
    End If
    Exit Function

ErrorHandler:
    GetFolderFromFilePath = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "GetFolderFromFilePath", Err
End Function
