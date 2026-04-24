Option Compare Database
Option Explicit

'===============================================================================
' Module    : modDocumentExportHandler
' Purpose   : UI-facing document export handler delegating PDF export to services.
' Author    : Codex
' Version   : 0.3.0
'===============================================================================

Private Const MODULE_NAME As String = "modDocumentExportHandler"

Public Function ExportDocumentToPdfLegacy(ByVal DocumentId As Long, Optional ByVal OutputPath As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim exportedFilePath As String
    Dim targetPdfPath As String

    ExportDocumentToPdfLegacy = False

    If DocumentId <= 0 Then
        Exit Function
    End If

    If LenB(Trim$(OutputPath)) > 0 Then
        targetPdfPath = modOutputPathService.BuildLegacyDocumentPdfPath(DocumentId, OutputPath)
        If LenB(targetPdfPath) = 0 Then
            modLoggingHandler.LogWarning MODULE_NAME & ".ExportDocumentToPdfLegacy", _
                "Export skipped because OutputPath could not be resolved for DocumentId=" & CStr(DocumentId) & "."
            Exit Function
        End If

        ExportDocumentToPdfLegacy = modPdfExportService.ExportDocumentToPdfAtPath(DocumentId, targetPdfPath, exportedFilePath)
        Exit Function
    End If

    ExportDocumentToPdfLegacy = modPdfExportService.ExportDocumentToPdf(DocumentId, exportedFilePath)
    Exit Function

ErrorHandler:
    ExportDocumentToPdfLegacy = False
    modErrorHandler.HandleError MODULE_NAME, "ExportDocumentToPdfLegacy", Err
End Function
