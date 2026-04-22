Option Compare Database
Option Explicit

'===============================================================================
' Module    : modDocumentPostingHandler
' Purpose   : Handles document posting workflow for validated business documents.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modDocumentPostingHandler"
Private Const STATUS_DRAFT As String = "DRAFT"
Private Const STATUS_POSTED As String = "POSTED"

Public Function PostDocument(ByVal DocumentId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim currentStatus As String
    Dim documentNo As String

    PostDocument = False

    If DocumentId <= 0 Then
        Exit Function
    End If

    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".PostDocument", _
            "Document posting skipped because backend configuration is not valid."
        Exit Function
    End If

    If Not modDocumentRepository.DocumentExists(DocumentId) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".PostDocument", _
            "Document posting failed because DocumentId=" & CStr(DocumentId) & " does not exist."
        Exit Function
    End If

    If modDocumentRepository.CountDocumentPositions(DocumentId) <= 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".PostDocument", _
            "Document posting failed because DocumentId=" & CStr(DocumentId) & " has no positions."
        Exit Function
    End If

    currentStatus = UCase$(Trim$(modDocumentRepository.GetDocumentStatus(DocumentId, STATUS_DRAFT)))
    If currentStatus = STATUS_POSTED Then
        modLoggingHandler.LogInfo MODULE_NAME & ".PostDocument", _
            "DocumentId=" & CStr(DocumentId) & " is already posted."
        PostDocument = True
        Exit Function
    End If

    If Not modDocumentRepository.UpdateDocumentTotals(DocumentId) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".PostDocument", _
            "Document posting failed because totals could not be updated for DocumentId=" & CStr(DocumentId) & "."
        Exit Function
    End If

    documentNo = Trim$(modDocumentRepository.GetDocumentNumber(DocumentId, vbNullString))
    If LenB(documentNo) = 0 Then
        If Not modDocumentRepository.AssignDocumentNumber(DocumentId) Then
            modLoggingHandler.LogWarning MODULE_NAME & ".PostDocument", _
                "Document posting failed because document number could not be assigned for DocumentId=" & CStr(DocumentId) & "."
            Exit Function
        End If
    End If

    If Not modDocumentRepository.SetDocumentStatus(DocumentId, STATUS_POSTED) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".PostDocument", _
            "Document posting failed because status could not be set to POSTED for DocumentId=" & CStr(DocumentId) & "."
        Exit Function
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".PostDocument", _
        "DocumentId=" & CStr(DocumentId) & " posted successfully."
    PostDocument = True
    Exit Function

ErrorHandler:
    PostDocument = False
    modLoggingHandler.LogError MODULE_NAME & ".PostDocument", _
        "Failed to post document DocumentId=" & CStr(DocumentId) & ".", Err.Number
    modErrorHandler.HandleError MODULE_NAME, "PostDocument", Err
End Function