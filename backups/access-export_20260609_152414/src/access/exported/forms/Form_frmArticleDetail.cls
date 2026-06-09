Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmArticleDetail
' Purpose   : Workspace-aware detail form for simple Article master data.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmArticleDetail"

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    modFormRuntime.InitializeForm Me
    modFwTranslationRuntime.ApplyTranslations Me
    modArticleService.ConfigureArticleDetailCombos Me
    modArticleService.ApplyDefaultValues Me
    UpdateHeaderCaptionUi
    UpdateLocalCommandButtons

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "Article detail loaded."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_SAVE_ERROR", "Fehler beim Speichern des Artikels."), vbExclamation, MODULE_NAME
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    UpdateHeaderCaptionUi
    UpdateLocalCommandButtons
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo ErrorHandler

    modArticleService.ApplyDefaultValues Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_BeforeInsert", Err
    Cancel = True
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo ErrorHandler

    modArticleService.PrepareArticleForSave Me

    If Not modArticleService.ValidateArticleForm(Me) Then
        Cancel = True
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_BeforeUpdate", Err
    Cancel = True
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler

    DetailSave
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdSave_Click", Err
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo ErrorHandler

    DetailCancel
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdCancel_Click", Err
End Sub

Public Function CanLeaveWorkspace() As Boolean
    On Error GoTo ErrorHandler

    If Not HasUnsavedChanges() Then
        CanLeaveWorkspace = True
        Exit Function
    End If

    If MsgBox(modFwTranslationRuntime.ResolveText("MSG.ARTICLE_CANCEL_CONFIRM", "Aenderungen verwerfen?"), vbQuestion + vbYesNo + vbDefaultButton2, MODULE_NAME) = vbYes Then
        CancelCurrentEdit False, False
        CanLeaveWorkspace = True
    End If
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CanLeaveWorkspace", Err
End Function

Public Function SupportsWorkspaceCommandBar() As Boolean
    SupportsWorkspaceCommandBar = True
End Function

Public Function GetWorkspaceCommandBarConfig() As Object
    Dim cfg As Object

    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = vbTextCompare

    cfg("search.visible") = False
    cfg("search.enabled") = False

    cfg("L1.visible") = True
    cfg("L1.enabled") = True
    cfg("L1.caption") = modFwTranslationRuntime.ResolveText("COMMON.HOME", "Home")
    cfg("L1.command") = modAppShell.WCMD_NAV_HOME

    cfg("L2.visible") = True
    cfg("L2.enabled") = modAppWorkspaceService.CanGoBack()
    cfg("L2.caption") = modFwTranslationRuntime.ResolveText("COMMON.BACK", "Zurueck")
    cfg("L2.command") = modAppShell.WCMD_NAV_BACK

    cfg("L3.visible") = False
    cfg("L3.enabled") = False
    cfg("L3.caption") = vbNullString
    cfg("L3.command") = vbNullString

    cfg("R2.visible") = True
    cfg("R2.enabled") = True
    cfg("R2.caption") = modFwTranslationRuntime.ResolveText("COMMON.CANCEL", "Abbrechen")
    cfg("R2.command") = modAppShell.WCMD_DETAIL_CANCEL

    cfg("R1.visible") = True
    cfg("R1.enabled") = True
    cfg("R1.caption") = modFwTranslationRuntime.ResolveText("COMMON.SAVE", "Speichern")
    cfg("R1.command") = modAppShell.WCMD_DETAIL_SAVE

    Set GetWorkspaceCommandBarConfig = cfg
End Function

Public Function CanExecuteWorkspaceCommand(ByVal commandKey As String) As Boolean
    Select Case UCase$(Trim$(commandKey))
        Case modAppShell.WCMD_DETAIL_SAVE, modAppShell.WCMD_DETAIL_CANCEL
            CanExecuteWorkspaceCommand = True
    End Select
End Function

Public Sub ExecuteWorkspaceCommand(ByVal commandKey As String, Optional ByVal commandValue As String = "")
    Select Case UCase$(Trim$(commandKey))
        Case modAppShell.WCMD_DETAIL_SAVE
            DetailSave
        Case modAppShell.WCMD_DETAIL_CANCEL
            DetailCancel
    End Select
End Sub

Public Function GetCurrentWorkspaceRecordId() As Long
    GetCurrentWorkspaceRecordId = GetCurrentArticleId()
End Function

Public Function GetWorkspaceState() As String
    Dim stateText As String

    stateText = AddWorkspaceStatePair(stateText, "article_id", CStr(GetCurrentArticleId()))
    stateText = AddWorkspaceStatePair(stateText, "is_new", IIf(Me.NewRecord, "1", "0"))

    GetWorkspaceState = stateText
End Function

Public Sub RestoreWorkspaceState(ByVal stateText As String)
    On Error GoTo ErrorHandler

    Dim articleId As Long

    articleId = modDaoHelper.NzLong(GetWorkspaceStateValue(stateText, "article_id"), 0)
    If articleId > 0 Then
        RestoreCurrentArticleSelection articleId
    ElseIf HasUnsavedChanges() Then
        modArticleService.ApplyDefaultValues Me
    End If

    UpdateHeaderCaptionUi
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreWorkspaceState", Err
End Sub

Public Function HasUnsavedChanges() As Boolean
    On Error GoTo SafeExit

    HasUnsavedChanges = Me.Dirty

SafeExit:
End Function

Public Sub DetailSave()
    SaveCurrentRecord
End Sub

Public Sub DetailCancel()
    CancelCurrentEdit
End Sub

Private Sub SaveCurrentRecord()
    On Error GoTo ErrorHandler

    If Me.Dirty Then
        modArticleService.PrepareArticleForSave Me
        modAuditHelper.ApplyAuditFields Me
        Me.Dirty = False
    End If

    UpdateHeaderCaptionUi

    modLoggingHandler.LogInfo MODULE_NAME & ".SaveCurrentRecord", _
        "Article saved for article_id=" & CStr(GetCurrentArticleId()) & "."

    ReturnToPreviousWorkspace "SaveCurrentRecord"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SaveCurrentRecord", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_SAVE_ERROR", "Fehler beim Speichern des Artikels."), vbExclamation, MODULE_NAME
End Sub

Private Sub CancelCurrentEdit( _
    Optional ByVal showConfirmation As Boolean = True, _
    Optional ByVal navigateBack As Boolean = True)
    On Error GoTo ErrorHandler

    If Not HasUnsavedChanges() Then
        If navigateBack Then
            ReturnToPreviousWorkspace "CancelCurrentEdit"
        End If
        Exit Sub
    End If

    If showConfirmation Then
        If MsgBox(modFwTranslationRuntime.ResolveText("MSG.ARTICLE_CANCEL_CONFIRM", "Aenderungen verwerfen?"), vbQuestion + vbYesNo + vbDefaultButton2, MODULE_NAME) <> vbYes Then
            Exit Sub
        End If
    End If

    Me.Undo
    modArticleService.ApplyDefaultValues Me
    UpdateHeaderCaptionUi

    modLoggingHandler.LogInfo MODULE_NAME & ".CancelCurrentEdit", _
        "Unsaved Article changes were discarded."

    If navigateBack Then
        ReturnToPreviousWorkspace "CancelCurrentEdit"
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CancelCurrentEdit", Err
End Sub

Private Sub UpdateHeaderCaptionUi()
    On Error GoTo SafeExit

    Dim headerCaption As String
    Dim displayName As String

    headerCaption = modFwTranslationRuntime.ResolveText("FORM.FRMARTICLEDETAIL.FORM_TITLE", "Artikel")

    If GetCurrentArticleId() > 0 Then
        displayName = modArticleService.ResolveArticleDisplayName(Me)
        If LenB(displayName) > 0 Then
            headerCaption = modFwTranslationRuntime.ResolveText("FORM.FRMARTICLEDETAIL.FORM_TITLE", "Artikel") & ": " & displayName
        Else
            headerCaption = modFwTranslationRuntime.ResolveText("FORM.FRMARTICLEDETAIL.FORM_TITLE", "Artikel") & ": " & CStr(GetCurrentArticleId())
        End If
    End If

    If HasControl("lblFormTitle") Then
        Me!lblFormTitle.Caption = headerCaption
    ElseIf HasControl("txtFormTitle") Then
        Me!txtFormTitle.Value = headerCaption
    End If

SafeExit:
End Sub

Private Sub UpdateLocalCommandButtons()
    On Error GoTo SafeExit

    If HasControl("cmdSave") Then
        Me!cmdSave.Visible = False
    End If

    If HasControl("cmdCancel") Then
        Me!cmdCancel.Visible = False
    End If

SafeExit:
End Sub

Private Function GetCurrentArticleId() As Long
    On Error GoTo SafeExit

    If Me.NewRecord Then
        Exit Function
    End If

    GetCurrentArticleId = modArticleService.ResolveArticleId(Me)

SafeExit:
End Function

Private Sub RestoreCurrentArticleSelection(ByVal articleId As Long)
    On Error GoTo ErrorHandler

    Dim recordsetClone As DAO.Recordset

    Set recordsetClone = Me.RecordsetClone
    recordsetClone.FindFirst "[article_id]=" & CStr(articleId)

    If Not recordsetClone.NoMatch Then
        Me.Bookmark = recordsetClone.Bookmark
    End If

CleanExit:
    On Error Resume Next
    If Not recordsetClone Is Nothing Then recordsetClone.Close
    Set recordsetClone = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreCurrentArticleSelection", Err
    Resume CleanExit
End Sub

Private Function HasControl(ByVal controlName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentControl As Control

    For Each currentControl In Me.Controls
        If StrComp(currentControl.Name, controlName, vbTextCompare) = 0 Then
            HasControl = True
            Exit Function
        End If
    Next currentControl

SafeExit:
End Function

Private Function AddWorkspaceStatePair(ByVal stateText As String, ByVal keyName As String, ByVal valueText As String) As String
    If LenB(stateText) > 0 Then
        stateText = stateText & ";"
    End If

    AddWorkspaceStatePair = stateText & keyName & "=" & EscapeWorkspaceStateValue(valueText)
End Function

Private Function EscapeWorkspaceStateValue(ByVal valueText As String) As String
    valueText = Replace(modDaoHelper.NzString(valueText), "%", "%25")
    valueText = Replace(valueText, ";", "%3B")
    valueText = Replace(valueText, "=", "%3D")
    EscapeWorkspaceStateValue = valueText
End Function

Private Function GetWorkspaceStateValue(ByVal stateText As String, ByVal keyName As String) As String
    Dim parts() As String
    Dim pairText As Variant
    Dim separatorPosition As Long
    Dim currentKey As String

    If LenB(stateText) = 0 Then
        Exit Function
    End If

    parts = Split(stateText, ";")
    For Each pairText In parts
        separatorPosition = InStr(1, CStr(pairText), "=", vbBinaryCompare)
        If separatorPosition > 0 Then
            currentKey = Left$(CStr(pairText), separatorPosition - 1)
            If StrComp(currentKey, keyName, vbTextCompare) = 0 Then
                GetWorkspaceStateValue = UnescapeWorkspaceStateValue(Mid$(CStr(pairText), separatorPosition + 1))
                Exit Function
            End If
        End If
    Next pairText
End Function

Private Function UnescapeWorkspaceStateValue(ByVal valueText As String) As String
    valueText = Replace(valueText, "%3D", "=")
    valueText = Replace(valueText, "%3B", ";")
    valueText = Replace(valueText, "%25", "%")
    UnescapeWorkspaceStateValue = valueText
End Function

Private Sub ReturnToPreviousWorkspace(ByVal callerName As String)
    On Error GoTo ErrorHandler

    If Not modAppWorkspaceService.GoBack(Me) Then
        modLoggingHandler.LogWarning MODULE_NAME & "." & callerName, _
            "Workspace history was not available after detail action. Detail form remains open."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, callerName & ".ReturnToPreviousWorkspace", Err
End Sub
