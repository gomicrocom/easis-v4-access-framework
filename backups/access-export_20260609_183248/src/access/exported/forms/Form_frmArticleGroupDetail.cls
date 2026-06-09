VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmArticleGroupDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ExportKind: CODE_BEHIND_ONLY
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmArticleGroupDetail
' Purpose   : Workspace-aware detail form for tenant Article Groups.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmArticleGroupDetail"

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    modArticleGroupService.EnsureArticleGroupTable
    modFormRuntime.InitializeForm Me
    modFwTranslationRuntime.ApplyTranslations Me
    modArticleGroupService.ApplyDefaultValues Me
    UpdateHeaderCaptionUi
    UpdateLocalCommandButtons

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "Article Group detail loaded."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_GROUP_DETAIL_LOAD_ERROR", "Fehler beim Laden der Artikelgruppendetails: " & Err.description), vbExclamation, MODULE_NAME
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

    modArticleGroupService.ApplyDefaultValues Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_BeforeInsert", Err
    Cancel = True
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo ErrorHandler

    modArticleGroupService.PrepareArticleGroupForSave Me

    If Not modArticleGroupService.ValidateArticleGroupForm(Me) Then
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

Private Sub cmdRefresh_Click()
    On Error GoTo ErrorHandler

    RefreshCurrentRecord
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdRefresh_Click", Err
End Sub

Public Function CanLeaveWorkspace() As Boolean
    On Error GoTo ErrorHandler

    If Not HasUnsavedChanges() Then
        CanLeaveWorkspace = True
        Exit Function
    End If

    If MsgBox(modFwTranslationRuntime.ResolveText("MSG.ARTICLE_GROUP_DISCARD_CHANGES", "Ungespeicherte Aenderungen verwerfen?"), vbQuestion + vbYesNo + vbDefaultButton2, MODULE_NAME) = vbYes Then
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
    GetCurrentWorkspaceRecordId = GetCurrentArticleGroupId()
End Function

Public Function GetWorkspaceState() As String
    Dim stateText As String

    stateText = AddWorkspaceStatePair(stateText, "product_group_id", CStr(GetCurrentArticleGroupId()))
    stateText = AddWorkspaceStatePair(stateText, "is_new", IIf(Me.NewRecord, "1", "0"))

    GetWorkspaceState = stateText
End Function

Public Sub RestoreWorkspaceState(ByVal stateText As String)
    On Error GoTo ErrorHandler

    Dim productGroupId As Long

    productGroupId = modDaoHelper.NzLong(GetWorkspaceStateValue(stateText, "product_group_id"), 0)
    If productGroupId > 0 Then
        RestoreCurrentArticleGroupSelection productGroupId
    ElseIf HasUnsavedChanges() Then
        modArticleGroupService.ApplyDefaultValues Me
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
        modAuditHelper.ApplyAuditFields Me
        Me.Dirty = False
    End If

    UpdateHeaderCaptionUi

    modLoggingHandler.LogInfo MODULE_NAME & ".SaveCurrentRecord", _
        "Article Group saved for product_group_id=" & CStr(GetCurrentArticleGroupId()) & "."

    ReturnToPreviousWorkspace "SaveCurrentRecord"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SaveCurrentRecord", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_GROUP_SAVE_ERROR", "Fehler beim Speichern der Artikelgruppe: " & Err.description), vbExclamation, MODULE_NAME
End Sub

Private Sub CancelCurrentEdit( _
    Optional ByVal showConfirmation As Boolean = True, _
    Optional ByVal NavigateBack As Boolean = True)
    On Error GoTo ErrorHandler

    If Not HasUnsavedChanges() Then
        If NavigateBack Then
            ReturnToPreviousWorkspace "CancelCurrentEdit"
        End If
        Exit Sub
    End If

    If showConfirmation Then
        If MsgBox(modFwTranslationRuntime.ResolveText("MSG.ARTICLE_GROUP_CANCEL_CONFIRM", "Aenderungen verwerfen?"), vbQuestion + vbYesNo + vbDefaultButton2, MODULE_NAME) <> vbYes Then
            Exit Sub
        End If
    End If

    Me.Undo
    modArticleGroupService.ApplyDefaultValues Me
    UpdateHeaderCaptionUi

    modLoggingHandler.LogInfo MODULE_NAME & ".CancelCurrentEdit", _
        "Unsaved Article Group changes were discarded."

    If NavigateBack Then
        ReturnToPreviousWorkspace "CancelCurrentEdit"
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CancelCurrentEdit", Err
End Sub

Private Sub RefreshCurrentRecord()
    On Error GoTo ErrorHandler

    If HasUnsavedChanges() Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_GROUP_SAVE_OR_CANCEL_FIRST", "Bitte zuerst speichern oder abbrechen."), vbExclamation, MODULE_NAME
        Exit Sub
    End If

    Me.Requery
    UpdateHeaderCaptionUi

    modLoggingHandler.LogInfo MODULE_NAME & ".RefreshCurrentRecord", _
        "Article Group detail requeried."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RefreshCurrentRecord", Err
End Sub

Private Sub UpdateHeaderCaptionUi()
    On Error GoTo SafeExit

    Dim headerCaption As String
    Dim displayName As String

    headerCaption = modFwTranslationRuntime.ResolveText("NAV.NEW_ARTICLE_GROUP", "Neue Artikelgruppe")

    If GetCurrentArticleGroupId() > 0 Then
        displayName = modArticleGroupService.ResolveArticleGroupDisplayName(Me)
        If LenB(displayName) > 0 Then
            headerCaption = modFwTranslationRuntime.ResolveText("FORM.FRMARTICLEGROUPDETAIL.FORM_TITLE", "Artikelgruppe") & ": " & displayName
        Else
            headerCaption = modFwTranslationRuntime.ResolveText("FORM.FRMARTICLEGROUPDETAIL.FORM_TITLE", "Artikelgruppe") & ": " & CStr(GetCurrentArticleGroupId())
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

Private Function GetCurrentArticleGroupId() As Long
    On Error GoTo SafeExit

    If Me.NewRecord Then
        Exit Function
    End If

    GetCurrentArticleGroupId = modArticleGroupService.ResolveArticleGroupId(Me)

SafeExit:
End Function

Private Sub RestoreCurrentArticleGroupSelection(ByVal productGroupId As Long)
    On Error GoTo ErrorHandler

    Dim recordsetClone As DAO.Recordset

    Set recordsetClone = Me.recordsetClone
    recordsetClone.FindFirst "[product_group_id]=" & CStr(productGroupId)

    If Not recordsetClone.NoMatch Then
        Me.Bookmark = recordsetClone.Bookmark
    End If

CleanExit:
    On Error Resume Next
    If Not recordsetClone Is Nothing Then recordsetClone.Close
    Set recordsetClone = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreCurrentArticleGroupSelection", Err
    Resume CleanExit
End Sub

Private Function HasControl(ByVal ControlName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentControl As Control

    For Each currentControl In Me.Controls
        If StrComp(currentControl.Name, ControlName, vbTextCompare) = 0 Then
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

