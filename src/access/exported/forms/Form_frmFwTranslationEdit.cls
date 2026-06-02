VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFwTranslationEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmFwTranslationEdit
' Purpose   : Focused workspace editor for exactly one translation key across
'             the required framework languages.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmFwTranslationEdit"

Private m_currentTranslationKey As String
Private m_originalDeCh As String
Private m_originalEnUs As String
Private m_originalFrFr As String
Private m_isLoading As Boolean

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    modFormRuntime.InitializeForm Me
    modFwTranslationRuntime.ApplyTranslations Me
    ConfigureDeepLPlaceholder
    UpdateHeaderCaptionUi
    UpdateLocalCommandButtons

    If LenB(Trim$(modDaoHelper.NzString(Me.OpenArgs))) > 0 Then
        ApplyWorkspaceOpenArgs CStr(Me.OpenArgs)
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "Translation edit form loaded."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.TRANSLATION_EDIT_LOAD_ERROR", "Fehler beim Laden der Uebersetzung."), vbExclamation, MODULE_NAME
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

Private Sub cmdDeepL_Click()
    On Error GoTo ErrorHandler

    If LenB(Trim$(modDeepLTranslationService.ResolveDeepLApiKey())) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.DEEPL_API_KEY_MISSING", "DeepL API-Schluessel ist nicht konfiguriert."), vbExclamation, MODULE_NAME
        Exit Sub
    End If

    ApplyDeepLSuggestions
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdDeepL_Click", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.TRANSLATION_EDIT_DEEPL_ERROR", "Fehler beim Abrufen der DeepL-Vorschlaege."), vbExclamation, MODULE_NAME
End Sub

Private Sub txtTranslationDeCh_AfterUpdate()
    On Error GoTo ErrorHandler

    UpdateDeepLActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtTranslationDeCh_AfterUpdate", Err
End Sub

Private Sub txtTranslationDeCh_Change()
    On Error GoTo ErrorHandler

    UpdateDeepLActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtTranslationDeCh_Change", Err
End Sub

Private Sub txtTranslationEnUs_AfterUpdate()
    On Error GoTo ErrorHandler

    UpdateDeepLActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtTranslationEnUs_AfterUpdate", Err
End Sub

Private Sub txtTranslationEnUs_Change()
    On Error GoTo ErrorHandler

    UpdateDeepLActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtTranslationEnUs_Change", Err
End Sub

Private Sub txtTranslationFrFr_AfterUpdate()
    On Error GoTo ErrorHandler

    UpdateDeepLActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtTranslationFrFr_AfterUpdate", Err
End Sub

Private Sub txtTranslationFrFr_Change()
    On Error GoTo ErrorHandler

    UpdateDeepLActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtTranslationFrFr_Change", Err
End Sub

Public Function SupportsListCommandBar() As Boolean
    SupportsListCommandBar = False
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

Public Function CanLeaveWorkspace() As Boolean
    On Error GoTo ErrorHandler

    If Not HasUnsavedChanges() Then
        CanLeaveWorkspace = True
        Exit Function
    End If

    If MsgBox(modFwTranslationRuntime.ResolveText("MSG.TRANSLATION_EDIT_CANCEL_CONFIRM", "Ungespeicherte Aenderungen verwerfen?"), vbQuestion + vbYesNo + vbDefaultButton2, MODULE_NAME) = vbYes Then
        CancelCurrentEdit False, False
        CanLeaveWorkspace = True
    End If
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CanLeaveWorkspace", Err
End Function

Public Function GetCurrentWorkspaceRecordId() As Long
    GetCurrentWorkspaceRecordId = 0
End Function

Public Function GetWorkspaceState() As String
    Dim stateText As String

    stateText = AddWorkspaceStatePair(stateText, "translation_key", m_currentTranslationKey)
    GetWorkspaceState = stateText
End Function

Public Sub RestoreWorkspaceState(ByVal stateText As String)
    On Error GoTo ErrorHandler

    Dim translationKey As String

    translationKey = GetWorkspaceStateValue(stateText, "translation_key")
    If LenB(translationKey) > 0 Then
        LoadTranslationKey translationKey
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreWorkspaceState", Err
End Sub

Public Sub ApplyWorkspaceOpenArgs(ByVal openArgs As String)
    LoadTranslationKey openArgs
End Sub

Public Sub LoadTranslationKey(ByVal translationKey As String)
    On Error GoTo ErrorHandler

    translationKey = Trim$(modDaoHelper.NzString(translationKey))
    If LenB(translationKey) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.TRANSLATION_KEY_REQUIRED", "Uebersetzungsschluessel ist erforderlich."), vbExclamation, MODULE_NAME
        Exit Sub
    End If

    m_isLoading = True
    modFwTranslationEditService.LoadTranslationEditContext Me, translationKey
    m_currentTranslationKey = translationKey
    CaptureOriginalValues
    UpdateHeaderCaptionUi
    m_isLoading = False
    UpdateDeepLActionState

    modLoggingHandler.LogInfo MODULE_NAME & ".LoadTranslationKey", _
        "Loaded translation key '" & translationKey & "' into the editor."
    Exit Sub

ErrorHandler:
    m_isLoading = False
    modErrorHandler.HandleError MODULE_NAME, "LoadTranslationKey", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.TRANSLATION_EDIT_LOAD_ERROR", "Fehler beim Laden der Uebersetzung."), vbExclamation, MODULE_NAME
End Sub

Public Sub FocusTranslationKey(ByVal translationKey As String)
    LoadTranslationKey translationKey
End Sub

Public Function HasUnsavedChanges() As Boolean
    On Error GoTo SafeExit

    If m_isLoading Then
        Exit Function
    End If

    HasUnsavedChanges = _
        (StrComp(GetEditorText("txtTranslationDeCh"), m_originalDeCh, vbBinaryCompare) <> 0) Or _
        (StrComp(GetEditorText("txtTranslationEnUs"), m_originalEnUs, vbBinaryCompare) <> 0) Or _
        (StrComp(GetEditorText("txtTranslationFrFr"), m_originalFrFr, vbBinaryCompare) <> 0)

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

    If LenB(m_currentTranslationKey) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.TRANSLATION_KEY_REQUIRED", "Uebersetzungsschluessel ist erforderlich."), vbExclamation, MODULE_NAME
        Exit Sub
    End If

    modFwTranslationEditService.SaveTranslationValues Me, m_currentTranslationKey
    CaptureOriginalValues
    UpdateHeaderCaptionUi

    modLoggingHandler.LogInfo MODULE_NAME & ".SaveCurrentRecord", _
        "Saved translation editor values for key '" & m_currentTranslationKey & "'."

    ReturnToPreviousWorkspace "SaveCurrentRecord"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SaveCurrentRecord", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.TRANSLATION_EDIT_SAVE_ERROR", "Fehler beim Speichern der Uebersetzung."), vbExclamation, MODULE_NAME
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
        If MsgBox(modFwTranslationRuntime.ResolveText("MSG.TRANSLATION_EDIT_CANCEL_CONFIRM", "Ungespeicherte Aenderungen verwerfen?"), vbQuestion + vbYesNo + vbDefaultButton2, MODULE_NAME) <> vbYes Then
            Exit Sub
        End If
    End If

    ReloadCurrentTranslationContext

    modLoggingHandler.LogInfo MODULE_NAME & ".CancelCurrentEdit", _
        "Unsaved translation editor changes were discarded for key '" & m_currentTranslationKey & "'."

    If navigateBack Then
        ReturnToPreviousWorkspace "CancelCurrentEdit"
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CancelCurrentEdit", Err
End Sub

Private Sub ReloadCurrentTranslationContext()
    On Error GoTo SafeExit

    If LenB(m_currentTranslationKey) > 0 Then
        LoadTranslationKey m_currentTranslationKey
    End If

SafeExit:
End Sub

Private Sub CaptureOriginalValues()
    m_originalDeCh = GetEditorText("txtTranslationDeCh")
    m_originalEnUs = GetEditorText("txtTranslationEnUs")
    m_originalFrFr = GetEditorText("txtTranslationFrFr")
End Sub

Private Sub ConfigureDeepLPlaceholder()
    On Error GoTo SafeExit

    UpdateDeepLActionState

SafeExit:
End Sub

Private Sub UpdateHeaderCaptionUi()
    On Error GoTo SafeExit

    Dim headerCaption As String

    headerCaption = modFwTranslationRuntime.ResolveText("FORM.FRMFWTRANSLATIONEDIT.FORM_TITLE", "Uebersetzung bearbeiten")
    If LenB(m_currentTranslationKey) > 0 Then
        headerCaption = headerCaption & ": " & m_currentTranslationKey
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

Private Function GetControlText(ByVal controlName As String) As String
    On Error GoTo SafeExit

    If HasControl(controlName) Then
        GetControlText = modDaoHelper.NzString(Me.Controls(controlName).Value)
    End If

SafeExit:
End Function

Private Function GetEditorText(ByVal controlName As String) As String
    On Error GoTo SafeExit

    If Not HasControl(controlName) Then
        Exit Function
    End If

    GetEditorText = modDaoHelper.NzString(Me.Controls(controlName).Value)

    On Error Resume Next
    If StrComp(Me.ActiveControl.Name, controlName, vbTextCompare) = 0 Then
        GetEditorText = modDaoHelper.NzString(Me.Controls(controlName).Text, GetEditorText)
    End If
    On Error GoTo SafeExit

SafeExit:
End Function

Private Sub UpdateDeepLActionState()
    On Error GoTo SafeExit

    Dim hasSourceText As Boolean
    Dim hasMissingTarget As Boolean

    If Not HasControl("cmdDeepL") Then
        Exit Sub
    End If

    hasSourceText = (LenB(Trim$(GetEditorText("txtTranslationDeCh"))) > 0)
    hasMissingTarget = (LenB(Trim$(GetEditorText("txtTranslationEnUs"))) = 0) Or _
                       (LenB(Trim$(GetEditorText("txtTranslationFrFr"))) = 0)

    Me!cmdDeepL.Enabled = hasSourceText And hasMissingTarget

SafeExit:
End Sub

Private Sub ApplyDeepLSuggestions()
    On Error GoTo ErrorHandler

    Dim sourceText As String
    Dim enUsText As String
    Dim frFrText As String
    Dim overwriteFilledTargets As Boolean
    Dim suggestions As Object

    sourceText = Trim$(GetEditorText("txtTranslationDeCh"))
    enUsText = Trim$(GetEditorText("txtTranslationEnUs"))
    frFrText = Trim$(GetEditorText("txtTranslationFrFr"))

    If LenB(sourceText) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.TRANSLATION_EDIT_DEEPL_SOURCE_REQUIRED", "Bitte zuerst einen DE-CH Ausgangstext erfassen."), vbExclamation, MODULE_NAME
        Exit Sub
    End If

    If LenB(enUsText) > 0 And LenB(frFrText) > 0 Then
        If MsgBox(modFwTranslationRuntime.ResolveText("MSG.TRANSLATION_EDIT_DEEPL_OVERWRITE_CONFIRM", "Bestehende Zieltexte mit DeepL-Vorschlaegen ersetzen?"), vbQuestion + vbYesNo + vbDefaultButton2, MODULE_NAME) <> vbYes Then
            Exit Sub
        End If
        overwriteFilledTargets = True
    End If

    Set suggestions = modDeepLTranslationService.SuggestMissingTranslations( _
        sourceText, _
        "DE-CH", _
        enUsText, _
        frFrText, _
        BuildDeepLContextText(), _
        overwriteFilledTargets)

    If suggestions.Exists("EN-US") Then
        Me!txtTranslationEnUs.Value = suggestions.Item("EN-US")
    End If

    If suggestions.Exists("FR-FR") Then
        Me!txtTranslationFrFr.Value = suggestions.Item("FR-FR")
    End If

    UpdateDeepLActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyDeepLSuggestions", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function BuildDeepLContextText() As String
    Dim contextText As String

    contextText = AppendContextLine(contextText, "translation_key", GetControlText("txtTranslationKey"))
    contextText = AppendContextLine(contextText, "source_object", GetControlText("txtSourceObject"))
    contextText = AppendContextLine(contextText, "source_control", GetControlText("txtSourceControl"))
    contextText = AppendContextLine(contextText, "fallback_text", GetControlText("txtFallbackText"))

    BuildDeepLContextText = contextText
End Function

Private Function AppendContextLine(ByVal contextText As String, ByVal labelText As String, ByVal valueText As String) As String
    valueText = Trim$(modDaoHelper.NzString(valueText))
    If LenB(valueText) = 0 Then
        AppendContextLine = contextText
        Exit Function
    End If

    If LenB(contextText) > 0 Then
        contextText = contextText & vbCrLf
    End If

    AppendContextLine = contextText & labelText & ": " & valueText
End Function

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

    If Not Me.Parent Is Nothing Then
        If modAppWorkspaceService.GoBack(Me.Parent) Then
            Exit Sub
        End If
    End If

    modLoggingHandler.LogWarning MODULE_NAME & "." & callerName, _
        "Workspace history not available. Translation edit form remains open."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, callerName & ".ReturnToPreviousWorkspace", Err
End Sub
