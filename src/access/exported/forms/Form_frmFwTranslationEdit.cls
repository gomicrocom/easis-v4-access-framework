VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFwTranslationEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ExportKind: CODE_BEHIND_ONLY
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
    ConfigureSourceLanguageCombo
    ConfigureDeepLPlaceholder
    UpdateHeaderCaptionUi
    UpdateLocalCommandButtons

    If LenB(Trim$(modDaoHelper.NzString(Me.openArgs))) > 0 Then
        ApplyWorkspaceOpenArgs CStr(Me.openArgs)
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

Private Sub cboSourceLanguage_AfterUpdate()
    On Error GoTo ErrorHandler

    UpdateDeepLActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboSourceLanguage_AfterUpdate", Err
End Sub

Public Function SupportsListCommandBar() As Boolean
    SupportsListCommandBar = False
End Function

Public Function SupportsWorkspaceCommandBar() As Boolean
    SupportsWorkspaceCommandBar = True
End Function

Public Function GetWorkspaceCommandBarConfig() As Object
    On Error GoTo ErrorHandler

    Dim cfg As Object
    Dim shellForm As Access.Form
    Dim hasShellHost As Boolean

    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = vbTextCompare
    Set shellForm = ResolveHostingShellForm()
    hasShellHost = Not shellForm Is Nothing

    cfg("search.visible") = False
    cfg("search.enabled") = False

    cfg("L1.visible") = True
    cfg("L1.enabled") = hasShellHost
    cfg("L1.caption") = modFwTranslationRuntime.ResolveText("COMMON.HOME", "Home")
    cfg("L1.command") = modAppShell.WCMD_NAV_HOME

    cfg("L2.visible") = True
    cfg("L2.enabled") = hasShellHost And modAppWorkspaceService.CanGoBack()
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
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetWorkspaceCommandBarConfig", Err
End Function

Public Function CanExecuteWorkspaceCommand(ByVal commandKey As String) As Boolean
    On Error GoTo ErrorHandler

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()

    Select Case UCase$(Trim$(commandKey))
        Case modAppShell.WCMD_NAV_HOME
            CanExecuteWorkspaceCommand = Not shellForm Is Nothing
        Case modAppShell.WCMD_NAV_BACK
            CanExecuteWorkspaceCommand = (Not shellForm Is Nothing) And modAppWorkspaceService.CanGoBack()
        Case modAppShell.WCMD_LIST_REFRESH
            CanExecuteWorkspaceCommand = (LenB(m_currentTranslationKey) > 0)
        Case modAppShell.WCMD_DETAIL_SAVE, modAppShell.WCMD_DETAIL_CANCEL
            CanExecuteWorkspaceCommand = True
    End Select
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CanExecuteWorkspaceCommand", Err
End Function

Public Sub ExecuteWorkspaceCommand(ByVal commandKey As String, Optional ByVal commandValue As String = "")
    On Error GoTo ErrorHandler

    Select Case UCase$(Trim$(commandKey))
        Case modAppShell.WCMD_NAV_HOME
            NavigateHome
        Case modAppShell.WCMD_NAV_BACK
            NavigateBack
        Case modAppShell.WCMD_LIST_REFRESH
            RefreshCurrentTranslation
        Case modAppShell.WCMD_DETAIL_SAVE
            DetailSave
        Case modAppShell.WCMD_DETAIL_CANCEL
            DetailCancel
    End Select
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ExecuteWorkspaceCommand", Err
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
    Optional ByVal NavigateBack As Boolean = True)
    On Error GoTo ErrorHandler

    If Not HasUnsavedChanges() Then
        If NavigateBack Then
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

    If NavigateBack Then
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

Private Sub ConfigureSourceLanguageCombo()
    On Error GoTo SafeExit

    If Not HasControl("cboSourceLanguage") Then
        Exit Sub
    End If

    Me!cboSourceLanguage.RowSourceType = "Value List"
    Me!cboSourceLanguage.rowSource = "de-CH;en-US;fr-CH"

    If LenB(Trim$(modDaoHelper.NzString(Me!cboSourceLanguage.Value))) = 0 Then
        Me!cboSourceLanguage.Value = "de-CH"
    End If

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

Private Function ResolveHostingShellForm() As Access.Form
    On Error GoTo SafeExit

    Dim parentForm As Access.Form

    Set parentForm = Me.Parent
    If parentForm Is Nothing Then
        Exit Function
    End If

    If HasControlOnForm(parentForm, "subWorkspaceHost") Then
        Set ResolveHostingShellForm = parentForm
    End If

SafeExit:
End Function

Private Function GetControlText(ByVal ControlName As String) As String
    On Error GoTo SafeExit

    If HasControl(ControlName) Then
        GetControlText = modDaoHelper.NzString(Me.Controls(ControlName).Value)
    End If

SafeExit:
End Function

Private Function GetEditorText(ByVal ControlName As String) As String
    On Error GoTo SafeExit

    If Not HasControl(ControlName) Then
        Exit Function
    End If

    GetEditorText = modDaoHelper.NzString(Me.Controls(ControlName).Value)

    On Error Resume Next
    If StrComp(Me.ActiveControl.Name, ControlName, vbTextCompare) = 0 Then
        GetEditorText = modDaoHelper.NzString(Me.Controls(ControlName).text, GetEditorText)
    End If
    On Error GoTo SafeExit

SafeExit:
End Function

Private Sub UpdateDeepLActionState()
    On Error GoTo SafeExit

    Dim hasSourceText As Boolean
    Dim hasMissingTarget As Boolean
    Dim sourceLanguageCode As String

    If Not HasControl("cmdDeepL") Then
        Exit Sub
    End If

    sourceLanguageCode = ResolveSelectedSourceLanguage()
    hasSourceText = Not IsEmptyTranslationValue(GetTranslationValueByLanguage(sourceLanguageCode))
    hasMissingTarget = HasEmptyTargetTranslation(sourceLanguageCode)

    Me!cmdDeepl.Enabled = hasSourceText And hasMissingTarget

SafeExit:
End Sub

Private Sub ApplyDeepLSuggestions()
    On Error GoTo ErrorHandler

    Dim sourceLanguageCode As String
    Dim sourceText As String
    Dim targetLanguageCodes As Collection
    Dim targetLanguageCode As Variant
    Dim rawTargetText As String
    Dim hasMissingTarget As Boolean
    Dim hasExistingTargets As Boolean
    Dim overwriteFilledTargets As Boolean
    Dim targetSuggestions As Object

    sourceLanguageCode = ResolveSelectedSourceLanguage()
    sourceText = NormalizeTranslationValueForDeepL(GetTranslationValueByLanguage(sourceLanguageCode))
    Set targetLanguageCodes = BuildTargetLanguageCodes(sourceLanguageCode)

    If IsEmptyTranslationValue(sourceText) Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.TRANSLATION_EDIT_DEEPL_SOURCE_REQUIRED", "Bitte zuerst einen Text in der Ausgangssprache erfassen."), vbExclamation, MODULE_NAME
        Exit Sub
    End If

    hasMissingTarget = False
    hasExistingTargets = False

    For Each targetLanguageCode In targetLanguageCodes
        rawTargetText = GetTranslationValueByLanguage(CStr(targetLanguageCode))
        If IsEmptyTranslationValue(rawTargetText) Then
            hasMissingTarget = True
        Else
            hasExistingTargets = True
        End If
    Next targetLanguageCode

    If (Not hasMissingTarget) And hasExistingTargets Then
        If MsgBox(modFwTranslationRuntime.ResolveText("MSG.TRANSLATION_EDIT_DEEPL_OVERWRITE_CONFIRM", "Bestehende Zieltexte mit DeepL-Vorschlaegen ersetzen?"), vbQuestion + vbYesNo + vbDefaultButton2, MODULE_NAME) <> vbYes Then
            Exit Sub
        End If
        overwriteFilledTargets = True
    End If

    Set targetSuggestions = CreateObject("Scripting.Dictionary")

    For Each targetLanguageCode In targetLanguageCodes
        rawTargetText = GetTranslationValueByLanguage(CStr(targetLanguageCode))
        If overwriteFilledTargets Or IsEmptyTranslationValue(rawTargetText) Then
            targetSuggestions(CStr(targetLanguageCode)) = modDeepLTranslationService.TranslateText( _
                sourceText, _
                sourceLanguageCode, _
                CStr(targetLanguageCode), _
                BuildDeepLContextText())
        End If
    Next targetLanguageCode

    For Each targetLanguageCode In targetSuggestions.Keys
        SetTranslationValueByLanguage CStr(targetLanguageCode), CStr(targetSuggestions(CStr(targetLanguageCode)))
    Next targetLanguageCode

    UpdateDeepLActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyDeepLSuggestions", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Function IsEmptyTranslationValue(ByVal valueText As String) As Boolean
    valueText = Trim$(modDaoHelper.NzString(valueText))
    IsEmptyTranslationValue = (LenB(valueText) = 0) Or IsPlaceholderTranslationValue(valueText)
End Function

Private Function IsPlaceholderTranslationValue(ByVal valueText As String) As Boolean
    valueText = Trim$(modDaoHelper.NzString(valueText))
    IsPlaceholderTranslationValue = (StrComp(valueText, "<neu>", vbTextCompare) = 0)
End Function

Private Function NormalizeTranslationValueForDeepL(ByVal valueText As String) As String
    valueText = Trim$(modDaoHelper.NzString(valueText))

    If IsEmptyTranslationValue(valueText) Then
        NormalizeTranslationValueForDeepL = vbNullString
    Else
        NormalizeTranslationValueForDeepL = valueText
    End If
End Function

Private Function ResolveSelectedSourceLanguage() As String
    Dim sourceLanguageCode As String

    sourceLanguageCode = UCase$(Trim$(GetControlText("cboSourceLanguage")))

    Select Case sourceLanguageCode
        Case "de-CH", "en-US", "fr-CH"
            ResolveSelectedSourceLanguage = sourceLanguageCode
        Case Else
            ResolveSelectedSourceLanguage = "de-CH"
    End Select
End Function

Private Function BuildTargetLanguageCodes(ByVal sourceLanguageCode As String) As Collection
    Dim languageCodes As Collection

    Set languageCodes = New Collection

    If StrComp(sourceLanguageCode, "de-CH", vbTextCompare) <> 0 Then
        languageCodes.Add "de-CH"
    End If

    If StrComp(sourceLanguageCode, "en-US", vbTextCompare) <> 0 Then
        languageCodes.Add "en-US"
    End If

    If StrComp(sourceLanguageCode, "fr-CH", vbTextCompare) <> 0 Then
        languageCodes.Add "fr-CH"
    End If

    Set BuildTargetLanguageCodes = languageCodes
End Function

Private Function GetTranslationControlName(ByVal languageCode As String) As String
    Select Case UCase$(Trim$(languageCode))
        Case "de-CH"
            GetTranslationControlName = "txtTranslationDeCh"
        Case "en-US"
            GetTranslationControlName = "txtTranslationEnUs"
        Case "fr-CH"
            GetTranslationControlName = "txtTranslationFrFr"
    End Select
End Function

Private Function GetTranslationValueByLanguage(ByVal languageCode As String) As String
    Dim ControlName As String

    ControlName = GetTranslationControlName(languageCode)
    If LenB(ControlName) = 0 Then
        Exit Function
    End If

    GetTranslationValueByLanguage = GetEditorText(ControlName)
End Function

Private Sub SetTranslationValueByLanguage(ByVal languageCode As String, ByVal valueText As String)
    Dim ControlName As String

    ControlName = GetTranslationControlName(languageCode)
    If LenB(ControlName) = 0 Or Not HasControl(ControlName) Then
        Exit Sub
    End If

    Me.Controls(ControlName).Value = valueText
End Sub

Private Function HasEmptyTargetTranslation(ByVal sourceLanguageCode As String) As Boolean
    Dim targetLanguageCode As Variant
    Dim targetLanguageCodes As Collection

    Set targetLanguageCodes = BuildTargetLanguageCodes(sourceLanguageCode)

    For Each targetLanguageCode In targetLanguageCodes
        If IsEmptyTranslationValue(GetTranslationValueByLanguage(CStr(targetLanguageCode))) Then
            HasEmptyTargetTranslation = True
            Exit Function
        End If
    Next targetLanguageCode
End Function

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

Private Function HasControlOnForm(ByVal targetForm As Access.Form, ByVal ControlName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentControl As Control

    If targetForm Is Nothing Then
        Exit Function
    End If

    For Each currentControl In targetForm.Controls
        If StrComp(currentControl.Name, ControlName, vbTextCompare) = 0 Then
            HasControlOnForm = True
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

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()
    If Not shellForm Is Nothing Then
        If modAppWorkspaceService.GoBack(shellForm) Then
            Exit Sub
        End If
    End If

    modLoggingHandler.LogWarning MODULE_NAME & "." & callerName, _
        "Workspace history not available. Translation edit form remains open."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, callerName & ".ReturnToPreviousWorkspace", Err
End Sub

Private Sub NavigateHome()
    On Error GoTo ErrorHandler

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()
    If shellForm Is Nothing Then
        Exit Sub
    End If

    modAppShell.LoadDefaultWorkspace shellForm
    modAppShell.RefreshShellStatus shellForm
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NavigateHome", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub NavigateBack()
    On Error GoTo ErrorHandler

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()
    If shellForm Is Nothing Then
        Exit Sub
    End If

    If modAppWorkspaceService.GoBack(shellForm) Then
        modAppShell.RefreshShellStatus shellForm
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NavigateBack", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub RefreshCurrentTranslation()
    On Error GoTo ErrorHandler

    If LenB(m_currentTranslationKey) = 0 Then
        Exit Sub
    End If

    If HasUnsavedChanges() Then
        If MsgBox(modFwTranslationRuntime.ResolveText("MSG.TRANSLATION_EDIT_CANCEL_CONFIRM", "Ungespeicherte Aenderungen verwerfen?"), vbQuestion + vbYesNo + vbDefaultButton2, MODULE_NAME) <> vbYes Then
            Exit Sub
        End If
    End If

    ReloadCurrentTranslationContext
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RefreshCurrentTranslation", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

