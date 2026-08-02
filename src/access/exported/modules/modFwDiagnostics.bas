Attribute VB_Name = "modFwDiagnostics"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFwDiagnostics
' Purpose   : Lightweight runtime diagnostics and health logging for shell,
'             workspace, translation, and form state investigation.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modFwDiagnostics"
Private Const DEFAULT_DIAGNOSTICS_ENABLED As Boolean = True
Private Const APPLICATION_DIAGNOSTICS_KEY As String = "DiagnosticsEnabled"
Private Const WORKSPACE_SUBFORM_CONTROL As String = "subWorkspaceHost"
Private Const NAVIGATION_SUBFORM_CONTROL As String = "subNavigationHost"
Private Const TEXT_QUICK_SEARCH As String = "txtQuickSearch"

Public Sub LogSystemSnapshot(ByVal contextName As String)
    On Error Resume Next

    If Not DiagnosticsEnabled() Then
        Exit Sub
    End If

    SafeLogInfo MODULE_NAME & ".LogSystemSnapshot", _
        "context=" & SafeToken(contextName) & _
        "; timestamp=" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
        "; app_name=" & SafeToken(APP_NAME) & _
        "; app_version=" & SafeToken(APP_VERSION) & _
        "; tenant_code=" & SafeToken(ResolveTenantCode()) & _
        "; current_user=" & SafeToken(ResolveCurrentUserName()) & _
        "; session_user_id=" & SafeToken(ResolveCurrentUserId()) & _
        "; session_started_at=" & SafeToken(ResolveSessionStartedAt()) & _
        "; language_code=" & SafeToken(modFwTranslationRuntime.GetCurrentLanguageCode()) & _
        "; open_forms_count=" & CStr(GetOpenFormsCount()) & _
        "; active_form=" & SafeToken(ResolveActiveFormName()) & _
        "; workspace_source_object=" & SafeToken(ResolveShellSourceObject(WORKSPACE_SUBFORM_CONTROL)) & _
        "; navigation_source_object=" & SafeToken(ResolveShellSourceObject(NAVIGATION_SUBFORM_CONTROL)) & _
        "; backend_path=" & SafeToken(modDb.GetCurrentTenantBackendPath()) & _
        "; log_level=" & SafeToken(CurrentLogLevel) & _
        "; environment=" & SafeToken(CurrentEnvironment)
End Sub

Public Sub LogOpenForms(ByVal contextName As String)
    On Error Resume Next

    Dim frm As Access.Form
    Dim recordsetState As String
    Dim sourceObject As String

    If Not DiagnosticsEnabled() Then
        Exit Sub
    End If

    SafeLogInfo MODULE_NAME & ".LogOpenForms", _
        "context=" & SafeToken(contextName) & "; open_forms_count=" & CStr(GetOpenFormsCount())

    For Each frm In Forms
        recordsetState = ResolveRecordsetState(frm)
        sourceObject = ResolveHostSourceObject(frm)

        SafeLogInfo MODULE_NAME & ".LogOpenForms", _
            "context=" & SafeToken(contextName) & _
            "; form_name=" & SafeToken(frm.Name) & _
            "; is_loaded=True" & _
            "; record_source=" & SafeToken(TryGetPropertyText(frm, "RecordSource")) & _
            "; dirty=" & BoolText(TryGetPropertyBoolean(frm, "Dirty")) & _
            "; new_record=" & BoolText(TryGetPropertyBoolean(frm, "NewRecord")) & _
            "; current_record=" & SafeToken(CStr(TryGetPropertyLong(frm, "CurrentRecord"))) & _
            "; recordset_state=" & SafeToken(recordsetState) & _
            "; host_source_object=" & SafeToken(sourceObject)
    Next frm
End Sub

Public Sub LogWorkspaceState(ByVal contextName As String, Optional ByVal shellForm As Variant)
    On Error Resume Next

    Dim effectiveShellForm As Access.Form

    If Not DiagnosticsEnabled() Then
        Exit Sub
    End If

    Set effectiveShellForm = ResolveShellForm(shellForm)

    SafeLogInfo MODULE_NAME & ".LogWorkspaceState", _
        "context=" & SafeToken(contextName) & _
        "; shell_form=" & SafeToken(ResolveFormName(effectiveShellForm)) & _
        "; workspace_source_object=" & SafeToken(GetControlPropertyText(effectiveShellForm, WORKSPACE_SUBFORM_CONTROL, "SourceObject")) & _
        "; navigation_source_object=" & SafeToken(GetControlPropertyText(effectiveShellForm, NAVIGATION_SUBFORM_CONTROL, "SourceObject")) & _
        "; can_go_back=" & BoolText(modAppWorkspaceService.CanGoBack())

    If Not effectiveShellForm Is Nothing Then
        LogCommandBarState contextName, effectiveShellForm
    End If
End Sub

Public Sub LogTranslationRuntimeState(ByVal contextName As String)
    On Error Resume Next

    Dim currentLanguageCode As String
    Dim tenantDefaultLanguage As String
    Dim iniLanguage As String
    Dim userLanguage As String
    Dim translationCount As Long
    Dim noteText As String

    If Not DiagnosticsEnabled() Then
        Exit Sub
    End If

    currentLanguageCode = modFwTranslationRuntime.GetCurrentLanguageCode()
    tenantDefaultLanguage = modTenantRepository.GetTenantParameter("DEFAULT_LANGUAGE", vbNullString)
    iniLanguage = modConfigIni.GetConfigValue(CONFIG_SECTION_APP, "Language", CurrentLanguage, ConfigFilePath)
    userLanguage = ResolveUserLanguageCode()
    translationCount = ResolveTranslationCount()

    If StrComp(currentLanguageCode, "de-CH", vbTextCompare) = 0 Then
        noteText = AppendNote(noteText, "Current language resolves to de-CH.")
    End If
    If LenB(Trim$(userLanguage)) = 0 Then
        noteText = AppendNote(noteText, "No user language resolved.")
    End If
    If LenB(Trim$(tenantDefaultLanguage)) = 0 Then
        noteText = AppendNote(noteText, "No tenant DEFAULT_LANGUAGE resolved.")
    End If
    If LenB(Trim$(iniLanguage)) = 0 Then
        noteText = AppendNote(noteText, "No INI Application.Language resolved.")
    End If

    SafeLogInfo MODULE_NAME & ".LogTranslationRuntimeState", _
        "context=" & SafeToken(contextName) & _
        "; current_language=" & SafeToken(currentLanguageCode) & _
        "; global_default_language=" & SafeToken(DEFAULT_LANGUAGE) & _
        "; tenant_default_language=" & SafeToken(tenantDefaultLanguage) & _
        "; ini_language=" & SafeToken(iniLanguage) & _
        "; user_language=" & SafeToken(userLanguage) & _
        "; translation_row_count=" & CStr(translationCount) & _
        "; note=" & SafeToken(noteText)
End Sub

Public Sub LogFormDiagnostics(ByVal contextName As String, ByVal formInstance As Access.Form)
    On Error Resume Next

    Dim ctl As Control
    Dim controlCount As Long
    Dim translatedTagCount As Long
    Dim captionWithoutTagCount As Long
    Dim currentTag As String
    Dim CaptionValue As String

    If Not DiagnosticsEnabled() Then
        Exit Sub
    End If

    If formInstance Is Nothing Then
        SafeLogInfo MODULE_NAME & ".LogFormDiagnostics", _
            "context=" & SafeToken(contextName) & "; form_instance=<nothing>"
        Exit Sub
    End If

    For Each ctl In formInstance.Controls
        controlCount = controlCount + 1
        currentTag = TryGetPropertyText(ctl, "Tag")
        CaptionValue = ResolveControlCaptionText(ctl)

        If InStr(1, currentTag, "TR:", vbTextCompare) > 0 Then
            translatedTagCount = translatedTagCount + 1
        ElseIf LenB(Trim$(CaptionValue)) > 0 Then
            captionWithoutTagCount = captionWithoutTagCount + 1
        End If
    Next ctl

    SafeLogInfo MODULE_NAME & ".LogFormDiagnostics", _
        "context=" & SafeToken(contextName) & _
        "; form_name=" & SafeToken(formInstance.Name) & _
        "; record_source=" & SafeToken(TryGetPropertyText(formInstance, "RecordSource")) & _
        "; filter=" & SafeToken(TryGetPropertyText(formInstance, "Filter")) & _
        "; filter_on=" & BoolText(TryGetPropertyBoolean(formInstance, "FilterOn")) & _
        "; control_count=" & CStr(controlCount) & _
        "; controls_with_tr_tag=" & CStr(translatedTagCount) & _
        "; controls_without_tr_tag_but_caption=" & CStr(captionWithoutTagCount) & _
        "; missing_translation_count_last_apply=n/a"
End Sub

Public Sub LogCommandBarState(ByVal contextName As String, ByVal shellForm As Access.Form)
    On Error Resume Next

    If Not DiagnosticsEnabled() Then
        Exit Sub
    End If

    If shellForm Is Nothing Then
        Exit Sub
    End If

    LogCommandBarControl contextName, shellForm, "cmdL1"
    LogCommandBarControl contextName, shellForm, "cmdL2"
    LogCommandBarControl contextName, shellForm, "cmdL3"
    LogCommandBarControl contextName, shellForm, "cmdR5"
    LogCommandBarControl contextName, shellForm, "cmdR4"
    LogCommandBarControl contextName, shellForm, "cmdR3"
    LogCommandBarControl contextName, shellForm, "cmdR2"
    LogCommandBarControl contextName, shellForm, "cmdR1"

    SafeLogInfo MODULE_NAME & ".LogCommandBarState", _
        "context=" & SafeToken(contextName) & _
        "; control=" & TEXT_QUICK_SEARCH & _
        "; visible=" & BoolText(GetControlPropertyBoolean(shellForm, TEXT_QUICK_SEARCH, "Visible")) & _
        "; enabled=" & BoolText(GetControlPropertyBoolean(shellForm, TEXT_QUICK_SEARCH, "Enabled")) & _
        "; value=" & SafeToken(GetControlPropertyText(shellForm, TEXT_QUICK_SEARCH, "Value"))

    SafeLogInfo MODULE_NAME & ".LogCommandBarState", _
        "context=" & SafeToken(contextName) & _
        "; workspace_source_object=" & SafeToken(GetControlPropertyText(shellForm, WORKSPACE_SUBFORM_CONTROL, "SourceObject")) & _
        "; navigation_source_object=" & SafeToken(GetControlPropertyText(shellForm, NAVIGATION_SUBFORM_CONTROL, "SourceObject"))
End Sub

Public Function DiagnosticsEnabled() As Boolean
    On Error Resume Next

    Static cachedValue As Variant
    Dim defaultValue As Boolean

    If Not IsEmpty(cachedValue) Then
        DiagnosticsEnabled = CBool(cachedValue)
        Exit Function
    End If

    defaultValue = DEFAULT_DIAGNOSTICS_ENABLED
    If StrComp(CurrentEnvironment, ENV_PROD, vbTextCompare) = 0 Then
        defaultValue = False
    End If

    cachedValue = modConfigIni.GetIniBoolean(CONFIG_SECTION_APP, APPLICATION_DIAGNOSTICS_KEY, defaultValue, ConfigFilePath)
    DiagnosticsEnabled = CBool(cachedValue)
End Function

Private Sub LogCommandBarControl(ByVal contextName As String, ByVal shellForm As Access.Form, ByVal ControlName As String)
    On Error Resume Next

    SafeLogInfo MODULE_NAME & ".LogCommandBarState", _
        "context=" & SafeToken(contextName) & _
        "; control=" & ControlName & _
        "; visible=" & BoolText(GetControlPropertyBoolean(shellForm, ControlName, "Visible")) & _
        "; enabled=" & BoolText(GetControlPropertyBoolean(shellForm, ControlName, "Enabled")) & _
        "; caption=" & SafeToken(GetControlPropertyText(shellForm, ControlName, "Caption")) & _
        "; tag=" & SafeToken(GetControlPropertyText(shellForm, ControlName, "Tag"))
End Sub

Private Function ResolveShellForm(Optional ByVal shellForm As Variant) As Access.Form
    On Error Resume Next

    If Not IsMissing(shellForm) Then
        If IsObject(shellForm) Then
            If Not shellForm Is Nothing Then
                Set ResolveShellForm = shellForm
                Exit Function
            End If
        End If
    End If

    If FormIsOpen("frmAppShell") Then
        Set ResolveShellForm = Forms("frmAppShell")
    End If
End Function

Private Function ResolveShellSourceObject(ByVal ControlName As String) As String
    On Error Resume Next

    Dim shellForm As Access.Form

    Set shellForm = ResolveShellForm()
    ResolveShellSourceObject = GetControlPropertyText(shellForm, ControlName, "SourceObject")
End Function

Private Function ResolveActiveFormName() As String
    On Error Resume Next

    ResolveActiveFormName = Screen.ActiveForm.Name
End Function

Private Function ResolveCurrentUserName() As String
    On Error Resume Next

    If modSessionContext.IsSessionInitialized Then
        ResolveCurrentUserName = modSessionContext.CurrentUserName
    End If
End Function

Private Function ResolveCurrentUserId() As String
    On Error Resume Next

    If modSessionContext.IsSessionInitialized Then
        ResolveCurrentUserId = modSessionContext.currentUserId
    End If
End Function

Private Function ResolveSessionStartedAt() As String
    On Error Resume Next

    If modSessionContext.IsSessionInitialized Then
        If modSessionContext.SessionStartedAt > 0 Then
            ResolveSessionStartedAt = Format$(modSessionContext.SessionStartedAt, "yyyy-mm-dd hh:nn:ss")
        End If
    End If
End Function

Private Function ResolveTenantCode() As String
    On Error Resume Next

    If modTenantContext.IsTenantInitialized Then
        ResolveTenantCode = modTenantContext.currentTenantCode
    End If
End Function

Private Function ResolveUserLanguageCode() As String
    On Error Resume Next

    If modSessionContext.IsSessionInitialized Then
        ResolveUserLanguageCode = Trim$(modUserRepository.GetUserLanguageCode(modSessionContext.currentUserId, vbNullString))
    End If
End Function

Private Function ResolveTranslationCount() As Long
    On Error Resume Next

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = modDb.GetSystemDatabase()
    If db Is Nothing Then
        Exit Function
    End If
    Set rs = db.OpenRecordset("SELECT Count(*) AS row_count FROM fw_translation", dbOpenSnapshot)
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            ResolveTranslationCount = CLng(Nz(rs.Fields(0).Value, 0))
        End If
        rs.Close
    End If
    db.Close
End Function

Private Function ResolveRecordsetState(ByVal formInstance As Access.Form) As String
    On Error Resume Next

    Dim rs As Object

    Set rs = formInstance.recordsetClone
    If rs Is Nothing Then
        ResolveRecordsetState = "n/a"
        Exit Function
    End If

    ResolveRecordsetState = "BOF=" & BoolText(rs.BOF) & ",EOF=" & BoolText(rs.EOF)
    rs.Close
End Function

Private Function ResolveHostSourceObject(ByVal formInstance As Access.Form) As String
    On Error Resume Next

    Dim ctl As Control
    Dim parts As String
    Dim sourceObject As String

    For Each ctl In formInstance.Controls
        sourceObject = TryGetPropertyText(ctl, "SourceObject")
        If LenB(sourceObject) > 0 Then
            If LenB(parts) > 0 Then
                parts = parts & "|"
            End If
            parts = parts & ctl.Name & ":" & sourceObject
        End If
    Next ctl

    ResolveHostSourceObject = parts
End Function

Private Function TryGetHostedForm(ByVal hostControl As Control) As Access.Form
    On Error Resume Next

    If hostControl Is Nothing Then
        Exit Function
    End If

    If LenB(Trim$(TryGetPropertyText(hostControl, "SourceObject"))) = 0 Then
        Exit Function
    End If

    Set TryGetHostedForm = hostControl.Form
End Function

Private Function ResolveControlCaptionText(ByVal ctl As Control) As String
    On Error Resume Next

    ResolveControlCaptionText = TryGetPropertyText(ctl, "Caption")
    If LenB(Trim$(ResolveControlCaptionText)) = 0 Then
        ResolveControlCaptionText = TryGetPropertyText(ctl, "ControlTipText")
    End If
End Function

Private Function GetOpenFormsCount() As Long
    On Error Resume Next

    GetOpenFormsCount = Forms.count
End Function

Private Function FormIsOpen(ByVal FormName As String) As Boolean
    On Error Resume Next

    Dim frm As Access.Form

    For Each frm In Forms
        If StrComp(frm.Name, FormName, vbTextCompare) = 0 Then
            FormIsOpen = True
            Exit Function
        End If
    Next frm
End Function

Private Function ResolveFormName(ByVal formInstance As Access.Form) As String
    On Error Resume Next

    If Not formInstance Is Nothing Then
        ResolveFormName = formInstance.Name
    End If
End Function

Private Function GetControlPropertyText(ByVal formInstance As Access.Form, ByVal ControlName As String, ByVal propertyName As String) As String
    On Error Resume Next

    If formInstance Is Nothing Then
        Exit Function
    End If

    GetControlPropertyText = TryGetPropertyText(formInstance.Controls(ControlName), propertyName)
End Function

Private Function GetControlPropertyBoolean(ByVal formInstance As Access.Form, ByVal ControlName As String, ByVal propertyName As String) As Boolean
    On Error Resume Next

    If formInstance Is Nothing Then
        Exit Function
    End If

    GetControlPropertyBoolean = TryGetPropertyBoolean(formInstance.Controls(ControlName), propertyName)
End Function

Private Function TryGetPropertyText(ByVal TargetObject As Object, ByVal propertyName As String) As String
    On Error Resume Next

    Dim Value As Variant

    If TargetObject Is Nothing Then
        Exit Function
    End If

    Value = CallByName(TargetObject, propertyName, VbGet)
    If Not IsNull(Value) Then
        TryGetPropertyText = CStr(Value)
    End If
End Function

Private Function TryGetPropertyBoolean(ByVal TargetObject As Object, ByVal propertyName As String) As Boolean
    On Error Resume Next

    Dim Value As Variant

    If TargetObject Is Nothing Then
        Exit Function
    End If

    Value = CallByName(TargetObject, propertyName, VbGet)
    If Not IsNull(Value) Then
        TryGetPropertyBoolean = CBool(Value)
    End If
End Function

Private Function TryGetPropertyLong(ByVal TargetObject As Object, ByVal propertyName As String) As Long
    On Error Resume Next

    Dim Value As Variant

    If TargetObject Is Nothing Then
        Exit Function
    End If

    Value = CallByName(TargetObject, propertyName, VbGet)
    If Not IsNull(Value) Then
        TryGetPropertyLong = CLng(Value)
    End If
End Function

Private Function BoolText(ByVal Value As Boolean) As String
    If Value Then
        BoolText = "True"
    Else
        BoolText = "False"
    End If
End Function

Private Function SafeToken(ByVal valueText As String) As String
    valueText = Replace(Trim$(Nz(valueText, vbNullString)), vbCr, " ")
    valueText = Replace(valueText, vbLf, " ")
    valueText = Replace(valueText, ";", ",")
    If LenB(valueText) = 0 Then
        SafeToken = "<empty>"
    Else
        SafeToken = valueText
    End If
End Function

Private Function AppendNote(ByVal noteText As String, ByVal notePart As String) As String
    notePart = Trim$(notePart)
    If LenB(notePart) = 0 Then
        AppendNote = noteText
    ElseIf LenB(noteText) = 0 Then
        AppendNote = notePart
    Else
        AppendNote = noteText & " " & notePart
    End If
End Function

Private Sub SafeLogInfo(ByVal sourceName As String, ByVal messageText As String)
    On Error Resume Next
    modLoggingHandler.LogInfo sourceName, messageText
End Sub
