Option Compare Database
Option Explicit

'===============================================================================
' Module    : modAppShell
' Purpose   : Main application shell orchestration.
' Author    : Codex
' Version   : 0.3.0
'===============================================================================

Private Const MODULE_NAME As String = "modAppShell"
Private Const NAVIGATION_SUBFORM_CONTROL As String = "subNavigationHost"
Private Const WORKSPACE_SUBFORM_CONTROL As String = "subWorkspaceHost"
Private Const STATUS_APP_VERSION As String = "txtStatusAppVersion"
Private Const STATUS_CURRENT_USER As String = "txtStatusCurrentUser"
Private Const STATUS_CURRENT_TENANT As String = "txtStatusCurrentTenant"
Private Const STATUS_CURRENT_ROLE As String = "txtStatusCurrentRole"
Private Const STATUS_BACKEND As String = "txtStatusBackend"
Private Const STATUS_ENVIRONMENT As String = "txtStatusEnvironment"
Private Const HEADER_TITLE As String = "lblAppTitle"
Private Const HEADER_SUBTITLE As String = "lblAppSubtitle"
Private Const COMMAND_BACK As String = "cmdBack"
Private Const LABEL_USER As String = "lblStatusUser"
Private Const LABEL_TENANT As String = "lblStatusTenant"
Private Const LABEL_ROLE As String = "lblStatusRole"
Private Const LABEL_ENVIRONMENT As String = "lblStatusEnvironment"
Private Const LABEL_BACKEND As String = "lblStatusBackend"
Private Const COMMAND_HOME As String = "cmdHome"
Private Const LABEL_SHELL_SEARCH As String = "lblShellSearch"
Private Const TEXT_SHELL_SEARCH As String = "txtShellSearch"
Private Const COMMAND_SHELL_CLEAR_SEARCH As String = "cmdShellClearSearch"
Private Const COMMAND_SHELL_NEW As String = "cmdShellNew"
Private Const COMMAND_SHELL_EDIT As String = "cmdShellEdit"
Private Const COMMAND_SHELL_REFRESH As String = "cmdShellRefresh"

Private Const LIST_METHOD_SUPPORTS_BAR As String = "SupportsListCommandBar"
Private Const LIST_METHOD_SUPPORTS_NEW As String = "SupportsListNew"
Private Const LIST_METHOD_SUPPORTS_EDIT As String = "SupportsListEdit"
Private Const LIST_METHOD_SEARCH As String = "ListSearch"
Private Const LIST_METHOD_CLEAR_SEARCH As String = "ListClearSearch"
Private Const LIST_METHOD_NEW As String = "ListNew"
Private Const LIST_METHOD_EDIT As String = "ListEdit"
Private Const LIST_METHOD_REFRESH As String = "ListRefresh"

Private m_lastWorkspaceFormName As String

Public Function InitializeAppShell(ByVal shellForm As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    If shellForm Is Nothing Then
        Exit Function
    End If

    If Not modBootstrap.EnsureBootstrapped() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".InitializeAppShell", _
            "Bootstrap could not be completed for frmAppShell."
        Exit Function
    End If

    modFormRuntime.InitializeForm shellForm

    If Not modAppNavigationService.EnsureNavigationTables() Then
        Exit Function
    End If

    If Not modAppNavigationService.SeedDefaultNavigation() Then
        Exit Function
    End If

    LoadNavigationHost shellForm
    LoadDefaultWorkspace shellForm
    RefreshShellStatus shellForm

    InitializeAppShell = True

    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeAppShell", _
        "Application shell initialized successfully."
    Exit Function

ErrorHandler:
    InitializeAppShell = False
    modErrorHandler.HandleError MODULE_NAME, "InitializeAppShell", Err
End Function

Public Function RefreshShellStatus(ByVal shellForm As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    If shellForm Is Nothing Then
        RefreshShellStatus = True
        Exit Function
    End If

    ApplyShellStaticTranslations shellForm
    SetDisplayValueIfPresent shellForm, STATUS_APP_VERSION, APP_VERSION
    SetDisplayValueIfPresent shellForm, STATUS_CURRENT_USER, ResolveCurrentUserText()
    SetDisplayValueIfPresent shellForm, STATUS_CURRENT_TENANT, ResolveCurrentTenantText()
    SetDisplayValueIfPresent shellForm, STATUS_CURRENT_ROLE, ResolveCurrentRoleText()
    SetDisplayValueIfPresent shellForm, STATUS_BACKEND, ResolveBackendStatusText()
    SetDisplayValueIfPresent shellForm, STATUS_ENVIRONMENT, ResolveEnvironmentText()
    SetControlEnabledIfPresent shellForm, COMMAND_BACK, modAppWorkspaceService.CanGoBack()
    UpdateShellCommandBarState shellForm

    RefreshShellStatus = True
    Exit Function

ErrorHandler:
    RefreshShellStatus = False
    modErrorHandler.HandleError MODULE_NAME, "RefreshShellStatus", Err
End Function

Public Sub UpdateShellCommandBarState(ByVal shellForm As Access.Form)
    On Error GoTo ErrorHandler

    Dim workspaceForm As Access.Form
    Dim supportsCommandBar As Boolean
    Dim supportsNew As Boolean
    Dim supportsEdit As Boolean
    Dim currentWorkspaceFormName As String
    Dim clearSearch As Boolean

    Set workspaceForm = ResolveCurrentWorkspaceForm(shellForm)
    supportsCommandBar = ResolveWorkspaceBooleanCapability(workspaceForm, LIST_METHOD_SUPPORTS_BAR, False)

    If Not workspaceForm Is Nothing Then
        currentWorkspaceFormName = workspaceForm.Name
    End If

    clearSearch = (StrComp(m_lastWorkspaceFormName, currentWorkspaceFormName, vbTextCompare) <> 0)
    m_lastWorkspaceFormName = currentWorkspaceFormName

    If supportsCommandBar Then
        supportsNew = ResolveWorkspaceBooleanCapability(workspaceForm, LIST_METHOD_SUPPORTS_NEW, True)
        supportsEdit = ResolveWorkspaceBooleanCapability(workspaceForm, LIST_METHOD_SUPPORTS_EDIT, True)

        SetCommandBarVisible shellForm, True
        SetControlEnabledIfPresent shellForm, LABEL_SHELL_SEARCH, True
        SetControlEnabledIfPresent shellForm, TEXT_SHELL_SEARCH, True
        SetControlEnabledIfPresent shellForm, COMMAND_SHELL_CLEAR_SEARCH, True
        SetControlEnabledIfPresent shellForm, COMMAND_SHELL_NEW, supportsNew
        SetControlEnabledIfPresent shellForm, COMMAND_SHELL_EDIT, supportsEdit
        SetControlEnabledIfPresent shellForm, COMMAND_SHELL_REFRESH, True

        If clearSearch Then
            ClearShellSearchText shellForm
        End If
    Else
        SetCommandBarVisible shellForm, False
        ClearShellSearchText shellForm
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "UpdateShellCommandBarState", Err
End Sub

Public Sub ExecuteShellListSearch(ByVal shellForm As Access.Form, ByVal searchText As String)
    CallWorkspaceListMethodWithArg shellForm, LIST_METHOD_SEARCH, searchText
End Sub

Public Sub ExecuteShellListClearSearch(ByVal shellForm As Access.Form)
    ClearShellSearchText shellForm
    CallWorkspaceListMethodNoArg shellForm, LIST_METHOD_CLEAR_SEARCH
End Sub

Public Sub ExecuteShellListNew(ByVal shellForm As Access.Form)
    CallWorkspaceListMethodNoArg shellForm, LIST_METHOD_NEW
End Sub

Public Sub ExecuteShellListEdit(ByVal shellForm As Access.Form)
    CallWorkspaceListMethodNoArg shellForm, LIST_METHOD_EDIT
End Sub

Public Sub ExecuteShellListRefresh(ByVal shellForm As Access.Form)
    CallWorkspaceListMethodNoArg shellForm, LIST_METHOD_REFRESH
End Sub

Public Function LoadDefaultWorkspace(ByVal shellForm As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    LoadDefaultWorkspace = modAppWorkspaceService.LoadDashboard(shellForm)
    Exit Function

ErrorHandler:
    LoadDefaultWorkspace = False
    modErrorHandler.HandleError MODULE_NAME, "LoadDefaultWorkspace", Err
End Function

Public Function ResolveShellText( _
    ByVal translation_key As String, _
    ByVal fallback_text As String) As String
    On Error GoTo ErrorHandler

    Dim normalizedTranslationKey As String
    Dim translatedValue As String

    normalizedTranslationKey = Trim$(translation_key)
    If LenB(normalizedTranslationKey) = 0 Then
        ResolveShellText = fallback_text
        Exit Function
    End If

    translatedValue = LookupShellTranslationQuietly(normalizedTranslationKey)

    If LenB(Trim$(translatedValue)) = 0 Then
        ResolveShellText = fallback_text
    ElseIf StrComp(Trim$(translatedValue), normalizedTranslationKey, vbTextCompare) = 0 Then
        ResolveShellText = fallback_text
    ElseIf StrComp(Trim$(translatedValue), "TR:" & normalizedTranslationKey, vbTextCompare) = 0 Then
        ResolveShellText = fallback_text
    Else
        ResolveShellText = translatedValue
    End If
    Exit Function

ErrorHandler:
    ResolveShellText = fallback_text
    modErrorHandler.HandleError MODULE_NAME, "ResolveShellText", Err
End Function

Private Function LookupShellTranslationQuietly(ByVal translation_key As String) As String
    On Error GoTo ErrorHandler

    Dim currentLanguageCode As String
    Dim baseLanguageCode As String

    translation_key = Trim$(translation_key)
    If LenB(translation_key) = 0 Then
        Exit Function
    End If

    currentLanguageCode = modFwTranslationRuntime.GetCurrentLanguageCode()
    LookupShellTranslationQuietly = LookupShellTranslationByLanguage(translation_key, currentLanguageCode)
    If LenB(LookupShellTranslationQuietly) > 0 Then
        Exit Function
    End If

    baseLanguageCode = GetBaseLanguageCode(currentLanguageCode)
    If LenB(baseLanguageCode) > 0 Then
        If StrComp(baseLanguageCode, currentLanguageCode, vbTextCompare) <> 0 Then
            LookupShellTranslationQuietly = LookupShellTranslationByLanguage(translation_key, baseLanguageCode)
            If LenB(LookupShellTranslationQuietly) > 0 Then
                Exit Function
            End If
        End If
    End If

    If StrComp(currentLanguageCode, "EN", vbTextCompare) <> 0 Then
        LookupShellTranslationQuietly = LookupShellTranslationByLanguage(translation_key, "EN")
    End If
    Exit Function

ErrorHandler:
    LookupShellTranslationQuietly = vbNullString
End Function

Private Function LookupShellTranslationByLanguage( _
    ByVal translation_key As String, _
    ByVal languageCode As String) As String
    On Error GoTo ErrorHandler

    Dim lookupValue As Variant
    Dim criteria As String

    languageCode = Trim$(languageCode)
    If LenB(languageCode) = 0 Then
        Exit Function
    End If

    criteria = "translation_key = " & SqlText(translation_key) & _
               " AND language_code = " & SqlText(languageCode) & _
               " AND Nz(is_active, True) = True"

    lookupValue = DLookup("translation_value", "fw_translation", criteria)
    LookupShellTranslationByLanguage = Trim$(Nz(lookupValue, vbNullString))
    Exit Function

ErrorHandler:
    LookupShellTranslationByLanguage = vbNullString
End Function

Private Function GetBaseLanguageCode(ByVal languageCode As String) As String
    Dim separatorPosition As Long

    languageCode = Trim$(languageCode)
    separatorPosition = InStr(1, languageCode, "-", vbBinaryCompare)

    If separatorPosition > 0 Then
        GetBaseLanguageCode = Left$(languageCode, separatorPosition - 1)
    Else
        GetBaseLanguageCode = languageCode
    End If
End Function

Private Sub LoadNavigationHost(ByVal shellForm As Access.Form)
    On Error GoTo ErrorHandler

    If Not HasControl(shellForm, NAVIGATION_SUBFORM_CONTROL) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".LoadNavigationHost", _
            "Navigation host control '" & NAVIGATION_SUBFORM_CONTROL & "' was not found."
        Exit Sub
    End If

    shellForm.Controls(NAVIGATION_SUBFORM_CONTROL).SourceObject = vbNullString
    shellForm.Controls(NAVIGATION_SUBFORM_CONTROL).SourceObject = "Form.frmAppNavigation"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadNavigationHost", Err
End Sub

Private Sub ApplyShellStaticTranslations(ByVal shellForm As Access.Form)
    On Error GoTo SafeExit

    SetCaptionIfPresent shellForm, HEADER_TITLE, "FORM.FRMAPPSHELL.APP_TITLE", APP_NAME
    SetCaptionIfPresent shellForm, HEADER_SUBTITLE, "FORM.FRMAPPSHELL.APP_SUBTITLE", "Access Framework"
    SetCaptionIfPresent shellForm, COMMAND_HOME, "COMMON.HOME", "Home"
    SetCaptionIfPresent shellForm, COMMAND_BACK, "COMMON.BACK", "Zurueck"
    SetCaptionIfPresent shellForm, LABEL_SHELL_SEARCH, "COMMON.SEARCH", "Suche"
    SetCaptionIfPresent shellForm, COMMAND_SHELL_CLEAR_SEARCH, "COMMON.CLEAR_SEARCH", "Leeren"
    SetCaptionIfPresent shellForm, COMMAND_SHELL_NEW, "COMMON.NEW", "Neu"
    SetCaptionIfPresent shellForm, COMMAND_SHELL_EDIT, "COMMON.EDIT", "Bearbeiten"
    SetCaptionIfPresent shellForm, COMMAND_SHELL_REFRESH, "COMMON.REFRESH", "Aktualisieren"
    SetCaptionIfPresent shellForm, LABEL_USER, "FORM.FRMAPPSHELL.USER", "Benutzer"
    SetCaptionIfPresent shellForm, LABEL_TENANT, "FORM.FRMAPPSHELL.TENANT", "Mandant"
    SetCaptionIfPresent shellForm, LABEL_ROLE, "FORM.FRMAPPSHELL.ROLE", "Rolle"
    SetCaptionIfPresent shellForm, LABEL_ENVIRONMENT, "FORM.FRMAPPSHELL.ENVIRONMENT", "Umgebung"
    SetCaptionIfPresent shellForm, LABEL_BACKEND, "FORM.FRMAPPSHELL.BACKEND", "Backend"

SafeExit:
End Sub

Private Function ResolveCurrentUserText() As String
    If modSessionContext.IsSessionInitialized Then
        ResolveCurrentUserText = ResolveSafeText(modSessionContext.CurrentUserName)
    Else
        ResolveCurrentUserText = "n/a"
    End If
End Function

Private Function ResolveCurrentTenantText() As String
    If modTenantContext.IsTenantInitialized Then
        ResolveCurrentTenantText = ResolveSafeText(modTenantContext.CurrentTenantName)
        If LenB(Trim$(modTenantContext.currentTenantCode)) > 0 Then
            ResolveCurrentTenantText = ResolveCurrentTenantText & " (" & modTenantContext.currentTenantCode & ")"
        End If
    Else
        ResolveCurrentTenantText = "n/a"
    End If
End Function

Private Function ResolveCurrentRoleText() As String
    If modSessionContext.IsSessionInitialized Then
        ResolveCurrentRoleText = ResolveSafeText(modSessionContext.CurrentRoleCode)
    Else
        ResolveCurrentRoleText = "n/a"
    End If
End Function

Private Function ResolveBackendStatusText() As String
    Dim backendPath As String

    backendPath = Trim$(modDb.GetBackendPath())

    If modDb.ValidateBackendConfiguration() Then
        ResolveBackendStatusText = ResolveShellText("STATUS.READY", "Ready")
    Else
        ResolveBackendStatusText = "Unavailable"
    End If

    If LenB(backendPath) > 0 Then
        ResolveBackendStatusText = ResolveBackendStatusText & " | " & backendPath
    End If
End Function

Private Function ResolveEnvironmentText() As String
    ResolveEnvironmentText = ResolveSafeText(CurrentEnvironment)
End Function

Private Sub SetDisplayValueIfPresent( _
    ByVal formInstance As Access.Form, _
    ByVal ControlName As String, _
    ByVal displayValue As String)
    On Error GoTo SafeExit

    Dim ctl As Control

    If formInstance Is Nothing Then
        Exit Sub
    End If

    If Not HasControl(formInstance, ControlName) Then
        Exit Sub
    End If

    Set ctl = formInstance.Controls(ControlName)

    Select Case ctl.ControlType
        Case acLabel, acCommandButton, acCheckBox, acOptionButton, acToggleButton
            ctl.Caption = displayValue
        Case Else
            ctl.Value = displayValue
    End Select

SafeExit:
End Sub

Private Sub SetCaptionIfPresent( _
    ByVal formInstance As Access.Form, _
    ByVal ControlName As String, _
    ByVal translation_key As String, _
    ByVal fallback_text As String)
    On Error GoTo SafeExit

    If formInstance Is Nothing Then
        Exit Sub
    End If

    If Not HasControl(formInstance, ControlName) Then
        Exit Sub
    End If

    formInstance.Controls(ControlName).Caption = ResolveShellText(translation_key, fallback_text)

SafeExit:
End Sub

Private Sub SetControlEnabledIfPresent( _
    ByVal formInstance As Access.Form, _
    ByVal ControlName As String, _
    ByVal isEnabled As Boolean)
    On Error GoTo SafeExit

    If formInstance Is Nothing Then
        Exit Sub
    End If

    If Not HasControl(formInstance, ControlName) Then
        Exit Sub
    End If

    formInstance.Controls(ControlName).Enabled = isEnabled

SafeExit:
End Sub

Private Sub SetControlVisibleIfPresent( _
    ByVal formInstance As Access.Form, _
    ByVal ControlName As String, _
    ByVal isVisible As Boolean)
    On Error GoTo SafeExit

    If formInstance Is Nothing Then
        Exit Sub
    End If

    If Not HasControl(formInstance, ControlName) Then
        Exit Sub
    End If

    formInstance.Controls(ControlName).Visible = isVisible

SafeExit:
End Sub

Private Sub SetCommandBarVisible(ByVal shellForm As Access.Form, ByVal isVisible As Boolean)
    On Error GoTo SafeExit

    SetControlVisibleIfPresent shellForm, LABEL_SHELL_SEARCH, isVisible
    SetControlVisibleIfPresent shellForm, TEXT_SHELL_SEARCH, isVisible
    SetControlVisibleIfPresent shellForm, COMMAND_SHELL_CLEAR_SEARCH, isVisible
    SetControlVisibleIfPresent shellForm, COMMAND_SHELL_NEW, isVisible
    SetControlVisibleIfPresent shellForm, COMMAND_SHELL_EDIT, isVisible
    SetControlVisibleIfPresent shellForm, COMMAND_SHELL_REFRESH, isVisible

    SetControlEnabledIfPresent shellForm, LABEL_SHELL_SEARCH, isVisible
    SetControlEnabledIfPresent shellForm, TEXT_SHELL_SEARCH, isVisible
    SetControlEnabledIfPresent shellForm, COMMAND_SHELL_CLEAR_SEARCH, isVisible
    SetControlEnabledIfPresent shellForm, COMMAND_SHELL_NEW, isVisible
    SetControlEnabledIfPresent shellForm, COMMAND_SHELL_EDIT, isVisible
    SetControlEnabledIfPresent shellForm, COMMAND_SHELL_REFRESH, isVisible

SafeExit:
End Sub

Private Sub ClearShellSearchText(ByVal shellForm As Access.Form)
    On Error GoTo SafeExit

    If shellForm Is Nothing Then
        Exit Sub
    End If

    If HasControl(shellForm, TEXT_SHELL_SEARCH) Then
        shellForm.Controls(TEXT_SHELL_SEARCH).Value = vbNullString
    End If

SafeExit:
End Sub

Private Function ResolveCurrentWorkspaceForm(ByVal shellForm As Access.Form) As Access.Form
    On Error GoTo SafeExit

    Dim workspaceHost As Control

    If shellForm Is Nothing Then
        Exit Function
    End If

    If Not HasControl(shellForm, WORKSPACE_SUBFORM_CONTROL) Then
        Exit Function
    End If

    Set workspaceHost = shellForm.Controls(WORKSPACE_SUBFORM_CONTROL)
    If LenB(Trim$(Nz(workspaceHost.SourceObject, vbNullString))) = 0 Then
        Exit Function
    End If

    Set ResolveCurrentWorkspaceForm = workspaceHost.Form

SafeExit:
End Function

Private Function ResolveWorkspaceBooleanCapability( _
    ByVal workspaceForm As Access.Form, _
    ByVal methodName As String, _
    ByVal defaultValue As Boolean) As Boolean
    On Error GoTo SafeExit

    Dim formObject As Object

    ResolveWorkspaceBooleanCapability = defaultValue

    If workspaceForm Is Nothing Then
        ResolveWorkspaceBooleanCapability = False
        Exit Function
    End If

    Set formObject = workspaceForm
    ResolveWorkspaceBooleanCapability = CBool(CallByName(formObject, methodName, VbMethod))
    Exit Function

SafeExit:
    If Err.Number = 438 Then
        ResolveWorkspaceBooleanCapability = defaultValue
    ElseIf Not workspaceForm Is Nothing Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ResolveWorkspaceBooleanCapability", _
            "Capability '" & methodName & "' could not be resolved for form '" & workspaceForm.Name & "'."
    End If
End Function

Private Sub CallWorkspaceListMethodNoArg(ByVal shellForm As Access.Form, ByVal methodName As String)
    On Error GoTo ErrorHandler

    Dim workspaceForm As Access.Form
    Dim formObject As Object

    Set workspaceForm = ResolveCurrentWorkspaceForm(shellForm)
    If workspaceForm Is Nothing Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CallWorkspaceListMethodNoArg", _
            "No active workspace form is available for '" & methodName & "'."
        Exit Sub
    End If

    Set formObject = workspaceForm
    CallByName formObject, methodName, VbMethod
    RefreshShellStatus shellForm
    Exit Sub

ErrorHandler:
    HandleMissingWorkspaceListMethod shellForm, workspaceForm, methodName, Err
End Sub

Private Sub CallWorkspaceListMethodWithArg( _
    ByVal shellForm As Access.Form, _
    ByVal methodName As String, _
    ByVal argText As String)
    On Error GoTo ErrorHandler

    Dim workspaceForm As Access.Form
    Dim formObject As Object

    Set workspaceForm = ResolveCurrentWorkspaceForm(shellForm)
    If workspaceForm Is Nothing Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CallWorkspaceListMethodWithArg", _
            "No active workspace form is available for '" & methodName & "'."
        Exit Sub
    End If

    Set formObject = workspaceForm
    CallByName formObject, methodName, VbMethod, argText
    RefreshShellStatus shellForm
    Exit Sub

ErrorHandler:
    HandleMissingWorkspaceListMethod shellForm, workspaceForm, methodName, Err
End Sub

Private Sub HandleMissingWorkspaceListMethod( _
    ByVal shellForm As Access.Form, _
    ByVal workspaceForm As Access.Form, _
    ByVal methodName As String, _
    ByVal raisedError As ErrObject)
    On Error GoTo SafeExit

    If raisedError.Number = 438 Then
        If Not workspaceForm Is Nothing Then
            modLoggingHandler.LogWarning MODULE_NAME & ".HandleMissingWorkspaceListMethod", _
                "Workspace form '" & workspaceForm.Name & "' does not implement '" & methodName & "'."
        Else
            modLoggingHandler.LogWarning MODULE_NAME & ".HandleMissingWorkspaceListMethod", _
                "No workspace form available for '" & methodName & "'."
        End If
    Else
        modErrorHandler.HandleError MODULE_NAME, methodName, raisedError
    End If

    RefreshShellStatus shellForm

SafeExit:
End Sub

Private Function HasControl(ByVal formInstance As Access.Form, ByVal ControlName As String) As Boolean
    On Error GoTo SafeExit

    Dim ctl As Control

    If formInstance Is Nothing Then
        Exit Function
    End If

    For Each ctl In formInstance.Controls
        If StrComp(ctl.Name, ControlName, vbTextCompare) = 0 Then
            HasControl = True
            Exit Function
        End If
    Next ctl

SafeExit:
End Function

Private Function ResolveSafeText(ByVal Value As String) As String
    If LenB(Trim$(Value)) = 0 Then
        ResolveSafeText = "n/a"
    Else
        ResolveSafeText = Trim$(Value)
    End If
End Function
