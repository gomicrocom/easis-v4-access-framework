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

Private Const COMMAND_BAR_L1 As String = "cmdL1"
Private Const COMMAND_BAR_L2 As String = "cmdL2"
Private Const COMMAND_BAR_L3 As String = "cmdL3"
Private Const COMMAND_BAR_R5 As String = "cmdR5"
Private Const COMMAND_BAR_R4 As String = "cmdR4"
Private Const COMMAND_BAR_R3 As String = "cmdR3"
Private Const COMMAND_BAR_R2 As String = "cmdR2"
Private Const COMMAND_BAR_R1 As String = "cmdR1"
Private Const TEXT_QUICK_SEARCH As String = "txtQuickSearch"
Private Const COMMAND_QUICK_SEARCH_CLEAR As String = "cmdQuickSearchClear"

Private Const METHOD_SUPPORTS_WORKSPACE_BAR As String = "SupportsWorkspaceCommandBar"
Private Const METHOD_GET_WORKSPACE_BAR_CONFIG As String = "GetWorkspaceCommandBarConfig"
Private Const METHOD_CAN_EXECUTE_WORKSPACE_COMMAND As String = "CanExecuteWorkspaceCommand"
Private Const METHOD_EXECUTE_WORKSPACE_COMMAND As String = "ExecuteWorkspaceCommand"
Private Const LIST_METHOD_SUPPORTS_BAR As String = "SupportsListCommandBar"
Private Const LIST_METHOD_SUPPORTS_NEW As String = "SupportsListNew"
Private Const LIST_METHOD_SUPPORTS_EDIT As String = "SupportsListEdit"
Private Const LIST_METHOD_SEARCH As String = "ListSearch"
Private Const LIST_METHOD_CLEAR_SEARCH As String = "ListClearSearch"
Private Const LIST_METHOD_NEW As String = "ListNew"
Private Const LIST_METHOD_EDIT As String = "ListEdit"
Private Const LIST_METHOD_REFRESH As String = "ListRefresh"

Public Const WCMD_LIST_SEARCH As String = "LIST_SEARCH"
Public Const WCMD_LIST_CLEAR_SEARCH As String = "LIST_CLEAR_SEARCH"
Public Const WCMD_LIST_NEW As String = "LIST_NEW"
Public Const WCMD_LIST_EDIT As String = "LIST_EDIT"
Public Const WCMD_LIST_REFRESH As String = "LIST_REFRESH"
Public Const WCMD_NAV_HOME As String = "NAV_HOME"
Public Const WCMD_NAV_BACK As String = "NAV_BACK"
Public Const WCMD_DETAIL_SAVE As String = "DETAIL_SAVE"
Public Const WCMD_DETAIL_CANCEL As String = "DETAIL_CANCEL"

Public Const WSLOT_L1 As String = "L1"
Public Const WSLOT_L2 As String = "L2"
Public Const WSLOT_L3 As String = "L3"
Public Const WSLOT_R5 As String = "R5"
Public Const WSLOT_R4 As String = "R4"
Public Const WSLOT_R3 As String = "R3"
Public Const WSLOT_R2 As String = "R2"
Public Const WSLOT_R1 As String = "R1"

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
    RefreshWorkspaceCommandBar shellForm

    RefreshShellStatus = True
    Exit Function

ErrorHandler:
    RefreshShellStatus = False
    modErrorHandler.HandleError MODULE_NAME, "RefreshShellStatus", Err
End Function

Public Sub UpdateShellCommandBarState(ByVal shellForm As Access.Form)
    RefreshWorkspaceCommandBar shellForm
End Sub

Public Function RefreshWorkspaceCommandBar(ByVal shellForm As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    Dim workspaceForm As Access.Form
    Dim config As Object
    Dim currentWorkspaceFormName As String
    Dim clearSearch As Boolean
    Dim supportsWorkspaceBar As Boolean

    Set workspaceForm = ResolveCurrentWorkspaceForm(shellForm)
    If Not workspaceForm Is Nothing Then
        currentWorkspaceFormName = workspaceForm.Name
    End If

    clearSearch = (StrComp(m_lastWorkspaceFormName, currentWorkspaceFormName, vbTextCompare) <> 0)
    m_lastWorkspaceFormName = currentWorkspaceFormName

    ResetWorkspaceCommandBarUi shellForm

    supportsWorkspaceBar = SupportsWorkspaceCommandBarApi(workspaceForm)
    Set config = ResolveWorkspaceCommandBarConfig(workspaceForm, supportsWorkspaceBar)
    If config Is Nothing Then
        ClearQuickSearchText shellForm
        RefreshWorkspaceCommandBar = True
        Exit Function
    End If

    If clearSearch Then
        ClearQuickSearchText shellForm
    End If

    ApplyWorkspaceCommandBarConfig shellForm, workspaceForm, config
    RefreshWorkspaceCommandBar = True
    Exit Function

ErrorHandler:
    RefreshWorkspaceCommandBar = False
    modErrorHandler.HandleError MODULE_NAME, "RefreshWorkspaceCommandBar", Err
End Function

Public Sub ExecuteWorkspaceCommandBarSlot(ByVal shellForm As Access.Form, ByVal slotName As String)
    On Error GoTo ErrorHandler

    Dim commandKey As String

    commandKey = ResolveSlotCommandKey(shellForm, slotName)
    If LenB(commandKey) = 0 Then
        Exit Sub
    End If

    ExecuteWorkspaceCommandByKey shellForm, commandKey
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ExecuteWorkspaceCommandBarSlot", Err
End Sub

Public Sub ExecuteWorkspaceQuickSearch(ByVal shellForm As Access.Form, ByVal searchText As String)
    On Error GoTo ErrorHandler

    Dim commandKey As String

    commandKey = ResolveQuickSearchCommandKey(shellForm, False)
    If LenB(commandKey) = 0 Then
        commandKey = WCMD_LIST_SEARCH
    End If

    ExecuteWorkspaceCommandByKey shellForm, commandKey, searchText
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ExecuteWorkspaceQuickSearch", Err
End Sub

Public Sub ClearWorkspaceQuickSearch(ByVal shellForm As Access.Form)
    On Error GoTo ErrorHandler

    Dim commandKey As String

    ClearQuickSearchText shellForm

    commandKey = ResolveQuickSearchCommandKey(shellForm, True)
    If LenB(commandKey) = 0 Then
        commandKey = WCMD_LIST_CLEAR_SEARCH
    End If

    ExecuteWorkspaceCommandByKey shellForm, commandKey
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ClearWorkspaceQuickSearch", Err
End Sub

Public Sub ExecuteShellListSearch(ByVal shellForm As Access.Form, ByVal searchText As String)
    ExecuteWorkspaceQuickSearch shellForm, searchText
End Sub

Public Sub ExecuteShellListClearSearch(ByVal shellForm As Access.Form)
    ClearWorkspaceQuickSearch shellForm
End Sub

Public Sub ExecuteShellListNew(ByVal shellForm As Access.Form)
    ExecuteWorkspaceCommandByKey shellForm, WCMD_LIST_NEW
End Sub

Public Sub ExecuteShellListEdit(ByVal shellForm As Access.Form)
    ExecuteWorkspaceCommandByKey shellForm, WCMD_LIST_EDIT
End Sub

Public Sub ExecuteShellListRefresh(ByVal shellForm As Access.Form)
    ExecuteWorkspaceCommandByKey shellForm, WCMD_LIST_REFRESH
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

Private Sub ResetWorkspaceCommandBarUi(ByVal shellForm As Access.Form)
    On Error GoTo SafeExit

    ResetCommandBarSlot shellForm, WSLOT_L1
    ResetCommandBarSlot shellForm, WSLOT_L2
    ResetCommandBarSlot shellForm, WSLOT_L3
    ResetCommandBarSlot shellForm, WSLOT_R5
    ResetCommandBarSlot shellForm, WSLOT_R4
    ResetCommandBarSlot shellForm, WSLOT_R3
    ResetCommandBarSlot shellForm, WSLOT_R2
    ResetCommandBarSlot shellForm, WSLOT_R1

    SetControlVisibleIfPresent shellForm, TEXT_QUICK_SEARCH, False
    SetControlEnabledIfPresent shellForm, TEXT_QUICK_SEARCH, False
    SetControlVisibleIfPresent shellForm, COMMAND_QUICK_SEARCH_CLEAR, False
    SetControlEnabledIfPresent shellForm, COMMAND_QUICK_SEARCH_CLEAR, False
    ClearQuickSearchText shellForm

SafeExit:
End Sub

Private Sub ResetCommandBarSlot(ByVal shellForm As Access.Form, ByVal slotName As String)
    On Error GoTo SafeExit

    Dim controlName As String

    controlName = ResolveSlotControlName(slotName)
    If LenB(controlName) = 0 Then
        Exit Sub
    End If

    If Not HasControl(shellForm, controlName) Then
        Exit Sub
    End If

    shellForm.Controls(controlName).Visible = False
    shellForm.Controls(controlName).Enabled = False
    shellForm.Controls(controlName).Caption = vbNullString
    shellForm.Controls(controlName).Tag = vbNullString

SafeExit:
End Sub

Private Function ResolveWorkspaceCommandBarConfig( _
    ByVal workspaceForm As Access.Form, _
    ByVal supportsWorkspaceBar As Boolean) As Object
    On Error GoTo SafeExit

    Dim formObject As Object

    If workspaceForm Is Nothing Then
        Exit Function
    End If

    If supportsWorkspaceBar Then
        Set formObject = workspaceForm
        Set ResolveWorkspaceCommandBarConfig = CallByName(formObject, METHOD_GET_WORKSPACE_BAR_CONFIG, VbMethod)
        Exit Function
    End If

    Set ResolveWorkspaceCommandBarConfig = BuildLegacyListCommandBarConfig(workspaceForm)
    Exit Function

SafeExit:
    If Err.Number <> 0 And Err.Number <> 438 Then
        modErrorHandler.HandleError MODULE_NAME, "ResolveWorkspaceCommandBarConfig", Err
    End If
End Function

Private Function BuildLegacyListCommandBarConfig(ByVal workspaceForm As Access.Form) As Object
    On Error GoTo SafeExit

    Dim cfg As Object
    Dim supportsNew As Boolean
    Dim supportsEdit As Boolean

    If workspaceForm Is Nothing Then
        Exit Function
    End If

    If Not ResolveWorkspaceBooleanCapability(workspaceForm, LIST_METHOD_SUPPORTS_BAR, False, True) Then
        Exit Function
    End If

    supportsNew = ResolveWorkspaceBooleanCapability(workspaceForm, LIST_METHOD_SUPPORTS_NEW, True, True)
    supportsEdit = ResolveWorkspaceBooleanCapability(workspaceForm, LIST_METHOD_SUPPORTS_EDIT, True, True)

    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = vbTextCompare

    cfg("search.visible") = True
    cfg("search.enabled") = True
    cfg("search.command_change") = WCMD_LIST_SEARCH
    cfg("search.command_clear") = WCMD_LIST_CLEAR_SEARCH

    cfg(WSLOT_L1 & ".visible") = True
    cfg(WSLOT_L1 & ".enabled") = supportsNew
    cfg(WSLOT_L1 & ".caption") = ResolveShellText("COMMON.NEW", "Neu")
    cfg(WSLOT_L1 & ".command") = WCMD_LIST_NEW

    cfg(WSLOT_R3 & ".visible") = True
    cfg(WSLOT_R3 & ".enabled") = True
    cfg(WSLOT_R3 & ".caption") = ResolveShellText("COMMON.CLEAR_SEARCH", "Leeren")
    cfg(WSLOT_R3 & ".command") = WCMD_LIST_CLEAR_SEARCH

    cfg(WSLOT_R2 & ".visible") = True
    cfg(WSLOT_R2 & ".enabled") = supportsEdit
    cfg(WSLOT_R2 & ".caption") = ResolveShellText("COMMON.EDIT", "Bearbeiten")
    cfg(WSLOT_R2 & ".command") = WCMD_LIST_EDIT

    cfg(WSLOT_R1 & ".visible") = True
    cfg(WSLOT_R1 & ".enabled") = True
    cfg(WSLOT_R1 & ".caption") = ResolveShellText("COMMON.REFRESH", "Aktualisieren")
    cfg(WSLOT_R1 & ".command") = WCMD_LIST_REFRESH

    Set BuildLegacyListCommandBarConfig = cfg
    Exit Function

SafeExit:
    If Err.Number <> 0 Then
        modErrorHandler.HandleError MODULE_NAME, "BuildLegacyListCommandBarConfig", Err
    End If
End Function

Private Sub ApplyWorkspaceCommandBarConfig( _
    ByVal shellForm As Access.Form, _
    ByVal workspaceForm As Access.Form, _
    ByVal config As Object)
    On Error GoTo SafeExit

    If shellForm Is Nothing Then
        Exit Sub
    End If

    If config Is Nothing Then
        Exit Sub
    End If

    ApplyCommandBarSlot shellForm, workspaceForm, config, WSLOT_L1
    ApplyCommandBarSlot shellForm, workspaceForm, config, WSLOT_L2
    ApplyCommandBarSlot shellForm, workspaceForm, config, WSLOT_L3
    ApplyCommandBarSlot shellForm, workspaceForm, config, WSLOT_R5
    ApplyCommandBarSlot shellForm, workspaceForm, config, WSLOT_R4
    ApplyCommandBarSlot shellForm, workspaceForm, config, WSLOT_R3
    ApplyCommandBarSlot shellForm, workspaceForm, config, WSLOT_R2
    ApplyCommandBarSlot shellForm, workspaceForm, config, WSLOT_R1

    ApplyQuickSearchConfig shellForm, config

SafeExit:
End Sub

Private Sub ApplyCommandBarSlot( _
    ByVal shellForm As Access.Form, _
    ByVal workspaceForm As Access.Form, _
    ByVal config As Object, _
    ByVal slotName As String)
    On Error GoTo SafeExit

    Dim controlName As String
    Dim isVisible As Boolean
    Dim isEnabled As Boolean
    Dim captionText As String
    Dim commandKey As String

    controlName = ResolveSlotControlName(slotName)
    If LenB(controlName) = 0 Or Not HasControl(shellForm, controlName) Then
        Exit Sub
    End If

    isVisible = GetConfigBoolean(config, slotName & ".visible", False)
    commandKey = Trim$(GetConfigString(config, slotName & ".command", vbNullString))
    captionText = GetConfigString(config, slotName & ".caption", vbNullString)

    shellForm.Controls(controlName).Visible = isVisible
    shellForm.Controls(controlName).Tag = BuildControlCommandTag(commandKey)

    If Not isVisible Then
        shellForm.Controls(controlName).Enabled = False
        shellForm.Controls(controlName).Caption = vbNullString
        Exit Sub
    End If

    isEnabled = GetConfigBoolean(config, slotName & ".enabled", True)
    If LenB(commandKey) > 0 Then
        isEnabled = isEnabled And WorkspaceCanExecuteCommand(workspaceForm, commandKey)
    End If

    shellForm.Controls(controlName).Enabled = isEnabled
    shellForm.Controls(controlName).Caption = captionText

SafeExit:
End Sub

Private Sub ApplyQuickSearchConfig(ByVal shellForm As Access.Form, ByVal config As Object)
    On Error GoTo SafeExit

    Dim isVisible As Boolean
    Dim isEnabled As Boolean
    Dim currentValue As String
    Dim commandKey As String

    isVisible = GetConfigBoolean(config, "search.visible", False)
    isEnabled = GetConfigBoolean(config, "search.enabled", False)
    currentValue = GetConfigString(config, "search.value", vbNullString)
    commandKey = GetConfigString(config, "search.command_change", WCMD_LIST_SEARCH)

    If HasControl(shellForm, TEXT_QUICK_SEARCH) Then
        shellForm.Controls(TEXT_QUICK_SEARCH).Visible = isVisible
        shellForm.Controls(TEXT_QUICK_SEARCH).Enabled = isVisible And isEnabled
        shellForm.Controls(TEXT_QUICK_SEARCH).Tag = BuildControlCommandTag(commandKey)
        If LenB(currentValue) > 0 Then
            shellForm.Controls(TEXT_QUICK_SEARCH).Value = currentValue
        End If
    End If

    If HasControl(shellForm, COMMAND_QUICK_SEARCH_CLEAR) Then
        shellForm.Controls(COMMAND_QUICK_SEARCH_CLEAR).Visible = isVisible
        shellForm.Controls(COMMAND_QUICK_SEARCH_CLEAR).Enabled = isVisible And isEnabled
        shellForm.Controls(COMMAND_QUICK_SEARCH_CLEAR).Tag = BuildControlCommandTag(GetConfigString(config, "search.command_clear", WCMD_LIST_CLEAR_SEARCH))
    End If

SafeExit:
End Sub

Private Function WorkspaceCanExecuteCommand(ByVal workspaceForm As Access.Form, ByVal commandKey As String) As Boolean
    On Error GoTo SafeExit

    Dim formObject As Object

    WorkspaceCanExecuteCommand = True

    If workspaceForm Is Nothing Then
        WorkspaceCanExecuteCommand = False
        Exit Function
    End If

    If Not SupportsWorkspaceCommandBarApi(workspaceForm) Then
        WorkspaceCanExecuteCommand = ResolveLegacyCommandAvailability(workspaceForm, commandKey)
        Exit Function
    End If

    Set formObject = workspaceForm
    WorkspaceCanExecuteCommand = CBool(CallByName(formObject, METHOD_CAN_EXECUTE_WORKSPACE_COMMAND, VbMethod, commandKey))
    Exit Function

SafeExit:
    If Err.Number = 438 Then
        WorkspaceCanExecuteCommand = True
    Else
        WorkspaceCanExecuteCommand = False
    End If
End Function

Private Function ResolveLegacyCommandAvailability(ByVal workspaceForm As Access.Form, ByVal commandKey As String) As Boolean
    Select Case UCase$(Trim$(commandKey))
        Case WCMD_NAV_HOME, WCMD_NAV_BACK
            ResolveLegacyCommandAvailability = True
        Case WCMD_LIST_NEW
            ResolveLegacyCommandAvailability = ResolveWorkspaceBooleanCapability(workspaceForm, LIST_METHOD_SUPPORTS_NEW, True, True)
        Case WCMD_LIST_EDIT
            ResolveLegacyCommandAvailability = ResolveWorkspaceBooleanCapability(workspaceForm, LIST_METHOD_SUPPORTS_EDIT, True, True)
        Case WCMD_LIST_SEARCH, WCMD_LIST_CLEAR_SEARCH, WCMD_LIST_REFRESH
            ResolveLegacyCommandAvailability = True
        Case Else
            ResolveLegacyCommandAvailability = False
    End Select
End Function

Private Sub ExecuteWorkspaceCommandByKey( _
    ByVal shellForm As Access.Form, _
    ByVal commandKey As String, _
    Optional ByVal commandValue As String = "")
    On Error GoTo ErrorHandler

    Dim workspaceForm As Access.Form
    Dim formObject As Object

    Set workspaceForm = ResolveCurrentWorkspaceForm(shellForm)
    If workspaceForm Is Nothing Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ExecuteWorkspaceCommandByKey", _
            "No active workspace form is available for command '" & commandKey & "'."
        Exit Sub
    End If

    If SupportsWorkspaceCommandBarApi(workspaceForm) Then
        If WorkspaceCanExecuteCommand(workspaceForm, commandKey) Then
            Set formObject = workspaceForm
            CallByName formObject, METHOD_EXECUTE_WORKSPACE_COMMAND, VbMethod, commandKey, commandValue
        End If
    Else
        ExecuteLegacyWorkspaceCommand shellForm, workspaceForm, commandKey, commandValue
        Exit Sub
    End If

    RefreshShellStatus shellForm
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ExecuteWorkspaceCommandByKey", Err
End Sub

Private Function SupportsWorkspaceCommandBarApi(ByVal workspaceForm As Access.Form) As Boolean
    SupportsWorkspaceCommandBarApi = ResolveWorkspaceBooleanCapability(workspaceForm, METHOD_SUPPORTS_WORKSPACE_BAR, False, True)
End Function

Private Sub ExecuteLegacyWorkspaceCommand( _
    ByVal shellForm As Access.Form, _
    ByVal workspaceForm As Access.Form, _
    ByVal commandKey As String, _
    ByVal commandValue As String)
    Select Case UCase$(Trim$(commandKey))
        Case WCMD_NAV_HOME
            LoadDefaultWorkspace shellForm
        Case WCMD_NAV_BACK
            Call modAppWorkspaceService.GoBack(shellForm)
        Case WCMD_LIST_SEARCH
            CallWorkspaceListMethodWithArg shellForm, LIST_METHOD_SEARCH, commandValue
        Case WCMD_LIST_CLEAR_SEARCH
            CallWorkspaceListMethodNoArg shellForm, LIST_METHOD_CLEAR_SEARCH
        Case WCMD_LIST_NEW
            CallWorkspaceListMethodNoArg shellForm, LIST_METHOD_NEW
        Case WCMD_LIST_EDIT
            CallWorkspaceListMethodNoArg shellForm, LIST_METHOD_EDIT
        Case WCMD_LIST_REFRESH
            CallWorkspaceListMethodNoArg shellForm, LIST_METHOD_REFRESH
    End Select
End Sub

Private Function ResolveSlotCommandKey(ByVal shellForm As Access.Form, ByVal slotName As String) As String
    Dim controlName As String

    controlName = ResolveSlotControlName(slotName)
    If LenB(controlName) = 0 Then
        Exit Function
    End If

    If HasControl(shellForm, controlName) Then
        ResolveSlotCommandKey = ParseCommandKeyFromTag(modDaoHelper.NzString(shellForm.Controls(controlName).Tag))
    End If
End Function

Private Function ResolveQuickSearchCommandKey(ByVal shellForm As Access.Form, ByVal isClearCommand As Boolean) As String
    If isClearCommand Then
        If HasControl(shellForm, COMMAND_QUICK_SEARCH_CLEAR) Then
            ResolveQuickSearchCommandKey = ParseCommandKeyFromTag(modDaoHelper.NzString(shellForm.Controls(COMMAND_QUICK_SEARCH_CLEAR).Tag))
        End If
    Else
        If HasControl(shellForm, TEXT_QUICK_SEARCH) Then
            ResolveQuickSearchCommandKey = ParseCommandKeyFromTag(modDaoHelper.NzString(shellForm.Controls(TEXT_QUICK_SEARCH).Tag))
        End If
    End If
End Function

Private Function ResolveSlotControlName(ByVal slotName As String) As String
    Select Case UCase$(Trim$(slotName))
        Case WSLOT_L1
            ResolveSlotControlName = COMMAND_BAR_L1
        Case WSLOT_L2
            ResolveSlotControlName = COMMAND_BAR_L2
        Case WSLOT_L3
            ResolveSlotControlName = COMMAND_BAR_L3
        Case WSLOT_R5
            ResolveSlotControlName = COMMAND_BAR_R5
        Case WSLOT_R4
            ResolveSlotControlName = COMMAND_BAR_R4
        Case WSLOT_R3
            ResolveSlotControlName = COMMAND_BAR_R3
        Case WSLOT_R2
            ResolveSlotControlName = COMMAND_BAR_R2
        Case WSLOT_R1
            ResolveSlotControlName = COMMAND_BAR_R1
    End Select
End Function

Private Function BuildControlCommandTag(ByVal commandKey As String) As String
    commandKey = Trim$(commandKey)
    If LenB(commandKey) = 0 Then
        Exit Function
    End If

    BuildControlCommandTag = "WCMD=" & commandKey
End Function

Private Function ParseCommandKeyFromTag(ByVal tagValue As String) As String
    Dim parts() As String
    Dim partText As Variant
    Dim separatorPosition As Long

    If LenB(tagValue) = 0 Then
        Exit Function
    End If

    parts = Split(tagValue, ";")
    For Each partText In parts
        separatorPosition = InStr(1, CStr(partText), "=", vbBinaryCompare)
        If separatorPosition > 0 Then
            If StrComp(Left$(CStr(partText), separatorPosition - 1), "WCMD", vbTextCompare) = 0 Then
                ParseCommandKeyFromTag = Trim$(Mid$(CStr(partText), separatorPosition + 1))
                Exit Function
            End If
        End If
    Next partText
End Function

Private Function GetConfigBoolean(ByVal config As Object, ByVal keyName As String, ByVal defaultValue As Boolean) As Boolean
    On Error GoTo SafeExit

    If config Is Nothing Then
        GetConfigBoolean = defaultValue
        Exit Function
    End If

    If config.Exists(keyName) Then
        GetConfigBoolean = CBool(config(keyName))
    Else
        GetConfigBoolean = defaultValue
    End If
    Exit Function

SafeExit:
    GetConfigBoolean = defaultValue
End Function

Private Function GetConfigString(ByVal config As Object, ByVal keyName As String, ByVal defaultValue As String) As String
    On Error GoTo SafeExit

    If config Is Nothing Then
        GetConfigString = defaultValue
        Exit Function
    End If

    If config.Exists(keyName) Then
        GetConfigString = Trim$(modDaoHelper.NzString(config(keyName), defaultValue))
    Else
        GetConfigString = defaultValue
    End If
    Exit Function

SafeExit:
    GetConfigString = defaultValue
End Function

Private Sub ClearQuickSearchText(ByVal shellForm As Access.Form)
    On Error GoTo SafeExit

    If shellForm Is Nothing Then
        Exit Sub
    End If

    If HasControl(shellForm, TEXT_QUICK_SEARCH) Then
        shellForm.Controls(TEXT_QUICK_SEARCH).Value = vbNullString
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
    ByVal defaultValue As Boolean, _
    Optional ByVal suppressWarningLog As Boolean = False) As Boolean
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
    ElseIf Not suppressWarningLog And Not workspaceForm Is Nothing Then
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
