Option Compare Database
Option Explicit

'===============================================================================
' Module    : modAppShell
' Purpose   : Main application shell orchestration.
' Author    : Codex
' Version   : 0.2.0
'===============================================================================

Private Const MODULE_NAME As String = "modAppShell"
Private Const NAVIGATION_SUBFORM_CONTROL As String = "subNavigationHost"
Private Const STATUS_APP_VERSION As String = "txtStatusAppVersion"
Private Const STATUS_CURRENT_USER As String = "txtStatusCurrentUser"
Private Const STATUS_CURRENT_TENANT As String = "txtStatusCurrentTenant"
Private Const STATUS_CURRENT_ROLE As String = "txtStatusCurrentRole"
Private Const STATUS_BACKEND As String = "txtStatusBackend"
Private Const STATUS_ENVIRONMENT As String = "txtStatusEnvironment"
Private Const HEADER_TITLE As String = "lblAppTitle"
Private Const COMMAND_BACK As String = "cmdBack"

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

    SetDisplayValueIfPresent shellForm, HEADER_TITLE, APP_NAME
    SetDisplayValueIfPresent shellForm, STATUS_APP_VERSION, APP_VERSION
    SetDisplayValueIfPresent shellForm, STATUS_CURRENT_USER, ResolveCurrentUserText()
    SetDisplayValueIfPresent shellForm, STATUS_CURRENT_TENANT, ResolveCurrentTenantText()
    SetDisplayValueIfPresent shellForm, STATUS_CURRENT_ROLE, ResolveCurrentRoleText()
    SetDisplayValueIfPresent shellForm, STATUS_BACKEND, ResolveBackendStatusText()
    SetDisplayValueIfPresent shellForm, STATUS_ENVIRONMENT, ResolveEnvironmentText()
    SetControlEnabledIfPresent shellForm, COMMAND_BACK, modAppWorkspaceService.CanGoBack()

    RefreshShellStatus = True
    Exit Function

ErrorHandler:
    RefreshShellStatus = False
    modErrorHandler.HandleError MODULE_NAME, "RefreshShellStatus", Err
End Function

Public Function LoadDefaultWorkspace(ByVal shellForm As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    LoadDefaultWorkspace = modAppWorkspaceService.LoadDashboard(shellForm)
    Exit Function

ErrorHandler:
    LoadDefaultWorkspace = False
    modErrorHandler.HandleError MODULE_NAME, "LoadDefaultWorkspace", Err
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
        ResolveBackendStatusText = "Ready"
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