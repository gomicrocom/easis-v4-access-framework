Attribute VB_Name = "modAppDashboardService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modAppDashboardService
' Purpose   : Dashboard value population for a standard unbound Access form.
' Author    : Codex
' Version   : 0.2.0
'===============================================================================

Private Const MODULE_NAME As String = "modAppDashboardService"
Private Const CONTROL_CARD_TENANT As String = "txtCardTenant"
Private Const CONTROL_CARD_USER As String = "txtCardUser"
Private Const CONTROL_CARD_BACKEND As String = "txtCardBackend"
Private Const CONTROL_CARD_FRAMEWORK As String = "txtCardFramework"
Private Const LABEL_CARD_TENANT As String = "lblCardTenantTitle"
Private Const LABEL_CARD_USER As String = "lblCardUserTitle"
Private Const LABEL_CARD_BACKEND As String = "lblCardBackendTitle"
Private Const LABEL_CARD_FRAMEWORK As String = "lblCardFrameworkTitle"
Private Const LABEL_CARD_STATUS As String = "lblCardStatusTitle"

Public Function RefreshDashboard(ByVal dashboardForm As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    If dashboardForm Is Nothing Then
        RefreshDashboard = True
        Exit Function
    End If

    ApplyDashboardTranslations dashboardForm
    SetDisplayValueIfPresent dashboardForm, CONTROL_CARD_TENANT, ResolveTenantStatusText()
    SetDisplayValueIfPresent dashboardForm, CONTROL_CARD_USER, ResolveUserStatusText()
    SetDisplayValueIfPresent dashboardForm, CONTROL_CARD_BACKEND, ResolveBackendStatusText()
    SetDisplayValueIfPresent dashboardForm, CONTROL_CARD_FRAMEWORK, ResolveFrameworkStatusText()

    RefreshDashboard = True

    modLoggingHandler.LogInfo MODULE_NAME & ".RefreshDashboard", _
        "Dashboard refreshed."
    Exit Function

ErrorHandler:
    RefreshDashboard = False
    modErrorHandler.HandleError MODULE_NAME, "RefreshDashboard", Err
End Function

Private Sub ApplyDashboardTranslations(ByVal dashboardForm As Access.Form)
    On Error GoTo SafeExit

    SetCaptionIfPresent dashboardForm, LABEL_CARD_TENANT, "FORM.FRMAPPDASHBOARD.TENANT", "Mandant"
    SetCaptionIfPresent dashboardForm, LABEL_CARD_USER, "FORM.FRMAPPDASHBOARD.USER", "Benutzer"
    SetCaptionIfPresent dashboardForm, LABEL_CARD_BACKEND, "FORM.FRMAPPDASHBOARD.BACKEND", "Backend"
    SetCaptionIfPresent dashboardForm, LABEL_CARD_FRAMEWORK, "FORM.FRMAPPDASHBOARD.FRAMEWORK", "Framework"
    SetCaptionIfPresent dashboardForm, LABEL_CARD_STATUS, "FORM.FRMAPPDASHBOARD.STATUS", "Status"

SafeExit:
End Sub

Private Function ResolveTenantStatusText() As String
    If modTenantContext.IsTenantInitialized Then
        ResolveTenantStatusText = ComposeDisplayText( _
            modTenantContext.CurrentTenantName, _
            modTenantContext.currentTenantCode)
    Else
        ResolveTenantStatusText = "n/a"
    End If
End Function

Private Function ResolveUserStatusText() As String
    If modSessionContext.IsSessionInitialized Then
        ResolveUserStatusText = ComposeDisplayText( _
            modSessionContext.CurrentUserName, _
            modSessionContext.CurrentRoleCode)
    Else
        ResolveUserStatusText = "n/a"
    End If
End Function

Private Function ResolveBackendStatusText() As String
    Dim backendPath As String

    backendPath = Trim$(modDb.GetBackendPath())

    If modDb.ValidateBackendConfiguration() Then
        ResolveBackendStatusText = "Ready"
        If LenB(backendPath) > 0 Then
            ResolveBackendStatusText = ResolveBackendStatusText & " | " & backendPath
        End If
    ElseIf LenB(backendPath) > 0 Then
        ResolveBackendStatusText = "Unavailable | " & backendPath
    Else
        ResolveBackendStatusText = "n/a"
    End If
End Function

Private Function ResolveFrameworkStatusText() As String
    Dim sessionText As String

    If modSessionContext.IsSessionInitialized Then
        sessionText = Format$(modSessionContext.SessionStartedAt, "yyyy-mm-dd hh:nn")
    Else
        sessionText = "n/a"
    End If

    ResolveFrameworkStatusText = "Bootstrap=" & IIf(IsBootstrapped, "ready", "pending") & _
                                 " | Env=" & NzText(CurrentEnvironment) & _
                                 " | Session=" & sessionText
End Function

Private Function ComposeDisplayText(ByVal primaryValue As String, ByVal secondaryValue As String) As String
    primaryValue = NzText(primaryValue)
    secondaryValue = NzText(secondaryValue)

    If LenB(secondaryValue) > 0 Then
        ComposeDisplayText = primaryValue & " | " & secondaryValue
    Else
        ComposeDisplayText = primaryValue
    End If
End Function

Private Sub SetDisplayValueIfPresent( _
    ByVal formInstance As Access.Form, _
    ByVal ControlName As String, _
    ByVal valueText As String)
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
            ctl.Caption = valueText
        Case Else
            ctl.Value = valueText
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

    formInstance.Controls(ControlName).Caption = modAppShell.ResolveShellText(translation_key, fallback_text)

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

Private Function NzText(ByVal Value As String, Optional ByVal DefaultValue As String = "n/a") As String
    If LenB(Trim$(Value)) = 0 Then
        NzText = DefaultValue
    Else
        NzText = Trim$(Value)
    End If
End Function
