Option Compare Database
Option Explicit

'===============================================================================
' Module    : modAppWorkspaceService
' Purpose   : Workspace host helpers for the application shell.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modAppWorkspaceService"
Private Const WORKSPACE_SUBFORM_CONTROL As String = "subWorkspaceHost"
Private Const DASHBOARD_FORM_NAME As String = "frmAppDashboard"

Public Function OpenWorkspaceForm( _
    ByVal shellForm As Access.Form, _
    ByVal form_name As String, _
    Optional ByVal where_condition As String = "") As Boolean
    On Error GoTo ErrorHandler

    Dim hostForm As Access.Form
    Dim workspaceHost As Control

    If LenB(Trim$(form_name)) = 0 Then
        Exit Function
    End If

    If Not FormExists(form_name) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".OpenWorkspaceForm", _
            "Workspace form '" & form_name & "' was not found."
        Exit Function
    End If

    Set hostForm = ResolveWorkspaceHostForm(shellForm)

    If hostForm Is Nothing Then
        DoCmd.OpenForm form_name, acNormal, , where_condition
        OpenWorkspaceForm = True
        modLoggingHandler.LogInfo MODULE_NAME & ".OpenWorkspaceForm", _
            "Opened form '" & form_name & "' without shell host."
        Exit Function
    End If

    Set workspaceHost = GetWorkspaceHostControl(hostForm)
    If workspaceHost Is Nothing Then
        DoCmd.OpenForm form_name, acNormal, , where_condition
        OpenWorkspaceForm = True
        modLoggingHandler.LogWarning MODULE_NAME & ".OpenWorkspaceForm", _
            "Shell workspace host is missing. Opened form '" & form_name & "' standalone."
        Exit Function
    End If

    workspaceHost.SourceObject = vbNullString
    workspaceHost.SourceObject = "Form." & form_name

    If LenB(Trim$(where_condition)) > 0 Then
        workspaceHost.Form.Filter = where_condition
        workspaceHost.Form.FilterOn = True
    Else
        workspaceHost.Form.FilterOn = False
        workspaceHost.Form.Filter = vbNullString
    End If

    OpenWorkspaceForm = True
    modLoggingHandler.LogInfo MODULE_NAME & ".OpenWorkspaceForm", _
        "Loaded form '" & form_name & "' into the shell workspace."
    Exit Function

ErrorHandler:
    OpenWorkspaceForm = False
    modErrorHandler.HandleError MODULE_NAME, "OpenWorkspaceForm", Err
End Function

Public Function PreviewWorkspaceReport( _
    ByVal shellForm As Access.Form, _
    ByVal report_name As String, _
    Optional ByVal where_condition As String = "") As Boolean
    On Error GoTo ErrorHandler

    If LenB(Trim$(report_name)) = 0 Then
        Exit Function
    End If

    If Not ReportExists(report_name) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".PreviewWorkspaceReport", _
            "Workspace report '" & report_name & "' was not found."
        Exit Function
    End If

    DoCmd.OpenReport report_name, acViewPreview, , where_condition

    PreviewWorkspaceReport = True
    modLoggingHandler.LogInfo MODULE_NAME & ".PreviewWorkspaceReport", _
        "Opened report '" & report_name & "' in preview mode."
    Exit Function

ErrorHandler:
    PreviewWorkspaceReport = False
    modErrorHandler.HandleError MODULE_NAME, "PreviewWorkspaceReport", Err
End Function

Public Function LoadDashboard(ByVal shellForm As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    LoadDashboard = OpenWorkspaceForm(shellForm, DASHBOARD_FORM_NAME)
    Exit Function

ErrorHandler:
    LoadDashboard = False
    modErrorHandler.HandleError MODULE_NAME, "LoadDashboard", Err
End Function

Public Function ClearWorkspace(ByVal shellForm As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    Dim hostForm As Access.Form
    Dim workspaceHost As Control

    Set hostForm = ResolveWorkspaceHostForm(shellForm)

    If hostForm Is Nothing Then
        ClearWorkspace = True
        Exit Function
    End If

    Set workspaceHost = GetWorkspaceHostControl(hostForm)
    If workspaceHost Is Nothing Then
        ClearWorkspace = True
        Exit Function
    End If

    workspaceHost.SourceObject = vbNullString
    ClearWorkspace = True

    modLoggingHandler.LogInfo MODULE_NAME & ".ClearWorkspace", _
        "Workspace content cleared."
    Exit Function

ErrorHandler:
    ClearWorkspace = False
    modErrorHandler.HandleError MODULE_NAME, "ClearWorkspace", Err
End Function

Private Function ResolveWorkspaceHostForm(ByVal shellForm As Access.Form) As Access.Form
    On Error GoTo SafeExit

    If shellForm Is Nothing Then
        Exit Function
    End If

    If HasControl(shellForm, WORKSPACE_SUBFORM_CONTROL) Then
        Set ResolveWorkspaceHostForm = shellForm
        Exit Function
    End If

    Set ResolveWorkspaceHostForm = shellForm.Parent

    If Not ResolveWorkspaceHostForm Is Nothing Then
        If Not HasControl(ResolveWorkspaceHostForm, WORKSPACE_SUBFORM_CONTROL) Then
            Set ResolveWorkspaceHostForm = Nothing
        End If
    End If

SafeExit:
End Function

Private Function GetWorkspaceHostControl(ByVal shellForm As Access.Form) As Control
    On Error GoTo SafeExit

    If shellForm Is Nothing Then
        Exit Function
    End If

    If HasControl(shellForm, WORKSPACE_SUBFORM_CONTROL) Then
        Set GetWorkspaceHostControl = shellForm.Controls(WORKSPACE_SUBFORM_CONTROL)
    End If

SafeExit:
End Function

Private Function FormExists(ByVal form_name As String) As Boolean
    On Error GoTo SafeExit

    Dim accessObject As Access.accessObject

    For Each accessObject In CurrentProject.AllForms
        If StrComp(accessObject.Name, form_name, vbTextCompare) = 0 Then
            FormExists = True
            Exit Function
        End If
    Next accessObject

SafeExit:
End Function

Private Function ReportExists(ByVal report_name As String) As Boolean
    On Error GoTo SafeExit

    Dim accessObject As Access.accessObject

    For Each accessObject In CurrentProject.AllReports
        If StrComp(accessObject.Name, report_name, vbTextCompare) = 0 Then
            ReportExists = True
            Exit Function
        End If
    Next accessObject

SafeExit:
End Function

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