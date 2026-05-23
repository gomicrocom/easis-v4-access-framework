Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmAppDashboard
' Purpose   : Lightweight dashboard for the application shell workspace.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmAppDashboard"

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    modFormRuntime.InitializeForm Me
    modAppDashboardService.RefreshDashboard Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox "Fehler beim Laden des Dashboards: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    modAppDashboardService.RefreshDashboard Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Public Sub RefreshDashboardView()
    On Error GoTo ErrorHandler

    modAppDashboardService.RefreshDashboard Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RefreshDashboardView", Err
End Sub
