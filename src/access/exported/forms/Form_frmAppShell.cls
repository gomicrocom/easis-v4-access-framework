Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmAppShell
' Purpose   : Main application shell host form.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmAppShell"

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    Debug.Print "frmAppShell Form_Load started"

    If Not modAppShell.InitializeAppShell(Me) Then
        Debug.Print "InitializeAppShell returned False"
        modLoggingHandler.LogWarning MODULE_NAME & ".Form_Load", _
            "Application shell initialization returned False."
    Else
        Debug.Print "InitializeAppShell returned True"
    End If

    Debug.Print "Navigation SourceObject: " & Me!subNavigationHost.SourceObject
    Debug.Print "Workspace SourceObject: " & Me!subWorkspaceHost.SourceObject

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox "Fehler beim Laden der Anwendungsshell: " & Err.description, vbExclamation, MODULE_NAME
End Sub
Private Sub Form_Current()
    On Error GoTo ErrorHandler

    modAppShell.RefreshShellStatus Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub cmdRefreshShell_Click()
    On Error GoTo ErrorHandler

    modAppShell.RefreshShellStatus Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdRefreshShell_Click", Err
End Sub

Private Sub cmdHome_Click()
    On Error GoTo ErrorHandler

    modAppShell.LoadDefaultWorkspace Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdHome_Click", Err
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandler

    modAppWorkspaceService.ClearWorkspace Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Unload", Err
End Sub

