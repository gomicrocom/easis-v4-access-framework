VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAppShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmAppShell
' Purpose   : Main application shell host form.
' Author    : Codex
' Version   : 0.3.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmAppShell"

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    Debug.Print "frmAppShell Form_Load started"
    modFwTranslationRuntime.ApplyTranslations Me

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

Private Sub txtShellSearch_AfterUpdate()
    On Error GoTo ErrorHandler

    modAppShell.ExecuteShellListSearch Me, modDaoHelper.NzString(Me!txtShellSearch.Value)
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtShellSearch_AfterUpdate", Err
End Sub

Private Sub txtShellSearch_Change()
    On Error GoTo ErrorHandler

    modAppShell.ExecuteShellListSearch Me, Nz(Me!txtShellSearch.text, vbNullString)
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtShellSearch_Change", Err
End Sub

Private Sub txtShellSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler

    Select Case KeyCode
        Case vbKeyReturn
            modAppShell.ExecuteShellListSearch Me, modDaoHelper.NzString(Me!txtShellSearch.Value)
            KeyCode = 0

        Case vbKeyEscape
            modAppShell.ExecuteShellListClearSearch Me
            KeyCode = 0
    End Select
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtShellSearch_KeyDown", Err
End Sub

Private Sub cmdShellClearSearch_Click()
    On Error GoTo ErrorHandler

    modAppShell.ExecuteShellListClearSearch Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdShellClearSearch_Click", Err
End Sub

Private Sub cmdShellNew_Click()
    On Error GoTo ErrorHandler

    modAppShell.ExecuteShellListNew Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdShellNew_Click", Err
End Sub

Private Sub cmdShellEdit_Click()
    On Error GoTo ErrorHandler

    modAppShell.ExecuteShellListEdit Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdShellEdit_Click", Err
End Sub

Private Sub cmdShellRefresh_Click()
    On Error GoTo ErrorHandler

    modAppShell.ExecuteShellListRefresh Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdShellRefresh_Click", Err
End Sub

Private Sub cmdHome_Click()
    On Error GoTo ErrorHandler

    modAppShell.LoadDefaultWorkspace Me
    modAppShell.RefreshShellStatus Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdHome_Click", Err
End Sub

Private Sub cmdBack_Click()
    On Error GoTo ErrorHandler

    If modAppWorkspaceService.GoBack(Me) Then
        modAppShell.RefreshShellStatus Me
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdBack_Click", Err
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandler

    modAppWorkspaceService.ClearWorkspace Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Unload", Err
End Sub




