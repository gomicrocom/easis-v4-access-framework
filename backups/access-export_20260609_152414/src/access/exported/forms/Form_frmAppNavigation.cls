VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "
Form_frmAppNavigation
"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmAppNavigation
' Purpose   : Hierarchical navigation menu form for the application shell.
' Author    : Codex
' Version   : 0.3.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmAppNavigation"
Private Const DUPLICATE_CLICK_WINDOW_SECONDS As Double = 0.5

Private m_lastHandledNavigationId As Long
Private m_lastHandledAt As Date

Private Sub cmdNavClick_Click()
    On Error GoTo ErrorHandler

    HandleCurrentNavigationRow
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdNavClick_Click", Err
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    modFormRuntime.InitializeForm Me
    modFwTranslationRuntime.ApplyTranslations Me
    RefreshNavigationMenu

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox "Fehler beim Laden der Navigation: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub Detail_Click()
    On Error GoTo ErrorHandler

    HandleCurrentNavigationRow
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Detail_Click", Err
End Sub

Private Sub Detail_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler

    HandleNavigationInteraction
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Detail_DblClick", Err
End Sub

Private Sub cmdOpen_Click()
    On Error GoTo ErrorHandler

    HandleNavigationInteraction
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdOpen_Click", Err
End Sub

Private Sub txtDisplayCaption_Click()
    On Error GoTo ErrorHandler

    HandleCurrentNavigationRow
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtDisplayCaption_Click", Err
End Sub

Private Sub txtDisplayCaption_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler

    HandleNavigationInteraction
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtDisplayCaption_DblClick", Err
End Sub

Private Sub RefreshNavigationMenu()
    On Error GoTo ErrorHandler

    Me.RecordSource = modAppNavigationService.GetVisibleNavigationRowSource(ResolveCurrentRoleCode())
    Me.Requery

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RefreshNavigationMenu", Err
End Sub

Private Sub OpenSelectedNavigation()
    On Error GoTo ErrorHandler

    Dim shellForm As Access.Form
    Dim navigationId As Long
    Dim objectType As String

    navigationId = GetCurrentNavigationId()
    objectType = GetCurrentObjectType()
    Set shellForm = ResolveShellForm()

    Select Case UCase$(Trim$(objectType))
        Case "GROUP"
            If modAppNavigationService.ToggleNavigationGroup(navigationId) Then
                RefreshNavigationMenu
            End If

        Case "FORM", "REPORT", "ACTION"
            modAppNavigationService.HandleNavigationClick shellForm, navigationId

        Case Else
            modLoggingHandler.LogWarning MODULE_NAME & ".OpenSelectedNavigation", _
                "Unsupported object_type '" & objectType & "' for navigation_id=" & CStr(navigationId) & "."
    End Select
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "OpenSelectedNavigation", Err
End Sub

Private Sub HandleNavigationInteraction()
    On Error GoTo ErrorHandler

    Dim navigationId As Long

    navigationId = GetCurrentNavigationId()
    If navigationId <= 0 Then
        Exit Sub
    End If

    If IsDuplicateNavigationInteraction(navigationId) Then
        Exit Sub
    End If

    OpenSelectedNavigation
    RememberNavigationInteraction navigationId
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "HandleNavigationInteraction", Err
End Sub

Private Function GetCurrentNavigationId() As Long
    On Error GoTo SafeExit

    If HasControl("txtNavigationId") Then
        GetCurrentNavigationId = modDaoHelper.NzLong(Me!txtNavigationId.Value, 0)
    ElseIf HasControl("navigation_id") Then
        GetCurrentNavigationId = modDaoHelper.NzLong(Me!navigation_id.Value, 0)
    End If

SafeExit:
End Function

Private Function ResolveShellForm() As Access.Form
    On Error GoTo SafeExit

    Set ResolveShellForm = Me.Parent

SafeExit:
End Function

Private Function GetCurrentObjectType() As String
    On Error GoTo SafeExit

    If HasControl("txtObjectType") Then
        GetCurrentObjectType = Trim$(modDaoHelper.NzString(Me!txtObjectType.Value))
    ElseIf HasControl("object_type") Then
        GetCurrentObjectType = Trim$(modDaoHelper.NzString(Me!object_type.Value))
    End If

SafeExit:
End Function

Private Function ResolveCurrentRoleCode() As String
    If modSessionContext.IsSessionInitialized Then
        ResolveCurrentRoleCode = modSessionContext.CurrentRoleCode
    End If
End Function

Private Function IsDuplicateNavigationInteraction(ByVal navigationId As Long) As Boolean
    Dim elapsedDays As Double

    If m_lastHandledNavigationId <> navigationId Then
        Exit Function
    End If

    If m_lastHandledAt = 0 Then
        Exit Function
    End If

    elapsedDays = CDbl(Now()) - CDbl(m_lastHandledAt)
    If elapsedDays < 0 Then
        Exit Function
    End If

    IsDuplicateNavigationInteraction = _
        elapsedDays <= (DUPLICATE_CLICK_WINDOW_SECONDS / 86400#)
End Function

Private Sub RememberNavigationInteraction(ByVal navigationId As Long)
    m_lastHandledNavigationId = navigationId
    m_lastHandledAt = Now()
End Sub

Private Function HasControl(ByVal ControlName As String) As Boolean
    On Error GoTo SafeExit

    Dim ctl As Control

    For Each ctl In Me.Controls
        If StrComp(ctl.Name, ControlName, vbTextCompare) = 0 Then
            HasControl = True
            Exit Function
        End If
    Next ctl

SafeExit:
End Function


Private Sub HandleCurrentNavigationRow()
    On Error GoTo ErrorHandler

    Dim navigationId As Long
    Dim objectType As String

    navigationId = CLng(Nz(Me!txtNavigationId.Value, 0))
    objectType = UCase$(Trim$(Nz(Me!txtObjectType.Value, vbNullString)))

    If navigationId <= 0 Then
        Exit Sub
    End If

    If objectType = "GROUP" Then
        If modAppNavigationService.ToggleNavigationGroup(navigationId) Then
            Me.Requery
        End If
    Else
        Call modAppNavigationService.HandleNavigationClick(Me.Parent, navigationId)
    End If

ExitHere:
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError "Form_frmAppNavigation", "HandleCurrentNavigationRow", Err
    Resume ExitHere
End Sub

