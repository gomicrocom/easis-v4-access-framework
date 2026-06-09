Option Compare Database
Option Explicit

Private Sub cmdOpenTagComposer_Click()

    DoCmd.OpenForm "frmTagComposer", , , , , acDialog, Me.Name

End Sub
Private Sub cmdSave_Click()
    If Not modFormRuntime.ValidateFormPolicies(Me) Then
        Exit Sub
    End If
    
    DoCmd.RunCommand acCmdSaveRecord
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError "frmCustomer", "cmdSave_Click", Err
End Sub

Private Sub Form_Load()
    modFormRuntime.InitializeForm Me
End Sub