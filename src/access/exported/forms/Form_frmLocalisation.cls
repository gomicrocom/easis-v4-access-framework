VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLocalisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
