VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmTmpTagComposer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()
    On Error Resume Next

    If Me.NewRecord Then Exit Sub
    If Me.Parent Is Nothing Then Exit Sub
    If Nz(Me!ControlName, "") = Nz(Me.Parent.CurrentComposerControlName, "") Then Exit Sub

    Me.Parent.LoadSelectedControlFromSubform Nz(Me!ControlName, vbNullString), Nz(Me!TagValue, vbNullString)
End Sub
