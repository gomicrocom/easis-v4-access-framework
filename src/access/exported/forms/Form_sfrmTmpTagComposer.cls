Option Compare Database
Option Explicit

Private Const FIELD_CONTROL_NAME As String = "control_name"
Private Const FIELD_TAG_VALUE As String = "tag_value"

Private Sub Form_Current()
    On Error Resume Next

    If Me.NewRecord Then Exit Sub
    If Me.Parent Is Nothing Then Exit Sub
    If Nz(Me.Controls(FIELD_CONTROL_NAME).Value, "") = Nz(Me.Parent.CurrentComposerControlName, "") Then Exit Sub

    Me.Parent.LoadSelectedControlFromSubform _
        Nz(Me.Controls(FIELD_CONTROL_NAME).Value, vbNullString), _
        Nz(Me.Controls(FIELD_TAG_VALUE).Value, vbNullString)
End Sub

