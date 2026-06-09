Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_sfrmAddressContacts
' Purpose   : Contacts Listing
' Author    : Codex
' Version   : 0.2.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_sfrmAddressContacts"

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    ConfigureContactTypeCombo
    modFwTranslationRuntime.ApplyTranslations Me

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "Address Contacts loaded."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox "Fehler beim Laden der Adressliste: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub ConfigureContactTypeCombo()
    On Error GoTo ErrorHandler

    Dim currentLanguageCode As String

    If Not HasControl("cboContactType") Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ConfigureContactTypeCombo", _
            "Control cboContactType not found."
        Exit Sub
    End If

    currentLanguageCode = modFwTranslationRuntime.GetCurrentLanguageCode()

    With Me!cboContactType
        .ControlSource = "contact_type_code"
        .RowSourceType = "Table/Query"
        .RowSource = modFwTranslationRuntime.BuildTranslatedReferenceRowSource( _
            "ref_contact_type", _
            "contact_type_code", _
            currentLanguageCode)
        .BoundColumn = 1
        .ColumnCount = 2
        .ColumnWidths = "0cm;4cm"
        .LimitToList = True
        .Requery
        Debug.Print "cboContactType RowSource = " & Me!cboContactType.RowSource
        Debug.Print "cboContactType ListCount = " & Me!cboContactType.ListCount
    End With
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureContactTypeCombo", Err
    Err.Raise Err.Number, Err.Source, Err.description
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











