VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfrmAddressContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_sfrmAddressContacts
' Purpose   : Contacts Listing
' Author    : Codex
' Version   : 0.1.1
'===============================================================================

Private Const MODULE_NAME As String = "Form_sfrmAddressContacts"

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    modFwTranslationRuntime.ApplyTranslations Me
    
    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "Address Contacts loaded."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox "Fehler beim Laden der Adressliste: " & Err.description, vbExclamation, MODULE_NAME
End Sub





