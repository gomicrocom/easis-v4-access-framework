VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFwTranslationList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmFwTranslationList
' Purpose   : Editable translation maintenance form for a selected TranslationKey.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmFwTranslationList"
Private Const TABLE_FW_TRANSLATIONS As String = "fw_translation"
Private Const FIELD_TRANSLATION_KEY As String = "translation_key"
Private Const FIELD_LANGUAGE_CODE As String = "language_code"
Private Const FIELD_TRANSLATION_VALUE As String = "translation_value"

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    Me.AllowAdditions = True
    Me.AllowEdits = True
    Me.AllowDeletions = True

    Me!cboLanguageCode.RowSourceType = "Value List"
    Me!cboLanguageCode.RowSource = "DE-CH;DE-DE;EN-US;FR-FR;IT-CH"

    EnsureCurrentTranslationKeyDefault
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox "Fehler beim Initialisieren der ?bersetzungsliste: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    EnsureCurrentTranslationKeyDefault
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo ErrorHandler

    EnsureCurrentTranslationKeyDefault
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_BeforeInsert", Err
    Cancel = True
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo ErrorHandler

    Dim TranslationKey As String
    Dim LanguageCode As String
    Dim TranslationValue As String
    Dim duplicateCount As Long

    TranslationKey = Trim$(Nz(Me!txtTranslationKey.Value, vbNullString))
    LanguageCode = Trim$(Nz(Me!cboLanguageCode.Value, vbNullString))
    TranslationValue = Trim$(Nz(Me!txtTranslationValue.Value, vbNullString))

    If LenB(TranslationKey) = 0 Then
        MsgBox "TranslationKey fehlt.", vbExclamation, MODULE_NAME
        Cancel = True
        Exit Sub
    End If

    If LenB(LanguageCode) = 0 Then
        MsgBox "LanguageCode fehlt.", vbExclamation, MODULE_NAME
        Cancel = True
        Exit Sub
    End If

    If LenB(TranslationValue) = 0 Then
        MsgBox "TranslationValue fehlt.", vbExclamation, MODULE_NAME
        Cancel = True
        Exit Sub
    End If

    duplicateCount = DCount( _
        "*", _
        TABLE_FW_TRANSLATIONS, _
        "[" & FIELD_TRANSLATION_KEY & "] = " & SqlText(TranslationKey) & _
        " AND [" & FIELD_LANGUAGE_CODE & "] = " & SqlText(LanguageCode))

    If Me.NewRecord Then
        If duplicateCount > 0 Then
            MsgBox "?bersetzung f?r diese Sprache existiert bereits.", vbExclamation, MODULE_NAME
            Cancel = True
            Exit Sub
        End If
    ElseIf duplicateCount > 1 Then
        MsgBox "?bersetzung f?r diese Sprache existiert bereits.", vbExclamation, MODULE_NAME
        Cancel = True
        Exit Sub
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_BeforeUpdate", Err
    MsgBox "Fehler beim Pr?fen der ?bersetzung: " & Err.description, vbExclamation, MODULE_NAME
    Cancel = True
End Sub

Private Function CurrentTranslationKeyFromParent() As String
    On Error GoTo SafeExit

    CurrentTranslationKeyFromParent = Trim$(Nz(Forms!frmFwComposer!txtSuggestedTranslationKey.Value, vbNullString))
    Exit Function

SafeExit:
    CurrentTranslationKeyFromParent = vbNullString
End Function

Private Sub EnsureCurrentTranslationKeyDefault()
    Dim TranslationKey As String

    If Me.NewRecord Then
        TranslationKey = CurrentTranslationKeyFromParent()
        If LenB(TranslationKey) > 0 Then
            Me!txtTranslationKey = TranslationKey
        End If
    End If
End Sub

Private Function SqlText(ByVal Value As String) As String
    SqlText = "'" & Replace(Nz(Value, vbNullString), "'", "''") & "'"
End Function


