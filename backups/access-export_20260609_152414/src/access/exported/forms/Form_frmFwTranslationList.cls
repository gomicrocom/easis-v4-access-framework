Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmFwTranslationList
' Purpose   : Editable translation maintenance form for form-based and free
'             translation namespaces.
' Author    : Codex
' Version   : 0.2.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmFwTranslationList"
Private Const TABLE_FW_TRANSLATIONS As String = "fw_translation"
Private Const FIELD_TRANSLATION_KEY As String = "translation_key"
Private Const FIELD_LANGUAGE_CODE As String = "language_code"
Private Const FIELD_TRANSLATION_VALUE As String = "translation_value"
Private Const FIELD_MODULE_CODE As String = "module_code"
Private Const FIELD_IS_ACTIVE As String = "is_active"
Private Const FIELD_SORT_ORDER As String = "sort_order"
Private Const FIELD_UPDATED_AT As String = "updated_at"
Private Const FIELD_UPDATED_BY As String = "updated_by"
Private Const DEFAULT_LANGUAGE_CODE As String = "DE-CH"

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    Me.AllowAdditions = True
    Me.AllowEdits = True
    Me.AllowDeletions = True

    Me!cboLanguageCode.RowSourceType = "Value List"
    Me!cboLanguageCode.RowSource = "DE-CH;DE-DE;EN-US;FR-FR;IT-CH"

    EnsureCurrentTranslationDefaults
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox "Fehler beim Initialisieren der Uebersetzungsliste: " & Err.Description, vbExclamation, MODULE_NAME
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    EnsureCurrentTranslationDefaults
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo ErrorHandler

    EnsureCurrentTranslationDefaults
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_BeforeInsert", Err
    Cancel = True
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo ErrorHandler

    Dim translationKey As String
    Dim languageCode As String
    Dim translationValue As String
    Dim scopeValue As String
    Dim duplicateCount As Long

    translationKey = Trim$(Nz(Me!txtTranslationKey.Value, vbNullString))
    languageCode = Trim$(Nz(Me!cboLanguageCode.Value, vbNullString))
    translationValue = Trim$(Nz(Me!txtTranslationValue.Value, vbNullString))
    scopeValue = CurrentTranslationScopeValue()

    If LenB(translationKey) = 0 Then
        MsgBox "translation_key fehlt.", vbExclamation, MODULE_NAME
        Cancel = True
        Exit Sub
    End If

    If Left$(UCase$(translationKey), 3) = "TR:" Then
        MsgBox "translation_key darf keinen TR:-Praefix enthalten.", vbExclamation, MODULE_NAME
        Cancel = True
        Exit Sub
    End If

    If LenB(languageCode) = 0 Then
        MsgBox "language_code fehlt.", vbExclamation, MODULE_NAME
        Cancel = True
        Exit Sub
    End If

    If LenB(translationValue) = 0 Then
        MsgBox "translation_value fehlt.", vbExclamation, MODULE_NAME
        Cancel = True
        Exit Sub
    End If

    If Not TranslationKeyMatchesScope(scopeValue, translationKey) Then
        MsgBox "translation_key passt nicht zum gewaehlten Scope.", vbExclamation, MODULE_NAME
        Cancel = True
        Exit Sub
    End If

    duplicateCount = DCount( _
        "*", _
        TABLE_FW_TRANSLATIONS, _
        "[" & FIELD_TRANSLATION_KEY & "] = " & SqlText(translationKey) & _
        " AND [" & FIELD_LANGUAGE_CODE & "] = " & SqlText(languageCode))

    If Me.NewRecord Then
        If duplicateCount > 0 Then
            If UpdateExistingTranslationRow(translationKey, languageCode, translationValue) Then
                MsgBox "Vorhandene Uebersetzung wurde aktualisiert.", vbInformation, MODULE_NAME
                Me.Undo
                Me.Requery
                Cancel = True
                Exit Sub
            End If

            MsgBox "Uebersetzung fuer diese Sprache existiert bereits.", vbExclamation, MODULE_NAME
            Cancel = True
            Exit Sub
        End If
    ElseIf duplicateCount > 1 Then
        MsgBox "Uebersetzung fuer diese Sprache existiert bereits.", vbExclamation, MODULE_NAME
        Cancel = True
        Exit Sub
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_BeforeUpdate", Err
    MsgBox "Fehler beim Pruefen der Uebersetzung: " & Err.Description, vbExclamation, MODULE_NAME
    Cancel = True
End Sub

Public Sub PrepareNewTranslationRow()
    On Error GoTo ErrorHandler

    If Me.Dirty Then
        Me.Dirty = False
    End If

    DoCmd.GoToRecord , , acNewRec
    EnsureCurrentTranslationDefaults
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "PrepareNewTranslationRow", Err
    MsgBox "Fehler beim Anlegen der neuen Uebersetzung: " & Err.Description, vbExclamation, MODULE_NAME
End Sub

Private Function CurrentTranslationKeyFromParent() As String
    On Error GoTo SafeExit

    CurrentTranslationKeyFromParent = Trim$(Nz(Forms!frmFwTranslations.CurrentTranslationKeyValue(), vbNullString))
    Exit Function

SafeExit:
    CurrentTranslationKeyFromParent = vbNullString
End Function

Private Function CurrentTranslationScopeValue() As String
    On Error GoTo SafeExit

    CurrentTranslationScopeValue = Trim$(Nz(Forms!frmFwTranslations.CurrentTranslationScopeValue(), "FORM"))
    Exit Function

SafeExit:
    CurrentTranslationScopeValue = "FORM"
End Function

Private Function DefaultModuleCodeFromParent() As String
    On Error GoTo SafeExit

    DefaultModuleCodeFromParent = Trim$(Nz(Forms!frmFwTranslations.DefaultModuleCodeForCurrentScope(), vbNullString))
    Exit Function

SafeExit:
    DefaultModuleCodeFromParent = vbNullString
End Function

Private Function NextSortOrderFromParent() As Long
    On Error GoTo SafeExit

    NextSortOrderFromParent = CLng(Forms!frmFwTranslations.NextSortOrderForCurrentScope())
    Exit Function

SafeExit:
    NextSortOrderFromParent = 10
End Function

Private Function CurrentLanguageFromRuntime() As String
    On Error GoTo SafeExit

    CurrentLanguageFromRuntime = Trim$(Nz(modFwTranslationRuntime.GetCurrentLanguageCode(), DEFAULT_LANGUAGE_CODE))
    If LenB(CurrentLanguageFromRuntime) = 0 Then
        CurrentLanguageFromRuntime = DEFAULT_LANGUAGE_CODE
    End If
    Exit Function

SafeExit:
    CurrentLanguageFromRuntime = DEFAULT_LANGUAGE_CODE
End Function

Private Sub EnsureCurrentTranslationDefaults()
    Dim translationKey As String

    If Not Me.NewRecord Then
        Exit Sub
    End If

    translationKey = CurrentTranslationKeyFromParent()
    If LenB(translationKey) > 0 Then
        Me!txtTranslationKey = translationKey
    End If

    If HasFormControl("txtModuleCode") Then
        If LenB(Trim$(Nz(Me!txtModuleCode.Value, vbNullString))) = 0 Then
            Me!txtModuleCode = DefaultModuleCodeFromParent()
        End If
    End If

    If LenB(Trim$(Nz(Me!cboLanguageCode.Value, vbNullString))) = 0 Then
        Me!cboLanguageCode = CurrentLanguageFromRuntime()
    End If

    If HasFormControl("chkIsActive") Then
        If IsNull(Me!chkIsActive.Value) Then
            Me!chkIsActive = True
        End If
    ElseIf HasFormControl(FIELD_IS_ACTIVE) Then
        If IsNull(Me.Controls(FIELD_IS_ACTIVE).Value) Then
            Me.Controls(FIELD_IS_ACTIVE).Value = True
        End If
    End If

    If HasFormControl("txtSortOrder") Then
        If Val(Nz(Me!txtSortOrder.Value, 0)) <= 0 Then
            Me!txtSortOrder = NextSortOrderFromParent()
        End If
    ElseIf HasFormControl(FIELD_SORT_ORDER) Then
        If Val(Nz(Me.Controls(FIELD_SORT_ORDER).Value, 0)) <= 0 Then
            Me.Controls(FIELD_SORT_ORDER).Value = NextSortOrderFromParent()
        End If
    End If
End Sub

Private Function TranslationKeyMatchesScope(ByVal scopeValue As String, ByVal translationKey As String) As Boolean
    Dim normalizedScope As String
    Dim uppercaseKey As String

    normalizedScope = UCase$(Trim$(scopeValue))
    uppercaseKey = UCase$(Trim$(translationKey))

    Select Case normalizedScope
        Case "FORM"
            TranslationKeyMatchesScope = _
                (Left$(uppercaseKey, 5) = "FORM.") Or _
                (Left$(uppercaseKey, 7) = "REPORT.")
        Case "NAV"
            TranslationKeyMatchesScope = (Left$(uppercaseKey, 4) = "NAV.")
        Case "MSG"
            TranslationKeyMatchesScope = _
                (Left$(uppercaseKey, 4) = "MSG.") Or _
                (Left$(uppercaseKey, 4) = "MSG_")
        Case "STATUS"
            TranslationKeyMatchesScope = (Left$(uppercaseKey, 7) = "STATUS.")
        Case "REPORT"
            TranslationKeyMatchesScope = (Left$(uppercaseKey, 7) = "REPORT.")
        Case "DOCUMENT"
            TranslationKeyMatchesScope = (Left$(uppercaseKey, 9) = "DOCUMENT.")
        Case "REF"
            TranslationKeyMatchesScope = _
                (Left$(uppercaseKey, 13) = "ADDRESS_TYPE.") Or _
                (Left$(uppercaseKey, 11) = "SALUTATION.") Or _
                (Left$(uppercaseKey, 13) = "CONTACT_TYPE.") Or _
                (Left$(uppercaseKey, 5) = "UNIT.") Or _
                (Left$(uppercaseKey, 4) = "VAT.") Or _
                (Left$(uppercaseKey, 4) = "REF.")
        Case "ALL"
            TranslationKeyMatchesScope = (Left$(uppercaseKey, 5) <> "FORM.")
        Case Else
            TranslationKeyMatchesScope = True
    End Select
End Function

Private Function UpdateExistingTranslationRow( _
    ByVal translationKey As String, _
    ByVal languageCode As String, _
    ByVal translationValue As String) As Boolean

    On Error GoTo ErrorHandler

    Dim currentDb As DAO.Database
    Dim sqlStatement As String
    Dim moduleCode As String
    Dim isActiveValue As Boolean
    Dim sortOrderValue As Variant

    moduleCode = vbNullString
    If HasFormControl("txtModuleCode") Then
        moduleCode = Trim$(Nz(Me!txtModuleCode.Value, vbNullString))
    ElseIf HasFormControl(FIELD_MODULE_CODE) Then
        moduleCode = Trim$(Nz(Me.Controls(FIELD_MODULE_CODE).Value, vbNullString))
    End If

    isActiveValue = True
    If HasFormControl("chkIsActive") Then
        isActiveValue = Nz(Me!chkIsActive.Value, True)
    ElseIf HasFormControl(FIELD_IS_ACTIVE) Then
        isActiveValue = Nz(Me.Controls(FIELD_IS_ACTIVE).Value, True)
    End If

    sortOrderValue = Null
    If HasFormControl("txtSortOrder") Then
        sortOrderValue = Me!txtSortOrder.Value
    ElseIf HasFormControl(FIELD_SORT_ORDER) Then
        sortOrderValue = Me.Controls(FIELD_SORT_ORDER).Value
    End If

    sqlStatement = _
        "UPDATE " & TABLE_FW_TRANSLATIONS & _
        " SET [" & FIELD_TRANSLATION_VALUE & "] = " & SqlText(translationValue) & _
        ", [" & FIELD_MODULE_CODE & "] = " & SqlNullableText(moduleCode) & _
        ", [" & FIELD_IS_ACTIVE & "] = " & SqlBoolean(isActiveValue) & _
        ", [" & FIELD_SORT_ORDER & "] = " & SqlLongOrNull(sortOrderValue) & _
        ", [" & FIELD_UPDATED_AT & "] = Now()" & _
        ", [" & FIELD_UPDATED_BY & "] = 'SYSTEM'" & _
        " WHERE [" & FIELD_TRANSLATION_KEY & "] = " & SqlText(translationKey) & _
        " AND [" & FIELD_LANGUAGE_CODE & "] = " & SqlText(languageCode)

    Set currentDb = CurrentDb
    currentDb.Execute sqlStatement, dbFailOnError
    UpdateExistingTranslationRow = (currentDb.RecordsAffected > 0)
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "UpdateExistingTranslationRow", Err
    UpdateExistingTranslationRow = False
End Function

Private Function HasFormControl(ByVal controlName As String) As Boolean
    On Error GoTo MissingControl

    Dim currentControl As Control

    Set currentControl = Me.Controls(controlName)
    HasFormControl = True
    Exit Function

MissingControl:
    HasFormControl = False
End Function
