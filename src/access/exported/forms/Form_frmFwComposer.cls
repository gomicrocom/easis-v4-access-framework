Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmFwComposer
' Purpose   : Minimal translation explorer and maintenance entry point for
'             future framework composer work.
' Author    : Codex
' Version   : 0.1.5
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmFwComposer"
Private Const PLACEHOLDER_TRANSLATION_VALUE As String = "<neu>"

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    PopulateObjectTypeCombo
    Me!cboObjectType = modFwComposerService.OBJECT_TYPE_REPORT
    LoadObjectList
    ClearControlSelection
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox "Fehler beim Initialisieren des Framework Composer: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cboObjectType_AfterUpdate()
    On Error GoTo ErrorHandler

    LoadObjectList
    ClearControlSelection
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboObjectType_AfterUpdate", Err
    MsgBox "Fehler beim Laden der Objektliste: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub lstObjects_AfterUpdate()
    On Error GoTo ErrorHandler

    LoadControlList
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "lstObjects_AfterUpdate", Err
    MsgBox "Fehler beim Laden der Steuerelemente: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub lstControls_AfterUpdate()
    On Error GoTo ErrorHandler

    RefreshTranslationPreview
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "lstControls_AfterUpdate", Err
    MsgBox "Fehler beim Laden der Uebersetzungsansicht: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdShowMissingTranslations_Click()
    On Error GoTo ErrorHandler

    ApplyMissingTranslationFilter
    modLoggingHandler.LogInfo MODULE_NAME & ".cmdShowMissingTranslations_Click", _
        "Applied missing translations filter in composer."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdShowMissingTranslations_Click", Err
    MsgBox "Fehler beim Filtern fehlender Uebersetzungen: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdValidateTranslations_Click()
    On Error GoTo ErrorHandler

    If ValidateCurrentTranslations(True) Then
        MsgBox "Alle Uebersetzungen sind vollstaendig.", vbInformation, MODULE_NAME
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdValidateTranslations_Click", Err
    MsgBox "Fehler beim Pruefen der Uebersetzungen: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdApplyTranslationTags_Click()
    On Error GoTo ErrorHandler

    Dim objectType As String
    Dim objectName As String
    Dim AppliedCount As Long
    Dim SkippedCount As Long

    objectType = Nz(Me!cboObjectType.Value, vbNullString)
    objectName = Nz(Me!lstObjects.Value, vbNullString)

    If LenB(Trim$(objectType)) = 0 Or LenB(Trim$(objectName)) = 0 Then
        MsgBox "Bitte zuerst ein Formular oder einen Report auswaehlen.", vbExclamation, MODULE_NAME
        Exit Sub
    End If

    If Not ValidateCurrentTranslations(True) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".cmdApplyTranslationTags_Click", _
            "Translation validation failed for " & objectType & "." & objectName & "."
        Exit Sub
    End If

    modFwComposerService.ApplyTranslationTagsToObject objectType, objectName, AppliedCount, SkippedCount

    modLoggingHandler.LogInfo MODULE_NAME & ".cmdApplyTranslationTags_Click", _
        "Applied TR captions for " & objectType & "." & objectName & _
        " | applied=" & CStr(AppliedCount) & _
        " | skipped=" & CStr(SkippedCount) & "."

    MsgBox "TR-Tags wurden erfolgreich angewendet.", vbInformation, MODULE_NAME
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdApplyTranslationTags_Click", Err
    MsgBox "Fehler beim Anwenden der TR-Tags: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub PopulateObjectTypeCombo()
    Me!cboObjectType.RowSourceType = "Value List"
    Me!cboObjectType.RowSource = _
        modFwComposerService.OBJECT_TYPE_REPORT & ";" & _
        modFwComposerService.OBJECT_TYPE_FORM
End Sub

Private Sub LoadObjectList()
    On Error GoTo ErrorHandler

    Dim objectList As Collection

    Set objectList = modFwComposerService.GetComposerObjectList( _
        Nz(Me!cboObjectType, modFwComposerService.OBJECT_TYPE_REPORT))

    LoadCollectionToListBox Me!lstObjects, objectList
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadObjectList", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub LoadControlList()
    On Error GoTo ErrorHandler

    Dim controlList As Collection
    Dim objectType As String
    Dim objectName As String

    objectType = Nz(Me!cboObjectType.Value, vbNullString)
    objectName = Nz(Me!lstObjects.Value, vbNullString)

    Me!lstControls.RowSourceType = "Value List"
    Me!lstControls.RowSource = vbNullString
    Me!txtSuggestedTranslationKey = Null
    Me!txtCaptionPreview = Null
    ApplyTranslationFilter vbNullString

    If LenB(Trim$(objectType)) = 0 Or LenB(Trim$(objectName)) = 0 Then
        Exit Sub
    End If

    Set controlList = modFwComposerService.GetComposerControlList(objectType, objectName, "lbl")

    If controlList Is Nothing Or controlList.count = 0 Then
        Me!txtSuggestedTranslationKey = "(keine lbl*-Controls gefunden)"
        Exit Sub
    End If

    LoadCollectionToListBox Me!lstControls, controlList
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadControlList", Err
    MsgBox "Fehler beim Laden der Steuerelemente: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub RefreshTranslationPreview()
    On Error GoTo ErrorHandler

    Dim objectType As String
    Dim objectName As String
    Dim ControlName As String
    Dim translationKey As String

    objectType = Nz(Me!cboObjectType, vbNullString)
    objectName = Nz(Me!lstObjects, vbNullString)
    ControlName = Nz(Me!lstControls, vbNullString)

    If LenB(Trim$(objectType)) = 0 _
       Or LenB(Trim$(objectName)) = 0 _
       Or LenB(Trim$(ControlName)) = 0 Then

        Me!txtSuggestedTranslationKey = Null
        Me!txtCaptionPreview = Null
        ApplyTranslationFilter vbNullString
        Exit Sub
    End If

    translationKey = modFwComposerService.SuggestTranslationKey( _
        objectType, _
        objectName, _
        ControlName)

    Me!txtSuggestedTranslationKey = translationKey
    Me!txtCaptionPreview = modFwComposerService.GetControlCaptionValue( _
        objectType, _
        objectName, _
        ControlName)

    modFwComposerService.EnsureTranslationPlaceholders translationKey
    ApplyTranslationFilter translationKey
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RefreshTranslationPreview", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub ClearControlSelection()
    On Error GoTo ErrorHandler

    Me!lstControls.RowSourceType = "Value List"
    Me!lstControls.RowSource = vbNullString
    Me!txtSuggestedTranslationKey = Null
    Me!txtCaptionPreview = Null
    ApplyTranslationFilter vbNullString
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ClearControlSelection", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub ApplyTranslationFilter(Optional ByVal translationKey As String = vbNullString)
    Dim filterExpression As String

    If LenB(Trim$(translationKey)) > 0 Then
        filterExpression = "[translation_key] = " & SqlText(translationKey)
    Else
        filterExpression = vbNullString
    End If

    ApplySubformFilter filterExpression
End Sub

Private Sub ApplyMissingTranslationFilter()
    Dim filterExpression As String
    Dim scopeFilter As String

    filterExpression = BuildMissingTranslationFilter()
    scopeFilter = BuildCurrentContextFilter()

    If LenB(scopeFilter) > 0 Then
        filterExpression = "(" & scopeFilter & ") AND (" & filterExpression & ")"
    End If

    ApplySubformFilter filterExpression
End Sub

Private Sub ApplySubformFilter(Optional ByVal filterExpression As String = vbNullString)
    On Error GoTo ErrorHandler

    With Me!sfrmTranslations.Form
        .AllowAdditions = True

        If LenB(Trim$(filterExpression)) = 0 Then
            .FilterOn = False
            .Filter = vbNullString
        Else
            .Filter = filterExpression
            .FilterOn = True
        End If

        .Requery
    End With
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplySubformFilter", Err
    MsgBox "Fehler beim Filtern der Uebersetzungen: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Function BuildCurrentContextFilter() As String
    Dim translationKey As String
    Dim objectType As String
    Dim objectName As String

    translationKey = Trim$(Nz(Me!txtSuggestedTranslationKey.Value, vbNullString))
    If LenB(translationKey) > 0 Then
        If Left$(translationKey, 1) <> "(" Then
            BuildCurrentContextFilter = "[translation_key] = " & SqlText(translationKey)
            Exit Function
        End If
    End If

    objectType = UCase$(Trim$(Nz(Me!cboObjectType.Value, vbNullString)))
    objectName = UCase$(Trim$(Nz(Me!lstObjects.Value, vbNullString)))

    Select Case objectType
        Case modFwComposerService.OBJECT_TYPE_FORM
            If LenB(objectName) > 0 Then
                BuildCurrentContextFilter = "[translation_key] Like " & SqlText("FORM." & objectName & ".*")
            Else
                BuildCurrentContextFilter = "[translation_key] Like " & SqlText("FORM.*")
            End If

        Case modFwComposerService.OBJECT_TYPE_REPORT
            BuildCurrentContextFilter = "[translation_key] Like " & SqlText("REPORT.*")

        Case Else
            BuildCurrentContextFilter = vbNullString
    End Select
End Function

Private Function BuildMissingTranslationFilter() As String
    BuildMissingTranslationFilter = _
        "([translation_value] = " & SqlText(PLACEHOLDER_TRANSLATION_VALUE) & _
        " OR [translation_value] Is Null" & _
        " OR Trim(Nz([translation_value], '')) = '')"
End Function

Private Function ValidateCurrentTranslations(Optional ByVal ShowMessage As Boolean = True) As Boolean
    Dim objectType As String
    Dim objectName As String
    Dim isValid As Boolean

    objectType = vbNullString
    objectName = vbNullString

    If LenB(Trim$(Nz(Me!lstObjects.Value, vbNullString))) > 0 Then
        objectType = Nz(Me!cboObjectType.Value, vbNullString)
        objectName = Nz(Me!lstObjects.Value, vbNullString)
    End If

    isValid = modFwComposerService.ValidateTranslationsReady(objectType, objectName)

    If Not isValid Then
        If ShowMessage Then
            MsgBox "Es gibt noch fehlende Uebersetzungen. Bitte '<neu>' ersetzen.", vbExclamation, MODULE_NAME
        End If
        ApplyMissingTranslationFilter
    End If

    ValidateCurrentTranslations = isValid
End Function

Private Sub LoadCollectionToListBox(ByVal lst As Access.ListBox, ByVal values As Collection)
    On Error GoTo ErrorHandler

    Dim itemValue As Variant

    lst.RowSourceType = "Value List"
    lst.RowSource = vbNullString

    If values Is Nothing Then Exit Sub

    For Each itemValue In values
        lst.AddItem CStr(itemValue)
    Next itemValue

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadCollectionToListBox", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Function SqlText(ByVal Value As String) As String
    SqlText = "'" & Replace(Nz(Value, vbNullString), "'", "''") & "'"
End Function



