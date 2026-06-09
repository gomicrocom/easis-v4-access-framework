Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmFwTranslations
' Purpose   : Deprecated broad translation maintenance form retained only during
'             the transition to frmFwTranslationAudit + frmFwTranslationEdit.
' Author    : Codex
' Version   : 0.3.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmFwTranslations"
Private Const PLACEHOLDER_TRANSLATION_VALUE As String = "<neu>"
Private Const SCOPE_FORM As String = "FORM"
Private Const SCOPE_NAV As String = "NAV"
Private Const SCOPE_MSG As String = "MSG"
Private Const SCOPE_STATUS As String = "STATUS"
Private Const SCOPE_REPORT As String = "REPORT"
Private Const SCOPE_DOCUMENT As String = "DOCUMENT"
Private Const SCOPE_REF As String = "REF"
Private Const SCOPE_ALL As String = "ALL"

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    modFwTranslationRuntime.ApplyTranslations Me
    PopulateTranslationScopeCombo
    PopulateObjectTypeCombo

    If HasControl("cboTranslationScope") Then
        Me!cboTranslationScope = SCOPE_FORM
    End If

    Me!cboObjectType = modFwComposerService.OBJECT_TYPE_REPORT
    RefreshTranslationScopeUi
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox "Fehler beim Initialisieren der Uebersetzungsverwaltung: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cboTranslationScope_AfterUpdate()
    On Error GoTo ErrorHandler

    RefreshTranslationScopeUi
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboTranslationScope_AfterUpdate", Err
    MsgBox "Fehler beim Wechseln des Uebersetzungsbereichs: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cboObjectType_AfterUpdate()
    On Error GoTo ErrorHandler

    If IsFormScope() Then
        LoadObjectList
        ClearControlSelection
    End If
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
        "Applied missing translations filter in translation maintenance form."
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

    If Not IsFormScope() Then
        MsgBox "Uebersetzungs-Tags koennen nur im Bereich FORM angewendet werden.", vbInformation, MODULE_NAME
        Exit Sub
    End If

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
        "Applied translation tags for " & objectType & "." & objectName & _
        " | applied=" & CStr(AppliedCount) & _
        " | skipped=" & CStr(SkippedCount) & "."

    MsgBox "Uebersetzungs-Tags wurden erfolgreich angewendet.", vbInformation, MODULE_NAME
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdApplyTranslationTags_Click", Err
    MsgBox "Fehler beim Anwenden der Uebersetzungs-Tags: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdNewKey_Click()
    On Error GoTo ErrorHandler

    CreateNewTranslationEntry
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdNewKey_Click", Err
    MsgBox "Fehler beim Erstellen des neuen Uebersetzungsschluessels: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler

    SaveCurrentTranslationRow
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdSave_Click", Err
    MsgBox "Fehler beim Speichern der Uebersetzung: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo ErrorHandler

    RefreshTranslationScopeUi
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdRefresh_Click", Err
    MsgBox "Fehler beim Aktualisieren der Uebersetzungen: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdDeactivate_Click()
    On Error GoTo ErrorHandler

    SetCurrentTranslationActiveState False
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdDeactivate_Click", Err
    MsgBox "Fehler beim Deaktivieren der Uebersetzung: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Public Function CurrentTranslationScopeValue() As String
    If HasControl("cboTranslationScope") Then
        CurrentTranslationScopeValue = NormalizeScopeValue(Nz(Me!cboTranslationScope.Value, SCOPE_FORM))
    Else
        CurrentTranslationScopeValue = SCOPE_FORM
    End If
End Function

Public Function CurrentTranslationKeyValue() As String
    CurrentTranslationKeyValue = Trim$(Nz(Me!txtSuggestedTranslationKey.Value, vbNullString))
End Function

Public Function DefaultModuleCodeForCurrentScope() As String
    Select Case CurrentTranslationScopeValue()
        Case SCOPE_NAV
            DefaultModuleCodeForCurrentScope = "NAV"
        Case SCOPE_MSG
            DefaultModuleCodeForCurrentScope = "MSG"
        Case SCOPE_STATUS
            DefaultModuleCodeForCurrentScope = "STATUS"
        Case SCOPE_REPORT
            DefaultModuleCodeForCurrentScope = "REPORT"
        Case SCOPE_DOCUMENT
            DefaultModuleCodeForCurrentScope = "DOCUMENT"
        Case SCOPE_REF
            DefaultModuleCodeForCurrentScope = "REF"
        Case SCOPE_ALL
            DefaultModuleCodeForCurrentScope = vbNullString
        Case Else
            DefaultModuleCodeForCurrentScope = "FORM"
    End Select
End Function

Public Function NextSortOrderForCurrentScope() As Long
    Dim maxSortOrder As Variant
    Dim filterExpression As String

    filterExpression = BuildScopePrefixFilter(CurrentTranslationScopeValue())
    If LenB(filterExpression) = 0 Then
        NextSortOrderForCurrentScope = 10
        Exit Function
    End If

    maxSortOrder = DMax("sort_order", "fw_translation", filterExpression)
    If IsNull(maxSortOrder) Or IsEmpty(maxSortOrder) Then
        NextSortOrderForCurrentScope = 10
    Else
        NextSortOrderForCurrentScope = CLng(maxSortOrder) + 10
    End If
End Function

Public Sub FocusTranslationKey(ByVal translationKey As String)
    On Error GoTo ErrorHandler

    Dim scopeValue As String

    translationKey = Trim$(Nz(translationKey, vbNullString))
    If LenB(translationKey) = 0 Then
        Exit Sub
    End If

    scopeValue = ResolveScopeFromTranslationKey(translationKey)

    If HasControl("cboTranslationScope") Then
        Me!cboTranslationScope = scopeValue
    End If

    RefreshTranslationScopeUi

    If HasControl("txtSuggestedTranslationKey") Then
        Me!txtSuggestedTranslationKey = translationKey
    End If

    ApplyTranslationFilter translationKey
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "FocusTranslationKey", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub PopulateTranslationScopeCombo()
    If Not HasControl("cboTranslationScope") Then
        Exit Sub
    End If

    Me!cboTranslationScope.RowSourceType = "Value List"
    Me!cboTranslationScope.RowSource = _
        SCOPE_FORM & ";" & _
        SCOPE_NAV & ";" & _
        SCOPE_MSG & ";" & _
        SCOPE_STATUS & ";" & _
        SCOPE_REPORT & ";" & _
        SCOPE_DOCUMENT & ";" & _
        SCOPE_REF & ";" & _
        SCOPE_ALL
End Sub

Private Sub PopulateObjectTypeCombo()
    Me!cboObjectType.RowSourceType = "Value List"
    Me!cboObjectType.RowSource = _
        modFwComposerService.OBJECT_TYPE_REPORT & ";" & _
        modFwComposerService.OBJECT_TYPE_FORM
End Sub

Private Sub RefreshTranslationScopeUi()
    Dim formScopeActive As Boolean

    formScopeActive = IsFormScope()

    ToggleControlEnabled "cboObjectType", formScopeActive
    ToggleControlEnabled "lstObjects", formScopeActive
    ToggleControlEnabled "lstControls", formScopeActive
    ToggleControlEnabled "cmdApplyTranslationTags", formScopeActive

    If formScopeActive Then
        LoadObjectList
        ClearControlSelection
    Else
        ClearControlSelection
        Me!txtSuggestedTranslationKey = DefaultTranslationKeySeed()
        ApplyTranslationFilter vbNullString
    End If

    SyncSubformScopeContext
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
    SetPreviewValueIfPresent "txtTagPreview", Null
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
        SetPreviewValueIfPresent "txtTagPreview", Null
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
    SetPreviewValueIfPresent "txtTagPreview", modFwComposerService.GetControlTagValue( _
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
    SetPreviewValueIfPresent "txtTagPreview", Null
    ApplyTranslationFilter vbNullString
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ClearControlSelection", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub SetPreviewValueIfPresent(ByVal ControlName As String, ByVal previewValue As Variant)
    On Error GoTo SafeExit

    If HasControl(ControlName) Then
        Me.Controls(ControlName).Value = previewValue
    End If

SafeExit:
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

Private Sub ApplyTranslationFilter(Optional ByVal translationKey As String = vbNullString)
    Dim filterExpression As String

    If LenB(Trim$(translationKey)) > 0 Then
        filterExpression = "[translation_key] = " & SqlText(translationKey)
    ElseIf Not IsFormScope() Then
        filterExpression = BuildScopePrefixFilter(CurrentTranslationScopeValue())
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

    If Not IsFormScope() Then
        BuildCurrentContextFilter = BuildScopePrefixFilter(CurrentTranslationScopeValue())
        Exit Function
    End If

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

    If Not IsFormScope() Then
        ValidateCurrentTranslations = True
        Exit Function
    End If

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

Private Function IsFormScope() As Boolean
    IsFormScope = (StrComp(CurrentTranslationScopeValue(), SCOPE_FORM, vbTextCompare) = 0)
End Function

Private Function NormalizeScopeValue(ByVal scopeValue As String) As String
    scopeValue = UCase$(Trim$(scopeValue))

    Select Case scopeValue
        Case SCOPE_NAV, SCOPE_MSG, SCOPE_STATUS, SCOPE_REPORT, SCOPE_DOCUMENT, SCOPE_REF, SCOPE_ALL
            NormalizeScopeValue = scopeValue
        Case Else
            NormalizeScopeValue = SCOPE_FORM
    End Select
End Function

Private Function ResolveScopeFromTranslationKey(ByVal translationKey As String) As String
    Dim uppercaseKey As String

    uppercaseKey = UCase$(Trim$(translationKey))

    Select Case True
        Case Left$(uppercaseKey, 5) = "FORM."
            ResolveScopeFromTranslationKey = SCOPE_FORM
        Case Left$(uppercaseKey, 4) = "NAV."
            ResolveScopeFromTranslationKey = SCOPE_NAV
        Case Left$(uppercaseKey, 4) = "MSG." Or Left$(uppercaseKey, 4) = "MSG_" Or Left$(uppercaseKey, 4) = "ERR_"
            ResolveScopeFromTranslationKey = SCOPE_MSG
        Case Left$(uppercaseKey, 7) = "STATUS."
            ResolveScopeFromTranslationKey = SCOPE_STATUS
        Case Left$(uppercaseKey, 7) = "REPORT."
            ResolveScopeFromTranslationKey = SCOPE_REPORT
        Case Left$(uppercaseKey, 9) = "DOCUMENT."
            ResolveScopeFromTranslationKey = SCOPE_DOCUMENT
        Case Left$(uppercaseKey, 13) = "ADDRESS_TYPE." Or _
             Left$(uppercaseKey, 11) = "SALUTATION." Or _
             Left$(uppercaseKey, 13) = "CONTACT_TYPE." Or _
             Left$(uppercaseKey, 5) = "UNIT." Or _
             Left$(uppercaseKey, 4) = "VAT." Or _
             Left$(uppercaseKey, 4) = "REF."
            ResolveScopeFromTranslationKey = SCOPE_REF
        Case Left$(uppercaseKey, 7) = "COMMON."
            ResolveScopeFromTranslationKey = SCOPE_ALL
        Case Else
            ResolveScopeFromTranslationKey = SCOPE_ALL
    End Select
End Function

Private Function BuildScopePrefix(ByVal scopeValue As String) As String
    Select Case NormalizeScopeValue(scopeValue)
        Case SCOPE_NAV
            BuildScopePrefix = "NAV."
        Case SCOPE_MSG
            BuildScopePrefix = "MSG."
        Case SCOPE_STATUS
            BuildScopePrefix = "STATUS."
        Case SCOPE_REPORT
            BuildScopePrefix = "REPORT."
        Case SCOPE_DOCUMENT
            BuildScopePrefix = "DOCUMENT."
        Case SCOPE_REF
            BuildScopePrefix = "REF."
        Case SCOPE_ALL
            BuildScopePrefix = vbNullString
        Case Else
            BuildScopePrefix = vbNullString
    End Select
End Function

Private Function BuildScopePrefixFilter(ByVal scopeValue As String) As String
    Select Case NormalizeScopeValue(scopeValue)
        Case SCOPE_NAV
            BuildScopePrefixFilter = "[translation_key] Like " & SqlText("NAV.*")

        Case SCOPE_MSG
            BuildScopePrefixFilter = _
                "([translation_key] Like " & SqlText("MSG.*") & _
                " OR [translation_key] Like " & SqlText("MSG_*") & ")"

        Case SCOPE_STATUS
            BuildScopePrefixFilter = "[translation_key] Like " & SqlText("STATUS.*")

        Case SCOPE_REPORT
            BuildScopePrefixFilter = "[translation_key] Like " & SqlText("REPORT.*")

        Case SCOPE_DOCUMENT
            BuildScopePrefixFilter = "[translation_key] Like " & SqlText("DOCUMENT.*")

        Case SCOPE_REF
            BuildScopePrefixFilter = _
                "([translation_key] Like " & SqlText("ADDRESS_TYPE.*") & _
                " OR [translation_key] Like " & SqlText("SALUTATION.*") & _
                " OR [translation_key] Like " & SqlText("CONTACT_TYPE.*") & _
                " OR [translation_key] Like " & SqlText("UNIT.*") & _
                " OR [translation_key] Like " & SqlText("VAT.*") & _
                " OR [translation_key] Like " & SqlText("REF.*") & ")"

        Case SCOPE_ALL
            BuildScopePrefixFilter = "Not ([translation_key] Like " & SqlText("FORM.*") & ")"

        Case Else
            BuildScopePrefixFilter = vbNullString
    End Select
End Function

Private Function DefaultTranslationKeySeed() As String
    DefaultTranslationKeySeed = BuildScopePrefix(CurrentTranslationScopeValue())
End Function

Private Sub ToggleControlEnabled(ByVal ControlName As String, ByVal isEnabled As Boolean)
    On Error GoTo SafeExit

    If HasControl(ControlName) Then
        Me.Controls(ControlName).Enabled = isEnabled
    End If

SafeExit:
End Sub

Private Sub SyncSubformScopeContext()
    On Error GoTo SafeExit

    If HasControl("sfrmTranslations") Then
        Me!sfrmTranslations.Form.Requery
    End If

SafeExit:
End Sub

Private Sub CreateNewTranslationEntry()
    On Error GoTo ErrorHandler

    If HasControl("txtSuggestedTranslationKey") Then
        If Not IsFormScope() Then
            Me!txtSuggestedTranslationKey = DefaultTranslationKeySeed()
        End If
    End If

    If HasControl("sfrmTranslations") Then
        Me!sfrmTranslations.Form.SetFocus
        CallByName Me!sfrmTranslations.Form, "PrepareNewTranslationRow", VbMethod
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CreateNewTranslationEntry", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub SaveCurrentTranslationRow()
    On Error GoTo ErrorHandler

    If HasControl("sfrmTranslations") Then
        If Me!sfrmTranslations.Form.Dirty Then
            Me!sfrmTranslations.Form.Dirty = False
        End If
        Me!sfrmTranslations.Form.Requery
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SaveCurrentTranslationRow", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub SetCurrentTranslationActiveState(ByVal isActive As Boolean)
    On Error GoTo ErrorHandler

    If Not HasControl("sfrmTranslations") Then
        Exit Sub
    End If

    If Me!sfrmTranslations.Form.NewRecord Then
        MsgBox "Bitte zuerst eine gespeicherte Uebersetzung auswaehlen.", vbExclamation, MODULE_NAME
        Exit Sub
    End If

    If HasFormControl(Me!sfrmTranslations.Form, "chkIsActive") Then
        Me!sfrmTranslations.Form!chkIsActive = isActive
    ElseIf HasFormControl(Me!sfrmTranslations.Form, "is_active") Then
        Me!sfrmTranslations.Form!is_active = isActive
    End If

    If Me!sfrmTranslations.Form.Dirty Then
        Me!sfrmTranslations.Form.Dirty = False
    End If
    Me!sfrmTranslations.Form.Requery
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SetCurrentTranslationActiveState", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Function HasFormControl(ByVal formInstance As Access.Form, ByVal ControlName As String) As Boolean
    On Error GoTo SafeExit

    Dim ctl As Control

    If formInstance Is Nothing Then
        Exit Function
    End If

    For Each ctl In formInstance.Controls
        If StrComp(ctl.Name, ControlName, vbTextCompare) = 0 Then
            HasFormControl = True
            Exit Function
        End If
    Next ctl

SafeExit:
End Function
