VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFwComposer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmFwComposer
' Purpose   : Minimal translation explorer and maintenance entry point for
'             future framework composer work.
' Author    : Codex
' Version   : 0.1.4
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

    Dim ObjectType As String
    Dim ObjectName As String
    Dim isValid As Boolean

    ObjectType = vbNullString
    ObjectName = vbNullString

    If LenB(Trim$(Nz(Me!lstObjects.Value, vbNullString))) > 0 Then
        ObjectType = Nz(Me!cboObjectType.Value, vbNullString)
        ObjectName = Nz(Me!lstObjects.Value, vbNullString)
    End If

    isValid = modFwComposerService.ValidateTranslationsReady(ObjectType, ObjectName)

    If isValid Then
        MsgBox "Alle Uebersetzungen sind vollstaendig.", vbInformation, MODULE_NAME
    Else
        MsgBox "Es gibt noch fehlende Uebersetzungen. Bitte '<neu>' ersetzen.", vbExclamation, MODULE_NAME
        ApplyMissingTranslationFilter
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdValidateTranslations_Click", Err
    MsgBox "Fehler beim Pruefen der Uebersetzungen: " & Err.description, vbExclamation, MODULE_NAME
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
    Dim ObjectType As String
    Dim ObjectName As String

    ObjectType = Nz(Me!cboObjectType.Value, vbNullString)
    ObjectName = Nz(Me!lstObjects.Value, vbNullString)

    Me!lstControls.RowSourceType = "Value List"
    Me!lstControls.RowSource = vbNullString
    Me!txtSuggestedTranslationKey = Null
    Me!txtCaptionPreview = Null
    ApplyTranslationFilter vbNullString

    If LenB(Trim$(ObjectType)) = 0 Or LenB(Trim$(ObjectName)) = 0 Then
        Exit Sub
    End If

    Set controlList = modFwComposerService.GetComposerControlList(ObjectType, ObjectName, "lbl")

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

    Dim ObjectType As String
    Dim ObjectName As String
    Dim ControlName As String
    Dim TranslationKey As String

    ObjectType = Nz(Me!cboObjectType, vbNullString)
    ObjectName = Nz(Me!lstObjects, vbNullString)
    ControlName = Nz(Me!lstControls, vbNullString)

    If LenB(Trim$(ObjectType)) = 0 _
       Or LenB(Trim$(ObjectName)) = 0 _
       Or LenB(Trim$(ControlName)) = 0 Then

        Me!txtSuggestedTranslationKey = Null
        Me!txtCaptionPreview = Null
        ApplyTranslationFilter vbNullString
        Exit Sub
    End If

    TranslationKey = modFwComposerService.SuggestTranslationKey( _
        ObjectType, _
        ObjectName, _
        ControlName)

    Me!txtSuggestedTranslationKey = TranslationKey
    Me!txtCaptionPreview = modFwComposerService.GetControlCaptionValue( _
        ObjectType, _
        ObjectName, _
        ControlName)

    modFwComposerService.EnsureTranslationPlaceholders TranslationKey
    ApplyTranslationFilter TranslationKey
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

Private Sub ApplyTranslationFilter(Optional ByVal TranslationKey As String = vbNullString)
    Dim filterExpression As String

    If LenB(Trim$(TranslationKey)) > 0 Then
        filterExpression = "[TranslationKey] = " & sqlText(TranslationKey)
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
    Dim TranslationKey As String
    Dim ObjectType As String
    Dim ObjectName As String

    TranslationKey = Trim$(Nz(Me!txtSuggestedTranslationKey.Value, vbNullString))
    If LenB(TranslationKey) > 0 Then
        If Left$(TranslationKey, 1) <> "(" Then
            BuildCurrentContextFilter = "[TranslationKey] = " & sqlText(TranslationKey)
            Exit Function
        End If
    End If

    ObjectType = UCase$(Trim$(Nz(Me!cboObjectType.Value, vbNullString)))
    ObjectName = UCase$(Trim$(Nz(Me!lstObjects.Value, vbNullString)))

    Select Case ObjectType
        Case modFwComposerService.OBJECT_TYPE_FORM
            If LenB(ObjectName) > 0 Then
                BuildCurrentContextFilter = "[TranslationKey] Like " & sqlText("FORM." & ObjectName & ".*")
            Else
                BuildCurrentContextFilter = "[TranslationKey] Like " & sqlText("FORM.*")
            End If

        Case modFwComposerService.OBJECT_TYPE_REPORT
            BuildCurrentContextFilter = "[TranslationKey] Like " & sqlText("REPORT.*")

        Case Else
            BuildCurrentContextFilter = vbNullString
    End Select
End Function

Private Function BuildMissingTranslationFilter() As String
    BuildMissingTranslationFilter = _
        "([TranslationValue] = " & sqlText(PLACEHOLDER_TRANSLATION_VALUE) & _
        " OR [TranslationValue] Is Null" & _
        " OR Trim(Nz([TranslationValue], '')) = '')"
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

Private Function sqlText(ByVal Value As String) As String
    sqlText = "'" & Replace(Nz(Value, vbNullString), "'", "''") & "'"
End Function


