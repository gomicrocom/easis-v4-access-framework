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
' Purpose   : Minimal read-only translation explorer for future framework
'             composer work.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmFwComposer"

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
    Me!txtSuggestedTranslationKey = Null
    Me!txtCaptionPreview = Null
    ApplyTranslationFilter vbNullString
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
    MsgBox "Fehler beim Laden der Übersetzungsansicht: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub PopulateObjectTypeCombo()
    Me!cboObjectType.RowSourceType = "Value List"
    Me!cboObjectType.RowSource = modFwComposerService.OBJECT_TYPE_REPORT & ";" & modFwComposerService.OBJECT_TYPE_FORM
End Sub

Private Sub RefreshTranslationPreview()
    On Error GoTo ErrorHandler

    Dim objectType As String
    Dim objectName As String
    Dim ControlName As String
    Dim TranslationKey As String

    objectType = Nz(Me!cboObjectType, vbNullString)
    objectName = Nz(Me!lstObjects, vbNullString)
    ControlName = Nz(Me!lstControls, vbNullString)

    If LenB(Trim$(objectType)) = 0 Or LenB(Trim$(objectName)) = 0 Or LenB(Trim$(ControlName)) = 0 Then
        Me!txtSuggestedTranslationKey = Null
        Me!txtCaptionPreview = Null
        ApplyTranslationFilter vbNullString
        Exit Sub
    End If

    TranslationKey = modFwComposerService.SuggestTranslationKey(objectType, objectName, ControlName)

    Me!txtSuggestedTranslationKey = TranslationKey
    Me!txtCaptionPreview = modFwComposerService.GetControlCaptionValue(objectType, objectName, ControlName)
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

Private Sub LoadObjectList()
    On Error GoTo ErrorHandler

    Dim objectList As Collection

    Set objectList = modFwComposerService.GetComposerObjectList(Nz(Me!cboObjectType, modFwComposerService.OBJECT_TYPE_REPORT))
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

    objectType = Nz(Me!cboObjectType, vbNullString)
    objectName = Nz(Me!lstObjects, vbNullString)

    If LenB(Trim$(objectName)) = 0 Then
        Me!lstControls.RowSourceType = "Value List"
        Me!lstControls.RowSource = vbNullString
        Exit Sub
    End If

    Set controlList = modFwComposerService.GetComposerControlList(objectType, objectName, "lbl")
    LoadCollectionToListBox Me!lstControls, controlList
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadControlList", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub ApplyTranslationFilter(Optional ByVal TranslationKey As String = vbNullString)
    On Error GoTo ErrorHandler

    If LenB(Trim$(TranslationKey)) = 0 Then
        Me!sfrmTranslations.Form.FilterOn = False
        Me!sfrmTranslations.Form.Filter = vbNullString
    Else
        Me!sfrmTranslations.Form.Filter = "[TranslationKey] = " & SqlText(TranslationKey)
        Me!sfrmTranslations.Form.FilterOn = True
    End If

    Me!sfrmTranslations.Form.Requery
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyTranslationFilter", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub
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

