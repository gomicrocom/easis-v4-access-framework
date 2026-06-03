VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFwTranslationTagGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmFwTranslationTagGenerator
' Purpose   : Focused workspace form for FORM translation-tag inspection,
'             staging, apply, and removal.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmFwTranslationTagGenerator"
Private Const WORK_TABLE As String = "tmp_fw_translation_tag_generator"
Private Const EXCLUDED_FORM_NAME As String = "frmFwTranslationTagGenerator"

Private m_sessionId As String
Private m_currentFormName As String

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    m_sessionId = BuildSessionId()

    modFormRuntime.InitializeForm Me
    modFwTranslationRuntime.ApplyTranslations Me

    PopulateFormCombo
    ConfigureDataSource
    UpdateActionState

    If LenB(Trim$(modDaoHelper.NzString(Me.OpenArgs))) > 0 Then
        SetSelectedFormName CStr(Me.OpenArgs)
        LoadSelectedFormControls
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "Translation tag generator loaded."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.FW_TRANSLATION_TAG_GENERATOR_LOAD_ERROR", "Fehler beim Laden des Translation-Tag-Generators."), vbExclamation, MODULE_NAME
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrorHandler

    modFwTranslationTagGeneratorService.ClearSession m_sessionId
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Unload", Err
End Sub

Private Sub cboFormName_AfterUpdate()
    On Error GoTo ErrorHandler

    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboFormName_AfterUpdate", Err
End Sub

Private Sub chkIncludeHidden_AfterUpdate()
    On Error GoTo ErrorHandler

    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "chkIncludeHidden_AfterUpdate", Err
End Sub

Private Sub chkShowAllControls_AfterUpdate()
    On Error GoTo ErrorHandler

    ApplyDisplayFilter
    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "chkShowAllControls_AfterUpdate", Err
End Sub

Private Sub cmdLoadForm_Click()
    On Error GoTo ErrorHandler

    LoadSelectedFormControls
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdLoadForm_Click", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.FW_TRANSLATION_TAG_GENERATOR_LOAD_FORM_ERROR", "Fehler beim Laden des ausgewaehlten Formulars."), vbExclamation, MODULE_NAME
End Sub

Private Sub cmdGenerateSuggestions_Click()
    On Error GoTo ErrorHandler

    EnsureLoadedForm
    modFwTranslationTagGeneratorService.RegenerateSuggestions m_sessionId
    Me.Requery
    ApplyDisplayFilter
    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdGenerateSuggestions_Click", Err
End Sub

Private Sub cmdSetMissingKeys_Click()
    On Error GoTo ErrorHandler

    EnsureLoadedForm

    If modFwTranslationTagGeneratorService.SetMissingKeys(m_sessionId) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.FW_TRANSLATION_TAG_GENERATOR_NO_MISSING_KEYS", "Keine fehlenden Translation-Keys zum Setzen gefunden."), vbInformation, MODULE_NAME
    End If

    Me.Requery
    ApplyDisplayFilter
    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdSetMissingKeys_Click", Err
End Sub

Private Sub cmdSetSelectedKey_Click()
    On Error GoTo ErrorHandler

    Dim controlName As String

    EnsureLoadedForm
    controlName = RequireCurrentControlName()
    EnsureCurrentControlCanReceiveKey

    If modFwTranslationTagGeneratorService.SetKeyForControl(m_sessionId, controlName) Then
        Me.Requery
        ApplyDisplayFilter
        RestoreCurrentControlSelection controlName
    End If

    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdSetSelectedKey_Click", Err
End Sub

Private Sub cmdRemoveSelectedKey_Click()
    On Error GoTo ErrorHandler

    Dim controlName As String

    EnsureLoadedForm
    controlName = RequireCurrentControlName()

    If modFwTranslationTagGeneratorService.RemoveKeyFromControl(m_sessionId, controlName) Then
        Me.Requery
        ApplyDisplayFilter
        RestoreCurrentControlSelection controlName
    End If

    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdRemoveSelectedKey_Click", Err
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler

    Dim updatedCount As Long

    EnsureLoadedForm

    If modFwTranslationTagGeneratorService.SaveTagChanges(m_sessionId, m_currentFormName, updatedCount) Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.FW_TRANSLATION_TAG_GENERATOR_SAVE_SUCCESS", "Translation-Tags wurden gespeichert."), vbInformation, MODULE_NAME
        Me.Requery
        ApplyDisplayFilter
    End If

    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdSave_Click", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.FW_TRANSLATION_TAG_GENERATOR_SAVE_ERROR", "Fehler beim Speichern der Translation-Tags."), vbExclamation, MODULE_NAME
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo ErrorHandler

    If Not CanDiscardChanges() Then
        Exit Sub
    End If

    modFwTranslationTagGeneratorService.ClearSession m_sessionId
    ReturnToPreviousWorkspace
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdCancel_Click", Err
End Sub

Public Function SupportsListCommandBar() As Boolean
    SupportsListCommandBar = False
End Function

Public Function CanLeaveWorkspace() As Boolean
    On Error GoTo ErrorHandler

    CanLeaveWorkspace = CanDiscardChanges()
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CanLeaveWorkspace", Err
End Function

Public Function GetCurrentWorkspaceRecordId() As Long
    GetCurrentWorkspaceRecordId = 0
End Function

Public Function GetWorkspaceState() As String
    Dim stateText As String

    stateText = AddWorkspaceStatePair(stateText, "form_name", m_currentFormName)
    stateText = AddWorkspaceStatePair(stateText, "include_hidden", IIf(IncludeHiddenControls(), "1", "0"))
    stateText = AddWorkspaceStatePair(stateText, "show_all_controls", IIf(ShowAllControls(), "1", "0"))
    stateText = AddWorkspaceStatePair(stateText, "control_name", GetCurrentControlName())

    GetWorkspaceState = stateText
End Function

Public Sub RestoreWorkspaceState(ByVal stateText As String)
    On Error GoTo ErrorHandler

    Dim formName As String
    Dim controlName As String
    Dim includeHidden As Boolean
    Dim showAllControls As Boolean

    formName = GetWorkspaceStateValue(stateText, "form_name")
    controlName = GetWorkspaceStateValue(stateText, "control_name")
    includeHidden = (StrComp(GetWorkspaceStateValue(stateText, "include_hidden"), "1", vbBinaryCompare) = 0)
    showAllControls = (StrComp(GetWorkspaceStateValue(stateText, "show_all_controls"), "1", vbBinaryCompare) = 0)

    SetIncludeHiddenControls includeHidden
    SetShowAllControls showAllControls
    SetSelectedFormName formName

    If LenB(formName) > 0 Then
        LoadSelectedFormControls
        If LenB(controlName) > 0 Then
            RestoreCurrentControlSelection controlName
        End If
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreWorkspaceState", Err
End Sub

Private Sub PopulateFormCombo()
    On Error GoTo ErrorHandler

    Dim formNames As Collection
    Dim formName As Variant
    Dim rowSource As String

    If Not HasControl("cboFormName") Then
        Exit Sub
    End If

    Set formNames = modFwComposerService.GetComposerObjectList(modFwComposerService.OBJECT_TYPE_FORM)

    For Each formName In formNames
        If StrComp(CStr(formName), EXCLUDED_FORM_NAME, vbTextCompare) = 0 Then
            GoTo NextFormName
        End If

        If LenB(rowSource) > 0 Then
            rowSource = rowSource & ";"
        End If
        rowSource = rowSource & Replace(CStr(formName), ";", ",")

NextFormName:
    Next formName

    Me!cboFormName.RowSourceType = "Value List"
    Me!cboFormName.RowSource = rowSource
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "PopulateFormCombo", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ConfigureDataSource()
    On Error GoTo ErrorHandler

    modFwTranslationTagGeneratorService.EnsureTranslationTagGeneratorTable
    Me.RecordSource = "SELECT " & _
                      "row_no, form_name, control_name, control_type, source_text, current_tag, current_translation_key, suggested_translation_key, tag_status, is_relevant, is_visible, is_dirty " & _
                      "FROM " & WORK_TABLE & " " & _
                      "WHERE session_id = " & SqlText(m_sessionId) & " " & _
                      "ORDER BY row_no;"
    Me.OrderBy = "row_no"
    Me.OrderByOn = True
    ApplyDisplayFilter
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureDataSource", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub LoadSelectedFormControls()
    On Error GoTo ErrorHandler

    Dim formName As String

    formName = Trim$(modDaoHelper.NzString(GetSelectedFormName()))
    If LenB(formName) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.FW_TRANSLATION_TAG_GENERATOR_FORM_REQUIRED", "Bitte zuerst ein Formular auswaehlen."), vbExclamation, MODULE_NAME
        Exit Sub
    End If

    If LenB(m_currentFormName) > 0 Then
        If StrComp(m_currentFormName, formName, vbTextCompare) <> 0 Then
            If Not CanDiscardChanges() Then
                Exit Sub
            End If
        End If
    End If

    m_currentFormName = formName
    modFwTranslationTagGeneratorService.LoadFormControlRows m_sessionId, formName, IncludeHiddenControls()
    Me.Requery
    ApplyDisplayFilter
    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadSelectedFormControls", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub UpdateActionState()
    On Error GoTo SafeExit

    Dim hasFormSelection As Boolean
    Dim hasLoadedForm As Boolean
    Dim hasCurrentControl As Boolean
    Dim canSetSelectedKey As Boolean
    Dim canRemoveKey As Boolean

    hasFormSelection = (LenB(GetSelectedFormName()) > 0)
    hasLoadedForm = (LenB(Trim$(m_currentFormName)) > 0)
    hasCurrentControl = (LenB(GetCurrentControlName()) > 0)
    canSetSelectedKey = hasCurrentControl And CurrentControlCanReceiveKey()
    canRemoveKey = hasCurrentControl And (LenB(GetCurrentTranslationKey()) > 0 Or StrComp(GetCurrentTagStatus(), modFwTranslationTagGeneratorService.STATUS_INVALID_TAG, vbTextCompare) = 0)

    SetControlEnabledIfPresent "cmdLoadForm", hasFormSelection
    SetControlEnabledIfPresent "cmdGenerateSuggestions", hasLoadedForm
    SetControlEnabledIfPresent "cmdSetMissingKeys", hasLoadedForm
    SetControlEnabledIfPresent "cmdSetSelectedKey", canSetSelectedKey
    SetControlEnabledIfPresent "cmdRemoveSelectedKey", canRemoveKey
    SetControlEnabledIfPresent "cmdSave", modFwTranslationTagGeneratorService.HasPendingChanges(m_sessionId)

SafeExit:
End Sub

Private Sub ApplyDisplayFilter()
    On Error GoTo ErrorHandler

    If ShowAllControls() Then
        Me.Filter = vbNullString
        Me.FilterOn = False
    Else
        Me.Filter = "[is_relevant] = True OR [tag_status] IN ('" & _
                    modFwTranslationTagGeneratorService.STATUS_HAS_KEY & "','" & _
                    modFwTranslationTagGeneratorService.STATUS_MISSING_KEY & "','" & _
                    modFwTranslationTagGeneratorService.STATUS_INVALID_TAG & "')"
        Me.FilterOn = True
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyDisplayFilter", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function CurrentControlCanReceiveKey() As Boolean
    CurrentControlCanReceiveKey = (StrComp(GetCurrentTagStatus(), modFwTranslationTagGeneratorService.STATUS_MISSING_KEY, vbTextCompare) = 0) Or _
                                  (StrComp(GetCurrentTagStatus(), modFwTranslationTagGeneratorService.STATUS_INVALID_TAG, vbTextCompare) = 0)
End Function

Private Sub EnsureCurrentControlCanReceiveKey()
    If Not CurrentControlCanReceiveKey() Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.FW_TRANSLATION_TAG_GENERATOR_NOT_TRANSLATABLE", "Das markierte Control ist nicht fuer einen Translation-Key geeignet."), vbExclamation, MODULE_NAME
        Err.Raise vbObjectError + 5450, MODULE_NAME & ".EnsureCurrentControlCanReceiveKey", "Selected control cannot receive a translation key."
    End If
End Sub

Private Function RequireCurrentControlName() As String
    RequireCurrentControlName = GetCurrentControlName()

    If LenB(RequireCurrentControlName) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.FW_TRANSLATION_TAG_GENERATOR_SELECT_CONTROL", "Bitte zuerst ein Control auswaehlen."), vbExclamation, MODULE_NAME
        Err.Raise vbObjectError + 5451, MODULE_NAME & ".RequireCurrentControlName", "No control is currently selected."
    End If
End Function

Private Sub EnsureFormSelected()
    If LenB(Trim$(m_currentFormName)) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.FW_TRANSLATION_TAG_GENERATOR_FORM_REQUIRED", "Bitte zuerst ein Formular auswaehlen."), vbExclamation, MODULE_NAME
        Err.Raise vbObjectError + 5452, MODULE_NAME & ".EnsureFormSelected", "No form is currently selected."
    End If
End Sub

Private Sub EnsureLoadedForm()
    EnsureFormSelected
End Sub

Private Function CanDiscardChanges() As Boolean
    If Not modFwTranslationTagGeneratorService.HasPendingChanges(m_sessionId) Then
        CanDiscardChanges = True
    ElseIf MsgBox(modFwTranslationRuntime.ResolveText("MSG.FW_TRANSLATION_TAG_GENERATOR_DISCARD_CONFIRM", "Ungespeicherte Tag-Aenderungen verwerfen?"), vbQuestion + vbYesNo + vbDefaultButton2, MODULE_NAME) = vbYes Then
        CanDiscardChanges = True
    End If
End Function

Private Sub ReturnToPreviousWorkspace()
    On Error GoTo SafeExit

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()
    If Not shellForm Is Nothing Then
        If modAppWorkspaceService.GoBack(shellForm) Then
            Exit Sub
        End If
    End If

    DoCmd.Close acForm, Me.Name

SafeExit:
End Sub

Private Function ResolveHostingShellForm() As Access.Form
    On Error GoTo SafeExit

    Dim parentForm As Access.Form

    Set parentForm = Me.Parent
    If parentForm Is Nothing Then
        Exit Function
    End If

    If HasControlOnForm(parentForm, "subWorkspaceHost") Then
        Set ResolveHostingShellForm = parentForm
    End If

SafeExit:
End Function

Private Function GetSelectedFormName() As String
    On Error GoTo SafeExit

    If HasControl("cboFormName") Then
        GetSelectedFormName = Trim$(modDaoHelper.NzString(Me!cboFormName.Value))
    End If

SafeExit:
End Function

Private Sub SetSelectedFormName(ByVal formName As String)
    On Error GoTo SafeExit

    If HasControl("cboFormName") Then
        Me!cboFormName.Value = NullIfEmpty(formName)
    End If

    m_currentFormName = Trim$(modDaoHelper.NzString(formName))

SafeExit:
End Sub

Private Function IncludeHiddenControls() As Boolean
    On Error GoTo SafeExit

    If HasControl("chkIncludeHidden") Then
        IncludeHiddenControls = modDaoHelper.NzBoolean(Me!chkIncludeHidden.Value, False)
    End If

SafeExit:
End Function

Private Sub SetIncludeHiddenControls(ByVal includeHidden As Boolean)
    On Error GoTo SafeExit

    If HasControl("chkIncludeHidden") Then
        Me!chkIncludeHidden.Value = includeHidden
    End If

SafeExit:
End Sub

Private Function ShowAllControls() As Boolean
    On Error GoTo SafeExit

    If HasControl("chkShowAllControls") Then
        ShowAllControls = modDaoHelper.NzBoolean(Me!chkShowAllControls.Value, False)
    End If

SafeExit:
End Function

Private Sub SetShowAllControls(ByVal showAllControls As Boolean)
    On Error GoTo SafeExit

    If HasControl("chkShowAllControls") Then
        Me!chkShowAllControls.Value = showAllControls
    End If

SafeExit:
End Sub

Private Function GetCurrentControlName() As String
    On Error GoTo SafeExit

    If Me.NewRecord Then
        Exit Function
    End If

    If HasControl("control_name") Then
        GetCurrentControlName = Trim$(modDaoHelper.NzString(Me!control_name.Value))
    End If

SafeExit:
End Function

Private Function GetCurrentTranslationKey() As String
    On Error GoTo SafeExit

    If Me.NewRecord Then
        Exit Function
    End If

    If HasControl("current_translation_key") Then
        GetCurrentTranslationKey = Trim$(modDaoHelper.NzString(Me!current_translation_key.Value))
    End If

SafeExit:
End Function

Private Function GetCurrentTagStatus() As String
    On Error GoTo SafeExit

    If Me.NewRecord Then
        Exit Function
    End If

    If HasControl("tag_status") Then
        GetCurrentTagStatus = Trim$(modDaoHelper.NzString(Me!tag_status.Value))
    End If

SafeExit:
End Function

Private Sub RestoreCurrentControlSelection(ByVal controlName As String)
    On Error GoTo ErrorHandler

    Dim rsClone As DAO.Recordset

    controlName = Replace(Trim$(modDaoHelper.NzString(controlName)), "'", "''")
    If LenB(controlName) = 0 Then
        Exit Sub
    End If

    Set rsClone = Me.RecordsetClone
    rsClone.FindFirst "[control_name] = '" & controlName & "'"
    If Not rsClone.NoMatch Then
        Me.Bookmark = rsClone.Bookmark
    End If

CleanExit:
    On Error Resume Next
    If Not rsClone Is Nothing Then rsClone.Close
    Set rsClone = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreCurrentControlSelection", Err
    Resume CleanExit
End Sub

Private Sub SetControlEnabledIfPresent(ByVal controlName As String, ByVal isEnabled As Boolean)
    On Error GoTo SafeExit

    If HasControl(controlName) Then
        Me.Controls(controlName).Enabled = isEnabled
    End If

SafeExit:
End Sub

Private Function HasControl(ByVal controlName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentControl As Control

    For Each currentControl In Me.Controls
        If StrComp(currentControl.Name, controlName, vbTextCompare) = 0 Then
            HasControl = True
            Exit Function
        End If
    Next currentControl

SafeExit:
End Function

Private Function HasControlOnForm(ByVal formInstance As Access.Form, ByVal controlName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentControl As Control

    If formInstance Is Nothing Then
        Exit Function
    End If

    For Each currentControl In formInstance.Controls
        If StrComp(currentControl.Name, controlName, vbTextCompare) = 0 Then
            HasControlOnForm = True
            Exit Function
        End If
    Next currentControl

SafeExit:
End Function

Private Function AddWorkspaceStatePair(ByVal stateText As String, ByVal keyName As String, ByVal valueText As String) As String
    If LenB(stateText) > 0 Then
        stateText = stateText & ";"
    End If

    AddWorkspaceStatePair = stateText & keyName & "=" & EscapeWorkspaceStateValue(valueText)
End Function

Private Function EscapeWorkspaceStateValue(ByVal valueText As String) As String
    valueText = Replace(modDaoHelper.NzString(valueText), "%", "%25")
    valueText = Replace(valueText, ";", "%3B")
    valueText = Replace(valueText, "=", "%3D")
    EscapeWorkspaceStateValue = valueText
End Function

Private Function GetWorkspaceStateValue(ByVal stateText As String, ByVal keyName As String) As String
    Dim parts() As String
    Dim pairText As Variant
    Dim separatorPosition As Long
    Dim currentKey As String

    If LenB(stateText) = 0 Then
        Exit Function
    End If

    parts = Split(stateText, ";")
    For Each pairText In parts
        separatorPosition = InStr(1, CStr(pairText), "=", vbBinaryCompare)
        If separatorPosition > 0 Then
            currentKey = Left$(CStr(pairText), separatorPosition - 1)
            If StrComp(currentKey, keyName, vbTextCompare) = 0 Then
                GetWorkspaceStateValue = UnescapeWorkspaceStateValue(Mid$(CStr(pairText), separatorPosition + 1))
                Exit Function
            End If
        End If
    Next pairText
End Function

Private Function UnescapeWorkspaceStateValue(ByVal valueText As String) As String
    valueText = Replace(valueText, "%3D", "=")
    valueText = Replace(valueText, "%3B", ";")
    valueText = Replace(valueText, "%25", "%")
    UnescapeWorkspaceStateValue = valueText
End Function

Private Function NullIfEmpty(ByVal valueText As String) As Variant
    valueText = modDaoHelper.NzString(valueText)

    If LenB(valueText) = 0 Then
        NullIfEmpty = Null
    Else
        NullIfEmpty = valueText
    End If
End Function

Private Function BuildSessionId() As String
    BuildSessionId = "TTG-" & Replace(Replace(Replace(CStr(Now), ":", vbNullString), ".", vbNullString), " ", "-") & "-" & CStr(Timer * 1000)
End Function
