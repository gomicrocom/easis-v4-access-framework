Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmTagComposer
' Purpose   : Tag composer for framework-managed control tags on forms and
'             reports.
' Author    : Codex
' Version   : 0.2.1
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmTagComposer"
Private Const OBJECT_TYPE_FORM As String = "FORM"
Private Const OBJECT_TYPE_REPORT As String = "REPORT"
Private Const TEMP_TABLE_NAME As String = "fw_tmp_tag_composer"
Private Const TAG_HELP_TABLE_NAME As String = "fw_tag_help"
Private Const FIELD_SESSION_ID As String = "session_id"
Private Const FIELD_ROW_NO As String = "row_no"
Private Const FIELD_CONTROL_NAME As String = "control_name"
Private Const FIELD_CONTROL_TYPE As String = "control_type"
Private Const FIELD_TAG_VALUE As String = "tag_value"
Private Const FIELD_IS_DIRTY As String = "is_dirty"

Private m_TargetObjectType As String
Private m_TargetObjectName As String
Private m_SessionId As String
Private m_CurrentControl As String

Public Property Get CurrentComposerControlName() As String
    CurrentComposerControlName = m_CurrentControl
End Property

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    m_SessionId = BuildSessionId()

    PopulateObjectTypeCombo
    Me!cboObjectType = OBJECT_TYPE_FORM
    LoadObjectList
    ResetComposerState
    UpdateLocalCommandButtons

    If LenB(Trim$(Nz(Me.openArgs, vbNullString))) > 0 Then
        PreselectLegacyForm Nz(Me.openArgs, vbNullString)
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox "Fehler beim Initialisieren des Tag-Composer: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cboObjectType_AfterUpdate()
    On Error GoTo ErrorHandler

    SaveCurrentControlTagToTemp False
    ResetComposerState
    LoadObjectList
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboObjectType_AfterUpdate", Err
    MsgBox "Fehler beim Laden der Objektliste: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub lstObjects_AfterUpdate()
    On Error GoTo ErrorHandler

    SaveCurrentControlTagToTemp False
    LoadSelectedObject
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "lstObjects_AfterUpdate", Err
    MsgBox "Fehler beim Laden des Objekts: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdHelp_Click()
    DoCmd.OpenForm "frmTagHelp", WindowMode:=acDialog
End Sub

Private Sub cmdApply_Click()
    If LenB(m_CurrentControl) = 0 Then
        MsgBox "Bitte zuerst ein Control in der Liste auswählen.", vbExclamation
        Exit Sub
    End If

    SaveCurrentControlTagToTemp True
    MsgBox "Tag für '" & m_CurrentControl & "' übernommen.", vbInformation
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler

    DetailSave
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdSave_Click", Err
End Sub

Private Sub cmdClose_Click()
    On Error GoTo ErrorHandler

    DetailCancel
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdClose_Click", Err
End Sub

Public Function SupportsListCommandBar() As Boolean
    SupportsListCommandBar = False
End Function

Public Function SupportsWorkspaceCommandBar() As Boolean
    SupportsWorkspaceCommandBar = True
End Function

Public Function GetWorkspaceCommandBarConfig() As Object
    On Error GoTo ErrorHandler

    Dim cfg As Object
    Dim shellForm As Access.Form
    Dim hasShellHost As Boolean

    Set cfg = CreateObject("Scripting.Dictionary")
    cfg.CompareMode = vbTextCompare
    Set shellForm = ResolveHostingShellForm()
    hasShellHost = Not shellForm Is Nothing

    cfg("search.visible") = False
    cfg("search.enabled") = False

    cfg("L1.visible") = True
    cfg("L1.enabled") = hasShellHost
    cfg("L1.caption") = modFwTranslationRuntime.ResolveText("COMMON.HOME", "Home")
    cfg("L1.command") = modAppShell.WCMD_NAV_HOME

    cfg("L2.visible") = True
    cfg("L2.enabled") = hasShellHost And modAppWorkspaceService.CanGoBack()
    cfg("L2.caption") = modFwTranslationRuntime.ResolveText("COMMON.BACK", "Zurueck")
    cfg("L2.command") = modAppShell.WCMD_NAV_BACK

    cfg("L3.visible") = False
    cfg("L3.enabled") = False
    cfg("L3.caption") = vbNullString
    cfg("L3.command") = vbNullString

    cfg("R2.visible") = True
    cfg("R2.enabled") = True
    cfg("R2.caption") = modFwTranslationRuntime.ResolveText("COMMON.CANCEL", "Abbrechen")
    cfg("R2.command") = modAppShell.WCMD_DETAIL_CANCEL

    cfg("R1.visible") = True
    cfg("R1.enabled") = HasSelectedTargetObject()
    cfg("R1.caption") = modFwTranslationRuntime.ResolveText("COMMON.SAVE", "Speichern")
    cfg("R1.command") = modAppShell.WCMD_DETAIL_SAVE

    Set GetWorkspaceCommandBarConfig = cfg
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetWorkspaceCommandBarConfig", Err
End Function

Public Function CanExecuteWorkspaceCommand(ByVal commandKey As String) As Boolean
    On Error GoTo ErrorHandler

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()

    Select Case UCase$(Trim$(commandKey))
        Case modAppShell.WCMD_NAV_HOME
            CanExecuteWorkspaceCommand = Not shellForm Is Nothing
        Case modAppShell.WCMD_NAV_BACK
            CanExecuteWorkspaceCommand = (Not shellForm Is Nothing) And modAppWorkspaceService.CanGoBack()
        Case modAppShell.WCMD_DETAIL_CANCEL
            CanExecuteWorkspaceCommand = True
        Case modAppShell.WCMD_DETAIL_SAVE
            CanExecuteWorkspaceCommand = HasSelectedTargetObject()
    End Select
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CanExecuteWorkspaceCommand", Err
End Function

Public Sub ExecuteWorkspaceCommand(ByVal commandKey As String, Optional ByVal commandValue As String = "")
    On Error GoTo ErrorHandler

    Select Case UCase$(Trim$(commandKey))
        Case modAppShell.WCMD_NAV_HOME
            NavigateHome
        Case modAppShell.WCMD_NAV_BACK
            NavigateBack
        Case modAppShell.WCMD_DETAIL_CANCEL
            DetailCancel
        Case modAppShell.WCMD_DETAIL_SAVE
            DetailSave
    End Select
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ExecuteWorkspaceCommand", Err
End Sub

Public Function CanLeaveWorkspace() As Boolean
    CanLeaveWorkspace = True
End Function

Public Sub DetailSave()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim WasLoaded As Boolean
    Dim WasVisible As Boolean
    Dim OpenedForSave As Boolean

    If Not HasSelectedTargetObject() Then
        MsgBox "Bitte zuerst ein Formular oder einen Report auswählen.", vbExclamation, MODULE_NAME
        Exit Sub
    End If

    SaveCurrentControlTagToTemp False

    WasLoaded = IsTargetObjectLoaded(m_TargetObjectType, m_TargetObjectName)
    WasVisible = GetTargetObjectVisible(m_TargetObjectType, m_TargetObjectName)

    PrepareTargetObjectForSave m_TargetObjectType, m_TargetObjectName, OpenedForSave

    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT " & FIELD_CONTROL_NAME & ", " & FIELD_TAG_VALUE & " FROM " & TEMP_TABLE_NAME & " " & _
        "WHERE " & FIELD_SESSION_ID & " = '" & Replace(m_SessionId, "'", "''") & "' " & _
        "ORDER BY " & FIELD_ROW_NO & ";", dbOpenSnapshot)

    Do While Not rs.EOF
        WriteTagToTargetControl m_TargetObjectType, m_TargetObjectName, _
            Nz(rs.Fields(FIELD_CONTROL_NAME).Value, vbNullString), _
            Nz(rs.Fields(FIELD_TAG_VALUE).Value, vbNullString)
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    SaveAndCloseTargetObject m_TargetObjectType, m_TargetObjectName
    RestoreTargetObjectState m_TargetObjectType, m_TargetObjectName, WasLoaded, WasVisible

    CleanupTempSession

    MsgBox "Tags wurden dauerhaft gespeichert.", vbInformation, MODULE_NAME
    DoCmd.Close acForm, Me.Name
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "DetailSave", Err
    On Error Resume Next
    If OpenedForSave Then
        CloseObjectNoSave m_TargetObjectType, m_TargetObjectName
    End If
    RestoreTargetObjectState m_TargetObjectType, m_TargetObjectName, WasLoaded, WasVisible
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    MsgBox "Fehler beim Speichern der Tags: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Public Sub DetailCancel()
    CleanupTempSession
    DoCmd.Close acForm, Me.Name
End Sub

Public Sub LoadSelectedControlFromSubform(ByVal ControlName As String, ByVal TagValue As String)
    On Error GoTo ErrorHandler

    SaveCurrentControlTagToTemp False

    m_CurrentControl = Nz(ControlName, vbNullString)

    ResetUI
    LoadFromTag Nz(TagValue, vbNullString)

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadSelectedControlFromSubform", Err
    MsgBox "Fehler beim Laden des Controls: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub PopulateObjectTypeCombo()
    Me!cboObjectType.RowSourceType = "Value List"
    Me!cboObjectType.RowSource = OBJECT_TYPE_FORM & ";" & OBJECT_TYPE_REPORT
End Sub

Private Sub LoadObjectList()
    On Error GoTo ErrorHandler

    Dim values As Collection
    Dim accessObject As Access.accessObject
    Dim objectType As String

    objectType = NormalizeObjectType(Nz(Me!cboObjectType.Value, OBJECT_TYPE_FORM))
    Set values = New Collection

    Select Case objectType
        Case OBJECT_TYPE_FORM
            For Each accessObject In CurrentProject.AllForms
                If Left$(accessObject.Name, 1) <> "~" Then
                    If Not IsComposerInternalObject(accessObject.Name) Then
                        CollectionAddSorted values, accessObject.Name
                    End If
                End If
            Next accessObject

        Case OBJECT_TYPE_REPORT
            For Each accessObject In CurrentProject.AllReports
                If Left$(accessObject.Name, 1) <> "~" Then
                    CollectionAddSorted values, accessObject.Name
                End If
            Next accessObject
    End Select

    LoadCollectionToListBox Me!lstObjects, values
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadObjectList", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub LoadSelectedObject()
    On Error GoTo ErrorHandler

    Dim selectedObjectName As String

    selectedObjectName = Trim$(Nz(Me!lstObjects.Value, vbNullString))
    If LenB(selectedObjectName) = 0 Then
        ResetComposerState False
        Exit Sub
    End If

    m_TargetObjectType = NormalizeObjectType(Nz(Me!cboObjectType.Value, OBJECT_TYPE_FORM))
    m_TargetObjectName = selectedObjectName
    m_CurrentControl = vbNullString

    InitializeTempComposerData
    BindTempSubform
    ResetUI
    RefreshShellStatusIfHosted
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadSelectedObject", Err
    MsgBox "Fehler beim Initialisieren des ausgewählten Objekts: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub ResetComposerState(Optional ByVal ClearObjectSelection As Boolean = True)
    On Error Resume Next

    CleanupTempSession

    m_TargetObjectType = vbNullString
    m_TargetObjectName = vbNullString
    m_CurrentControl = vbNullString

    If ClearObjectSelection Then
        Me!lstObjects = Null
    End If

    Me!txtResult = Null
    ResetUI
    BindEmptySubform
    UpdateLocalCommandButtons
    RefreshShellStatusIfHosted
End Sub

Private Sub InitializeTempComposerData()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim RowNo As Long
    Dim controlTypeName As String
    Dim OpenedByComposer As Boolean
    Dim ctl As Control

    If LenB(m_TargetObjectType) = 0 Or LenB(m_TargetObjectName) = 0 Then
        Exit Sub
    End If

    Set db = CurrentDb
    db.Execute "DELETE FROM " & TEMP_TABLE_NAME & " WHERE " & FIELD_SESSION_ID & " = '" & Replace(m_SessionId, "'", "''") & "'", dbFailOnError

    OpenTargetObjectForRead m_TargetObjectType, m_TargetObjectName, OpenedByComposer
    Set rs = db.OpenRecordset(TEMP_TABLE_NAME, dbOpenDynaset)

    RowNo = 0
    For Each ctl In GetTargetControls(m_TargetObjectType, m_TargetObjectName)
        RowNo = RowNo + 1
        controlTypeName = GetControlTypeName(ctl)

        rs.AddNew
        rs.Fields(FIELD_SESSION_ID).Value = m_SessionId
        rs.Fields(FIELD_ROW_NO).Value = RowNo
        rs.Fields(FIELD_CONTROL_NAME).Value = ctl.Name
        rs.Fields(FIELD_CONTROL_TYPE).Value = controlTypeName
        rs.Fields(FIELD_TAG_VALUE).Value = GetControlTagSafely(ctl)
        rs.Fields(FIELD_IS_DIRTY).Value = False
        rs.Update
    Next ctl

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If OpenedByComposer Then
        CloseObjectNoSave m_TargetObjectType, m_TargetObjectName
    End If
    Set db = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "InitializeTempComposerData", Err
    Resume CleanExit
End Sub

Private Sub BindTempSubform()
    On Error GoTo ErrorHandler

    Dim sqlStatement As String

    sqlStatement = "SELECT " & FIELD_SESSION_ID & ", " & FIELD_ROW_NO & ", " & _
                   FIELD_CONTROL_NAME & ", " & FIELD_CONTROL_TYPE & ", " & _
                   FIELD_TAG_VALUE & ", " & FIELD_IS_DIRTY & " " & _
              "FROM " & TEMP_TABLE_NAME & " " & _
              "WHERE " & FIELD_SESSION_ID & " = '" & Replace(m_SessionId, "'", "''") & "' " & _
                   "ORDER BY " & FIELD_ROW_NO & ";"

    Me!sfrmTmpTagComposer.Form.RecordSource = sqlStatement
    Me!sfrmTmpTagComposer.Form.Requery
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "BindTempSubform", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub BindEmptySubform()
    On Error Resume Next

    Me!sfrmTmpTagComposer.Form.RecordSource = _
        "SELECT " & FIELD_SESSION_ID & ", " & FIELD_ROW_NO & ", " & _
        FIELD_CONTROL_NAME & ", " & FIELD_CONTROL_TYPE & ", " & _
        FIELD_TAG_VALUE & ", " & FIELD_IS_DIRTY & " FROM " & TEMP_TABLE_NAME & " WHERE 1=0;"
    Me!sfrmTmpTagComposer.Form.Requery
End Sub

Private Function BuildSessionId() As String
    Randomize
    BuildSessionId = "TC_" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int((Rnd() * 90000) + 10000))
End Function

Private Sub PreselectLegacyForm(ByVal FormName As String)
    On Error GoTo ErrorHandler

    FormName = Trim$(FormName)
    If LenB(FormName) = 0 Then
        Exit Sub
    End If

    If Not AccessObjectExists(OBJECT_TYPE_FORM, FormName) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".PreselectLegacyForm", _
            "OpenArgs form not found: " & FormName & "."
        Exit Sub
    End If

    Me!cboObjectType = OBJECT_TYPE_FORM
    LoadObjectList
    SetListBoxValue Me!lstObjects, FormName
    LoadSelectedObject
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "PreselectLegacyForm", Err
End Sub

Private Sub SaveCurrentControlTagToTemp(Optional ByVal RefreshSubform As Boolean = False)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim sqlStatement As String
    Dim newTag As String

    If LenB(m_CurrentControl) = 0 Then Exit Sub

    newTag = BuildTagString()

    Set db = CurrentDb
    sqlStatement = "UPDATE " & TEMP_TABLE_NAME & " " & _
                   "SET " & FIELD_TAG_VALUE & " = '" & Replace(newTag, "'", "''") & "', " & _
                   "    " & FIELD_IS_DIRTY & " = True " & _
              "WHERE " & FIELD_SESSION_ID & " = '" & Replace(m_SessionId, "'", "''") & "' " & _
                   "  AND " & FIELD_CONTROL_NAME & " = '" & Replace(m_CurrentControl, "'", "''") & "';"

    db.Execute sqlStatement, dbFailOnError

    Me!txtResult = newTag

    If RefreshSubform Then
        RefreshSubformCurrentRow m_CurrentControl
    End If

    Set db = Nothing
    Exit Sub

ErrorHandler:
    Set db = Nothing
    modErrorHandler.HandleError MODULE_NAME, "SaveCurrentControlTagToTemp", Err
    MsgBox "Fehler beim Zwischenspeichern des Tags: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Function NormalizeObjectType(ByVal Value As String) As String
    Value = UCase$(Trim$(Value))

    Select Case Value
        Case OBJECT_TYPE_FORM, OBJECT_TYPE_REPORT
            NormalizeObjectType = Value
        Case Else
            Err.Raise vbObjectError + 5100, MODULE_NAME & ".NormalizeObjectType", _
                "Nicht unterstützter Objekttyp: " & Value
    End Select
End Function

Private Function IsComposerInternalObject(ByVal objectName As String) As Boolean
    Select Case UCase$(Trim$(objectName))
        Case "FRMTAGCOMPOSER", "FRMTAGHELP", "SFRMTMPTAGCOMPOSER", "FRMFWTRANSLATIONS", "FRMFWTRANSLATIONLIST"
            IsComposerInternalObject = True
        Case Else
            IsComposerInternalObject = False
    End Select
End Function

Private Function AccessObjectExists(ByVal objectType As String, ByVal objectName As String) As Boolean
    On Error GoTo SafeExit

    Dim accessObject As Access.accessObject

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            For Each accessObject In CurrentProject.AllForms
                If StrComp(accessObject.Name, objectName, vbTextCompare) = 0 Then
                    AccessObjectExists = True
                    Exit Function
                End If
            Next accessObject

        Case OBJECT_TYPE_REPORT
            For Each accessObject In CurrentProject.AllReports
                If StrComp(accessObject.Name, objectName, vbTextCompare) = 0 Then
                    AccessObjectExists = True
                    Exit Function
                End If
            Next accessObject
    End Select

SafeExit:
End Function

Private Function IsTargetObjectLoaded(ByVal objectType As String, ByVal objectName As String) As Boolean
    On Error GoTo SafeExit

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            IsTargetObjectLoaded = CurrentProject.AllForms(objectName).IsLoaded
        Case OBJECT_TYPE_REPORT
            IsTargetObjectLoaded = CurrentProject.AllReports(objectName).IsLoaded
    End Select

SafeExit:
End Function

Private Function GetTargetObjectVisible(ByVal objectType As String, ByVal objectName As String) As Boolean
    On Error GoTo SafeExit

    If NormalizeObjectType(objectType) = OBJECT_TYPE_FORM Then
        If CurrentProject.AllForms(objectName).IsLoaded Then
            GetTargetObjectVisible = Forms(objectName).Visible
        End If
    End If

SafeExit:
End Function

Private Sub OpenTargetObjectForRead(ByVal objectType As String, ByVal objectName As String, ByRef OpenedByComposer As Boolean)
    On Error GoTo ErrorHandler

    OpenedByComposer = False

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            If Not CurrentProject.AllForms(objectName).IsLoaded Then
                DoCmd.OpenForm objectName, acDesign, , , , acHidden
                OpenedByComposer = True
            End If

        Case OBJECT_TYPE_REPORT
            If Not CurrentProject.AllReports(objectName).IsLoaded Then
                DoCmd.OpenReport objectName, acViewDesign, , , acHidden
                OpenedByComposer = True
            End If
    End Select
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "OpenTargetObjectForRead", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub PrepareTargetObjectForSave(ByVal objectType As String, ByVal objectName As String, ByRef OpenedForSave As Boolean)
    On Error GoTo ErrorHandler

    OpenedForSave = False

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            If CurrentProject.AllForms(objectName).IsLoaded Then
                DoCmd.Close acForm, objectName, acSaveYes
            End If
            DoCmd.OpenForm objectName, acDesign, , , , acHidden
            OpenedForSave = True

        Case OBJECT_TYPE_REPORT
            If CurrentProject.AllReports(objectName).IsLoaded Then
                DoCmd.Close acReport, objectName, acSaveYes
            End If
            DoCmd.OpenReport objectName, acViewDesign, , , acHidden
            OpenedForSave = True
    End Select
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "PrepareTargetObjectForSave", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub SaveAndCloseTargetObject(ByVal objectType As String, ByVal objectName As String)
    On Error GoTo ErrorHandler

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            DoCmd.Save acForm, objectName
            DoCmd.Close acForm, objectName, acSaveYes

        Case OBJECT_TYPE_REPORT
            DoCmd.Save acReport, objectName
            DoCmd.Close acReport, objectName, acSaveYes
    End Select
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SaveAndCloseTargetObject", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub RestoreTargetObjectState( _
    ByVal objectType As String, _
    ByVal objectName As String, _
    ByVal WasLoaded As Boolean, _
    ByVal WasVisible As Boolean)
    On Error GoTo SafeExit

    If Not WasLoaded Then
        Exit Sub
    End If

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            If Not CurrentProject.AllForms(objectName).IsLoaded Then
                DoCmd.OpenForm objectName, acNormal
            End If
            If Not WasVisible Then
                Forms(objectName).Visible = False
            End If

        Case OBJECT_TYPE_REPORT
            If Not CurrentProject.AllReports(objectName).IsLoaded Then
                DoCmd.OpenReport objectName, acViewPreview
            End If
    End Select

SafeExit:
End Sub

Private Sub CloseObjectNoSave(ByVal objectType As String, ByVal objectName As String)
    On Error Resume Next

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            If CurrentProject.AllForms(objectName).IsLoaded Then
                DoCmd.Close acForm, objectName, acSaveNo
            End If

        Case OBJECT_TYPE_REPORT
            If CurrentProject.AllReports(objectName).IsLoaded Then
                DoCmd.Close acReport, objectName, acSaveNo
            End If
    End Select
End Sub

Private Function GetTargetControls(ByVal objectType As String, ByVal objectName As String) As Controls
    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            Set GetTargetControls = Forms(objectName).Controls
        Case OBJECT_TYPE_REPORT
            Set GetTargetControls = Reports(objectName).Controls
    End Select
End Function

Private Sub WriteTagToTargetControl( _
    ByVal objectType As String, _
    ByVal objectName As String, _
    ByVal ControlName As String, _
    ByVal TagValue As String)
    On Error GoTo ErrorHandler

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            Forms(objectName).Controls(ControlName).Tag = TagValue
        Case OBJECT_TYPE_REPORT
            Reports(objectName).Controls(ControlName).Tag = TagValue
    End Select
    Exit Sub

ErrorHandler:
    modLoggingHandler.LogWarning MODULE_NAME & ".WriteTagToTargetControl", _
        "Tag write failed for " & objectType & "." & objectName & "." & ControlName & ": " & Err.description
    Err.Clear
End Sub

Private Function GetControlTagSafely(ByVal ctl As Control) As String
    On Error GoTo SafeExit

    GetControlTagSafely = Nz(ctl.Tag, vbNullString)
    Exit Function

SafeExit:
    GetControlTagSafely = vbNullString
End Function

Private Function GetControlTypeName(ByVal ctl As Control) As String
    On Error GoTo SafeExit

    Select Case ctl.ControlType
        Case acTextBox: GetControlTypeName = "TextBox"
        Case acComboBox: GetControlTypeName = "ComboBox"
        Case acCheckBox: GetControlTypeName = "CheckBox"
        Case acLabel: GetControlTypeName = "Label"
        Case acCommandButton: GetControlTypeName = "Button"
        Case acListBox: GetControlTypeName = "ListBox"
        Case acSubform: GetControlTypeName = "Subform"
        Case acOptionGroup: GetControlTypeName = "OptionGroup"
        Case acOptionButton: GetControlTypeName = "OptionButton"
        Case acToggleButton: GetControlTypeName = "ToggleButton"
        Case acTabCtl: GetControlTypeName = "TabControl"
        Case Else: GetControlTypeName = CStr(ctl.ControlType)
    End Select
    Exit Function

SafeExit:
    GetControlTypeName = vbNullString
End Function

Private Sub CleanupTempSession()
    On Error Resume Next

    If LenB(m_SessionId) = 0 Then Exit Sub
    CurrentDb.Execute "DELETE FROM " & TEMP_TABLE_NAME & " WHERE " & FIELD_SESSION_ID & " = '" & Replace(m_SessionId, "'", "''") & "';", dbFailOnError
End Sub

Private Function BuildTagString() As String
    Dim parts As Collection
    Dim preservedParts As Collection
    Dim item As Variant
    Dim ExistingTag As String

    Set parts = New Collection

    ExistingTag = GetExistingTagForCurrentControl()
    Set preservedParts = GetPreservedTagParts(ExistingTag)

    For Each item In preservedParts
        parts.Add CStr(item)
    Next item

    If Nz(Me!chkRequired, False) Then parts.Add "REQUIRED"
    If Nz(Me!chkNumeric, False) Then parts.Add "NUMERIC"
    If Nz(Me!chkInteger, False) Then parts.Add "INTEGER"
    If Nz(Me!chkDate, False) Then parts.Add "DATE"

    If Len(Nz(Me!txtMin, "")) > 0 Then parts.Add "MIN:" & Trim$(Me!txtMin)
    If Len(Nz(Me!txtMax, "")) > 0 Then parts.Add "MAX:" & Trim$(Me!txtMax)

    If Len(Nz(Me!txtMinLen, "")) > 0 Then parts.Add "MINLEN:" & Trim$(Me!txtMinLen)
    If Len(Nz(Me!txtMaxLen, "")) > 0 Then parts.Add "MAXLEN:" & Trim$(Me!txtMaxLen)

    If Nz(Me!chkHidden, False) Then parts.Add "HIDDEN"
    If Nz(Me!chkSetFocus, False) Then parts.Add "SETFOCUS"
    If Nz(Me!chkReadOnly, False) Then parts.Add "READONLY"
    If Nz(Me!chkLocked, False) Then parts.Add "LOCKED"
    If Nz(Me!chkDisabled, False) Then parts.Add "DISABLED"

    If Len(Nz(Me!txtRole, "")) > 0 Then parts.Add "ROLE:" & Trim$(Me!txtRole)
    If Len(Nz(Me!txtMod, "")) > 0 Then parts.Add "MOD:" & Trim$(Me!txtMod)

    BuildTagString = JoinCollection(parts, ";")
End Function

Private Function JoinCollection(col As Collection, ByVal sep As String) As String
    Dim item As Variant

    For Each item In col
        If Len(JoinCollection) > 0 Then
            JoinCollection = JoinCollection & sep
        End If
        JoinCollection = JoinCollection & CStr(item)
    Next item
End Function

Private Sub UpdatePreview()
    Me!txtResult = BuildTagString()
End Sub

Private Sub chkRequired_AfterUpdate(): UpdatePreview: End Sub
Private Sub chkNumeric_AfterUpdate(): UpdatePreview: End Sub
Private Sub chkDate_AfterUpdate(): UpdatePreview: End Sub
Private Sub txtMinLen_AfterUpdate(): UpdatePreview: End Sub
Private Sub txtMaxLen_AfterUpdate(): UpdatePreview: End Sub
Private Sub chkHidden_AfterUpdate(): UpdatePreview: End Sub
Private Sub chkSetFocus_AfterUpdate(): UpdatePreview: End Sub
Private Sub chkReadOnly_AfterUpdate(): UpdatePreview: End Sub
Private Sub chkLocked_AfterUpdate(): UpdatePreview: End Sub
Private Sub chkDisabled_AfterUpdate(): UpdatePreview: End Sub
Private Sub txtRole_AfterUpdate(): UpdatePreview: End Sub
Private Sub txtMod_AfterUpdate(): UpdatePreview: End Sub

Private Sub chkInteger_AfterUpdate()
    If Me!chkInteger Then Me!chkNumeric = True
    UpdatePreview
End Sub

Private Sub txtMin_AfterUpdate()
    If Len(Nz(Me!txtMin, "")) > 0 Then Me!chkNumeric = True
    UpdatePreview
End Sub

Private Sub txtMax_AfterUpdate()
    If Len(Nz(Me!txtMax, "")) > 0 Then Me!chkNumeric = True
    UpdatePreview
End Sub

Public Sub LoadFromTag(ByVal TagValue As String)
    Dim tokens As Object
    Set tokens = ParseTagTokens(TagValue)

    ResetUI

    Me!chkRequired = tokens.Exists("REQUIRED")
    Me!chkNumeric = tokens.Exists("NUMERIC")
    Me!chkInteger = tokens.Exists("INTEGER")
    Me!chkDate = tokens.Exists("DATE")

    Me!chkHidden = tokens.Exists("HIDDEN")
    Me!chkSetFocus = tokens.Exists("SETFOCUS")
    Me!chkReadOnly = tokens.Exists("READONLY")
    Me!chkLocked = tokens.Exists("LOCKED")
    Me!chkDisabled = tokens.Exists("DISABLED")

    If tokens.Exists("MIN") Then Me!txtMin = tokens("MIN")
    If tokens.Exists("MAX") Then Me!txtMax = tokens("MAX")
    If tokens.Exists("MINLEN") Then Me!txtMinLen = tokens("MINLEN")
    If tokens.Exists("MAXLEN") Then Me!txtMaxLen = tokens("MAXLEN")
    If tokens.Exists("ROLE") Then Me!txtRole = tokens("ROLE")
    If tokens.Exists("MOD") Then Me!txtMod = tokens("MOD")

    UpdatePreview
End Sub

Private Function ParseTagTokens(ByVal TagValue As String) As Object
    On Error GoTo ErrorHandler

    Dim parsedTokens As Object
    Dim tagParts() As String
    Dim token As Variant
    Dim trimmedToken As String
    Dim separatorPosition As Long
    Dim TokenKey As String
    Dim tokenValue As String

    Set parsedTokens = CreateObject("Scripting.Dictionary")
    parsedTokens.CompareMode = vbTextCompare

    trimmedToken = Trim$(TagValue)
    If LenB(trimmedToken) = 0 Then
        Set ParseTagTokens = parsedTokens
        Exit Function
    End If

    tagParts = Split(trimmedToken, ";")
    For Each token In tagParts
        trimmedToken = Trim$(CStr(token))
        If LenB(trimmedToken) = 0 Then GoTo NextToken

        separatorPosition = InStr(1, trimmedToken, ":", vbTextCompare)
        If separatorPosition > 0 Then
            TokenKey = UCase$(Trim$(Left$(trimmedToken, separatorPosition - 1)))
            tokenValue = Trim$(Mid$(trimmedToken, separatorPosition + 1))

            If LenB(TokenKey) > 0 Then
                parsedTokens(TokenKey) = tokenValue
            End If
        Else
            TokenKey = UCase$(trimmedToken)
            If LenB(TokenKey) > 0 Then
                parsedTokens(TokenKey) = True
            End If
        End If

NextToken:
    Next token

    Set ParseTagTokens = parsedTokens
    Exit Function

ErrorHandler:
    Set ParseTagTokens = CreateObject("Scripting.Dictionary")
    ParseTagTokens.CompareMode = vbTextCompare
End Function

Private Sub ResetUI()
    On Error Resume Next

    Me!chkRequired = False
    Me!chkNumeric = False
    Me!chkInteger = False
    Me!chkDate = False

    Me!chkHidden = False
    Me!chkSetFocus = False
    Me!chkReadOnly = False
    Me!chkLocked = False
    Me!chkDisabled = False

    Me!txtMin = Null
    Me!txtMax = Null
    Me!txtMinLen = Null
    Me!txtMaxLen = Null
    Me!txtRole = Null
    Me!txtMod = Null
    Me!txtResult = Null
End Sub

Private Sub RefreshSubformCurrentRow(ByVal ControlName As String)
    On Error GoTo ErrorHandler

    Dim rsClone As DAO.Recordset

    Me!sfrmTmpTagComposer.Form.Requery

    Set rsClone = Me!sfrmTmpTagComposer.Form.RecordsetClone
    rsClone.FindFirst FIELD_CONTROL_NAME & " = '" & Replace(ControlName, "'", "''") & "'"

    If Not rsClone.NoMatch Then
        Me!sfrmTmpTagComposer.Form.Bookmark = rsClone.Bookmark
    End If

    rsClone.Close
    Set rsClone = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not rsClone Is Nothing Then rsClone.Close
    Set rsClone = Nothing
End Sub

Private Function GetPreservedTagParts(ByVal ExistingTag As String) As Collection
    Dim parts As Collection
    Dim tokens() As String
    Dim token As Variant
    Dim trimmedToken As String
    Dim TokenKey As String
    Dim separatorPosition As Long

    Set parts = New Collection

    ExistingTag = Trim$(Nz(ExistingTag, ""))
    If Len(ExistingTag) = 0 Then
        Set GetPreservedTagParts = parts
        Exit Function
    End If

    tokens = Split(ExistingTag, ";")

    For Each token In tokens
        trimmedToken = Trim$(CStr(token))
        If Len(trimmedToken) = 0 Then GoTo NextToken

        separatorPosition = InStr(1, trimmedToken, ":", vbTextCompare)
        If separatorPosition > 0 Then
            TokenKey = UCase$(Trim$(Left$(trimmedToken, separatorPosition - 1)))
        Else
            TokenKey = UCase$(trimmedToken)
        End If

        If Not IsComposerManagedToken(TokenKey) Then
            parts.Add trimmedToken
        End If

NextToken:
    Next token

    Set GetPreservedTagParts = parts
End Function

Private Function IsComposerManagedToken(ByVal TokenKey As String) As Boolean
    TokenKey = UCase$(Trim$(Nz(TokenKey, vbNullString)))

    If LenB(TokenKey) = 0 Then
        Exit Function
    End If

    If IsActiveTagHelpToken(TokenKey) Then
        Select Case TokenKey
            Case "REQUIRED", "NUMERIC", "INTEGER", "DATE", _
                 "MIN", "MAX", "MINLEN", "MAXLEN", _
                 "HIDDEN", "SETFOCUS", "READONLY", "LOCKED", "DISABLED", _
                 "ROLE", "MOD"
                IsComposerManagedToken = True
                Exit Function
        End Select
    End If

    Select Case TokenKey
        Case "REQUIRED", "NUMERIC", "INTEGER", "DATE", _
             "MIN", "MAX", "MINLEN", "MAXLEN", _
             "HIDDEN", "SETFOCUS", "READONLY", "LOCKED", "DISABLED", _
             "ROLE", "MOD"
            IsComposerManagedToken = True
    End Select
End Function

Private Function IsActiveTagHelpToken(ByVal TokenKey As String) As Boolean
    On Error GoTo SafeExit

    Dim rs As DAO.Recordset
    Dim sqlStatement As String

    If Not TableExists(TAG_HELP_TABLE_NAME) Then
        Exit Function
    End If

    sqlStatement = "SELECT TokenKey FROM " & TAG_HELP_TABLE_NAME & " " & _
              "WHERE IsActive = True AND TokenKey = '" & Replace(TokenKey, "'", "''") & "';"

    Set rs = CurrentDb.OpenRecordset(sqlStatement, dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then
        IsActiveTagHelpToken = True
    End If

SafeExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

Private Function GetExistingTagForCurrentControl() As String
    On Error GoTo SafeExit

    Dim rs As DAO.Recordset
    Dim sqlStatement As String

    If LenB(m_CurrentControl) = 0 Then Exit Function
    If LenB(m_SessionId) = 0 Then Exit Function

    sqlStatement = "SELECT " & FIELD_TAG_VALUE & " " & _
              "FROM " & TEMP_TABLE_NAME & " " & _
              "WHERE " & FIELD_SESSION_ID & " = '" & Replace(m_SessionId, "'", "''") & "' " & _
                   "  AND " & FIELD_CONTROL_NAME & " = '" & Replace(m_CurrentControl, "'", "''") & "';"

    Set rs = CurrentDb.OpenRecordset(sqlStatement, dbOpenSnapshot)

    If Not rs.EOF Then
        GetExistingTagForCurrentControl = Nz(rs.Fields(FIELD_TAG_VALUE).Value, vbNullString)
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

SafeExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

Private Sub LoadCollectionToListBox(ByVal lst As Access.ListBox, ByVal values As Collection)
    On Error GoTo ErrorHandler

    Dim itemValue As Variant
    Dim rowSourceText As String

    lst.RowSourceType = "Value List"
    lst.RowSource = vbNullString

    If values Is Nothing Then Exit Sub

    For Each itemValue In values
        If Len(rowSourceText) > 0 Then
            rowSourceText = rowSourceText & ";"
        End If

        rowSourceText = rowSourceText & Replace(CStr(itemValue), ";", "\;")
    Next itemValue

    lst.RowSource = rowSourceText

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadCollectionToListBox", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub SetListBoxValue(ByVal lst As Access.ListBox, ByVal ValueToSelect As String)
    On Error GoTo ErrorHandler

    Dim i As Long

    For i = 0 To lst.ListCount - 1
        If StrComp(Nz(lst.ItemData(i), vbNullString), ValueToSelect, vbTextCompare) = 0 Then
            lst = lst.ItemData(i)
            Exit For
        End If
    Next i
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SetListBoxValue", Err
End Sub

Private Sub CollectionAddSorted(ByVal target As Collection, ByVal itemValue As String)
    On Error GoTo ErrorHandler

    Dim tempValues() As String
    Dim i As Long
    Dim j As Long
    Dim inserted As Boolean
    Dim currentValue As String

    If target Is Nothing Then
        Exit Sub
    End If

    If target.count = 0 Then
        target.Add itemValue
        Exit Sub
    End If

    ReDim tempValues(1 To target.count + 1)

    For i = 1 To target.count
        tempValues(i) = CStr(target(i))
    Next i

    tempValues(target.count + 1) = itemValue

    For i = LBound(tempValues) To UBound(tempValues) - 1
        For j = i + 1 To UBound(tempValues)
            If StrComp(tempValues(i), tempValues(j), vbTextCompare) > 0 Then
                currentValue = tempValues(i)
                tempValues(i) = tempValues(j)
                tempValues(j) = currentValue
            End If
        Next j
    Next i

    Do While target.count > 0
        target.Remove 1
    Loop

    inserted = False
    currentValue = vbNullString

    For i = LBound(tempValues) To UBound(tempValues)
        If Not inserted Or StrComp(currentValue, tempValues(i), vbTextCompare) <> 0 Then
            target.Add tempValues(i)
            inserted = True
        End If
        currentValue = tempValues(i)
    Next i

    Exit Sub

ErrorHandler:
    target.Add itemValue
End Sub

Private Function TableExists(ByVal tableName As String) As Boolean
    On Error GoTo SafeExit

    Dim tdf As DAO.tableDef

    For Each tdf In CurrentDb.TableDefs
        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next tdf

SafeExit:
End Function

Private Function HasSelectedTargetObject() As Boolean
    HasSelectedTargetObject = (LenB(m_TargetObjectType) > 0 And LenB(m_TargetObjectName) > 0)
End Function

Private Sub UpdateLocalCommandButtons()
    On Error GoTo SafeExit

    If ControlExists("cmdSave") Then
        Me!cmdSave.Visible = False
    End If

    If ControlExists("cmdClose") Then
        Me!cmdClose.Visible = False
    End If

SafeExit:
End Sub

Private Function ControlExists(ByVal controlName As String) As Boolean
    On Error GoTo SafeExit

    Dim ctl As Control

    For Each ctl In Me.Controls
        If StrComp(ctl.Name, controlName, vbTextCompare) = 0 Then
            ControlExists = True
            Exit Function
        End If
    Next ctl

SafeExit:
End Function

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

Private Function HasControlOnForm(ByVal formInstance As Access.Form, ByVal controlName As String) As Boolean
    On Error GoTo SafeExit

    Dim ctl As Control

    If formInstance Is Nothing Then
        Exit Function
    End If

    For Each ctl In formInstance.Controls
        If StrComp(ctl.Name, controlName, vbTextCompare) = 0 Then
            HasControlOnForm = True
            Exit Function
        End If
    Next ctl

SafeExit:
End Function

Private Sub RefreshShellStatusIfHosted()
    On Error GoTo SafeExit

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()
    If shellForm Is Nothing Then
        Exit Sub
    End If

    modAppShell.UpdateShellCommandBarState shellForm

SafeExit:
End Sub

Private Sub NavigateHome()
    On Error GoTo ErrorHandler

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()
    If shellForm Is Nothing Then
        Exit Sub
    End If

    If Not modAppWorkspaceService.OpenWorkspaceForm(shellForm, "frmAppDashboard") Then
        Err.Raise vbObjectError + 5470, MODULE_NAME & ".NavigateHome", "Workspace navigation to dashboard failed."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NavigateHome", Err
End Sub

Private Sub NavigateBack()
    On Error GoTo ErrorHandler

    Dim shellForm As Access.Form

    Set shellForm = ResolveHostingShellForm()
    If shellForm Is Nothing Then
        Exit Sub
    End If

    If Not modAppWorkspaceService.GoBack(shellForm) Then
        Err.Raise vbObjectError + 5471, MODULE_NAME & ".NavigateBack", "Workspace back navigation failed."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NavigateBack", Err
End Sub

