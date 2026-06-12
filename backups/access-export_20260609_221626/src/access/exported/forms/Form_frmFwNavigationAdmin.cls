VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFwNavigationAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ExportKind: CODE_BEHIND_ONLY
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmFwNavigationAdmin
' Purpose   : Workspace-compatible administration form for fw_navigation.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmFwNavigationAdmin"
Private Const TABLE_NAVIGATION As String = "fw_navigation"
Private Const TYPE_GROUP As String = "GROUP"
Private Const TYPE_FORM As String = "FORM"
Private Const TYPE_REPORT As String = "REPORT"
Private Const TYPE_ACTION As String = "ACTION"
Private Const OPEN_MODE_NORMAL As String = "NORMAL"
Private Const OPEN_MODE_ADD As String = "ADD"
Private Const OPEN_MODE_EDIT As String = "EDIT"
Private Const OPEN_MODE_READONLY As String = "READONLY"
Private Const LIST_CONTROL As String = "lstNavigationEntries"
Private Const PARENT_COMBO As String = "cboParentNavigationId"
Private Const OBJECT_TYPE_COMBO As String = "cboObjectType"
Private Const OPEN_MODE_COMBO As String = "cboOpenMode"
Private Const OBJECT_NAME_COMBO As String = "cboObjectName"
Private Const TITLE_CONTROL As String = "txtFormTitle"
Private Const SHELL_NAVIGATION_HOST As String = "subNavigationHost"

Private m_pendingCreateAsGroup As Boolean

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    modFormRuntime.InitializeForm Me
    modFwTranslationRuntime.ApplyTranslations Me

    Me.RecordSource = TABLE_NAVIGATION
    Me.AllowAdditions = True
    Me.AllowEdits = True
    Me.AllowDeletions = False

    ConfigureCombos
    RefreshNavigationAdminList
    UpdateFormTitle
    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox "Fehler beim Laden der Navigationsverwaltung: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    SyncListSelectionToCurrentRecord
    UpdateFormTitle
    UpdateActionState
    UpdateFieldVisibility
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo ErrorHandler

    ApplyDefaultValuesForNewRecord m_pendingCreateAsGroup
    m_pendingCreateAsGroup = False
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_BeforeInsert", Err
    Cancel = True
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo ErrorHandler

    If Not ValidateCurrentEntry(Cancel) Then
        Exit Sub
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_BeforeUpdate", Err
    Cancel = True
End Sub

Private Sub Form_AfterUpdate()
    On Error GoTo ErrorHandler

    RefreshNavigationAdminList
    RefreshShellNavigationHost

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_AfterUpdate", _
        "Navigation entry saved for navigation_id=" & CStr(CurrentNavigationId()) & "."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_AfterUpdate", Err
End Sub

Private Sub Form_AfterInsert()
    On Error GoTo ErrorHandler

    RefreshNavigationAdminList
    RefreshShellNavigationHost
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_AfterInsert", Err
End Sub

Private Sub lstNavigationEntries_AfterUpdate()
    On Error GoTo ErrorHandler

    NavigateToSelectedEntry
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "lstNavigationEntries_AfterUpdate", Err
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler

    SaveCurrentEntry
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdSave_Click", Err
    MsgBox "Fehler beim Speichern des Navigationseintrags: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo ErrorHandler

    If Me.Dirty Then
        Me.Dirty = False
    End If

    RefreshNavigationAdminList
    RefreshShellNavigationHost
    Me.Requery
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdRefresh_Click", Err
    MsgBox "Fehler beim Aktualisieren der Navigationsverwaltung: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdNewGroup_Click()
    On Error GoTo ErrorHandler

    CreateNewEntry True
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdNewGroup_Click", Err
    MsgBox "Fehler beim Erstellen der neuen Gruppe: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdNewItem_Click()
    On Error GoTo ErrorHandler

    CreateNewEntry False
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdNewItem_Click", Err
    MsgBox "Fehler beim Erstellen des neuen Eintrags: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdDeactivate_Click()
    On Error GoTo ErrorHandler

    SetCurrentBooleanField "is_active", False
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdDeactivate_Click", Err
    MsgBox "Fehler beim Deaktivieren des Navigationseintrags: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdHide_Click()
    On Error GoTo ErrorHandler

    SetCurrentBooleanField "is_visible", False
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdHide_Click", Err
    MsgBox "Fehler beim Ausblenden des Navigationseintrags: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdShow_Click()
    On Error GoTo ErrorHandler

    SetCurrentBooleanField "is_visible", True
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdShow_Click", Err
    MsgBox "Fehler beim Einblenden des Navigationseintrags: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdTestOpen_Click()
    On Error GoTo ErrorHandler

    Dim shellForm As Access.Form

    If Me.NewRecord Or CurrentNavigationId() <= 0 Then
        MsgBox "Bitte zuerst einen gespeicherten Navigationseintrag auswaehlen.", vbExclamation, MODULE_NAME
        Exit Sub
    End If

    If Me.Dirty Then
        Me.Dirty = False
    End If

    Set shellForm = ResolveShellForm()
    If shellForm Is Nothing Then
        MsgBox "Testoeffnen ist nur innerhalb der Shell verfuegbar.", vbInformation, MODULE_NAME
        Exit Sub
    End If

    modAppNavigationService.HandleNavigationClick shellForm, CurrentNavigationId()
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdTestOpen_Click", Err
    MsgBox "Fehler beim Testoeffnen des Navigationseintrags: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cboObjectType_AfterUpdate()
    On Error GoTo ErrorHandler

    If UCase$(Trim$(Nz(Me!object_type.Value, vbNullString))) = TYPE_GROUP Then
        Me!parent_navigation_id = Null
        Me!object_name = Null
        Me!open_mode = OPEN_MODE_NORMAL
    ElseIf LenB(Trim$(Nz(Me!open_mode.Value, vbNullString))) = 0 Then
        Me!open_mode = OPEN_MODE_NORMAL
    End If

    UpdateActionState
    ConfigureObjectNameCombo
    UpdateFieldVisibility
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboObjectType_AfterUpdate", Err
End Sub

Private Sub cboParentNavigationId_AfterUpdate()
    On Error GoTo ErrorHandler

    If Not IsNull(Me!parent_navigation_id.Value) Then
        Me!navigation_group = ResolveParentNavigationGroup(Me!parent_navigation_id.Value)
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboParentNavigationId_AfterUpdate", Err
End Sub

Private Sub ConfigureCombos()
    ConfigureObjectTypeCombo
    ConfigureOpenModeCombo
    ConfigureParentCombo
    ConfigureObjectNameCombo
End Sub

Private Sub ConfigureObjectTypeCombo()
    If Not HasControl(OBJECT_TYPE_COMBO) Then
        Exit Sub
    End If

    With Me.Controls(OBJECT_TYPE_COMBO)
        .RowSourceType = "Value List"
        .rowSource = TYPE_GROUP & ";" & TYPE_FORM & ";" & TYPE_REPORT & ";" & TYPE_ACTION
        .LimitToList = True
    End With
End Sub

Private Sub ConfigureOpenModeCombo()
    If Not HasControl(OPEN_MODE_COMBO) Then
        Exit Sub
    End If

    With Me.Controls(OPEN_MODE_COMBO)
        .RowSourceType = "Value List"
        .rowSource = OPEN_MODE_NORMAL & ";" & OPEN_MODE_ADD & ";" & OPEN_MODE_EDIT & ";" & OPEN_MODE_READONLY
        .LimitToList = True
    End With
End Sub

Private Sub ConfigureParentCombo()
    Dim sqlStatement As String

    If Not HasControl(PARENT_COMBO) Then
        Exit Sub
    End If

    sqlStatement = ""
    sqlStatement = sqlStatement & "SELECT navigation_id, "
    sqlStatement = sqlStatement & "ResolveShellText(Nz(caption_key,''), Nz(fallback_caption,'')) AS display_caption "
    sqlStatement = sqlStatement & "FROM fw_navigation "
    sqlStatement = sqlStatement & "WHERE UCase(Nz(object_type,'')) = 'GROUP' "
    sqlStatement = sqlStatement & "ORDER BY sort_order, navigation_id;"

    With Me.Controls(PARENT_COMBO)
        .RowSourceType = "Table/Query"
        .rowSource = sqlStatement
        .BoundColumn = 1
        .columnCount = 2
        .columnWidths = "0cm;5cm"
        .LimitToList = True
    End With
End Sub

Private Sub RefreshNavigationAdminList()
    On Error GoTo ErrorHandler

    If HasControl(LIST_CONTROL) Then
        With Me.Controls(LIST_CONTROL)
            .RowSourceType = "Table/Query"
            .rowSource = BuildNavigationListRowSource()
            .BoundColumn = 1
            .columnCount = 2
            .columnWidths = "0cm;8cm"
            .Requery
        End With
    End If

    ConfigureParentCombo
    SyncListSelectionToCurrentRecord
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RefreshNavigationAdminList", Err
End Sub

Private Function BuildNavigationListRowSource() As String
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "SELECT "
    sqlStatement = sqlStatement & "n.navigation_id, "
    sqlStatement = sqlStatement & "IIf(n.parent_navigation_id Is Null, "
    sqlStatement = sqlStatement & "ResolveShellText(Nz(n.caption_key,''), Nz(n.fallback_caption,'')), "
    sqlStatement = sqlStatement & SqlTextNav("    - ") & " & ResolveShellText(Nz(n.caption_key,''), Nz(n.fallback_caption,''))) "
    sqlStatement = sqlStatement & "& IIf(Nz(n.is_active,True)=False, " & SqlText(" [inaktiv]") & ", '') "
    sqlStatement = sqlStatement & "& IIf(Nz(n.is_visible,True)=False, " & SqlText(" [ausgeblendet]") & ", '') "
    sqlStatement = sqlStatement & "AS display_text "
    sqlStatement = sqlStatement & "FROM fw_navigation AS n "
    sqlStatement = sqlStatement & "LEFT JOIN fw_navigation AS p "
    sqlStatement = sqlStatement & "ON n.parent_navigation_id = p.navigation_id "
    sqlStatement = sqlStatement & "ORDER BY "
    sqlStatement = sqlStatement & "IIf(n.parent_navigation_id Is Null, n.sort_order, p.sort_order), "
    sqlStatement = sqlStatement & "IIf(n.parent_navigation_id Is Null, 0, 1), "
    sqlStatement = sqlStatement & "n.sort_order, "
    sqlStatement = sqlStatement & "n.navigation_id;"

    BuildNavigationListRowSource = sqlStatement
End Function

Private Sub NavigateToSelectedEntry()
    On Error GoTo ErrorHandler

    Dim selectedNavigationId As Long

    If Not HasControl(LIST_CONTROL) Then
        Exit Sub
    End If

    selectedNavigationId = modDaoHelper.NzLong(Me.Controls(LIST_CONTROL).Value, 0)
    If selectedNavigationId <= 0 Then
        Exit Sub
    End If

    Me.recordsetClone.FindFirst "navigation_id = " & CStr(selectedNavigationId)
    If Not Me.recordsetClone.NoMatch Then
        Me.Bookmark = Me.recordsetClone.Bookmark
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NavigateToSelectedEntry", Err
End Sub

Private Sub SyncListSelectionToCurrentRecord()
    On Error GoTo SafeExit

    If HasControl(LIST_CONTROL) Then
        Me.Controls(LIST_CONTROL).Value = CurrentNavigationId()
    End If

SafeExit:
End Sub

Private Sub SaveCurrentEntry()
    If Me.Dirty Then
        Me.Dirty = False
    End If
End Sub

Private Sub CreateNewEntry(ByVal createGroup As Boolean)
    On Error GoTo ErrorHandler

    If Me.Dirty Then
        Me.Dirty = False
    End If

    m_pendingCreateAsGroup = createGroup
    DoCmd.GoToRecord , , acNewRec
    ApplyDefaultValuesForNewRecord createGroup
    UpdateFormTitle
    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CreateNewEntry", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub ApplyDefaultValuesForNewRecord(ByVal createGroup As Boolean)
    Dim defaultParentId As Variant

    Me!is_active = True
    Me!is_visible = True
    Me!is_expanded = False
    Me!sort_order = ResolveDefaultSortOrder()

    If createGroup Then
        Me!object_type = TYPE_GROUP
        Me!open_mode = OPEN_MODE_NORMAL
        Me!parent_navigation_id = Null
        Me!object_name = Null
        Me!navigation_group = "Neue Gruppe"
    Else
        defaultParentId = ResolveDefaultParentNavigationId()
        Me!object_type = TYPE_FORM
        Me!open_mode = OPEN_MODE_NORMAL
        Me!parent_navigation_id = defaultParentId
        Me!navigation_group = ResolveParentNavigationGroup(defaultParentId)
    End If
End Sub

Private Function ValidateCurrentEntry(ByRef Cancel As Integer) As Boolean
    Dim navigationGroup As String
    Dim captionKey As String
    Dim fallbackCaption As String
    Dim objectType As String
    Dim objectName As String
    Dim openMode As String
    Dim parentNavigationId As Long

    navigationGroup = Trim$(Nz(Me!navigation_group.Value, vbNullString))
    captionKey = UCase$(Trim$(Nz(Me!caption_key.Value, vbNullString)))
    fallbackCaption = Trim$(Nz(Me!fallback_caption.Value, vbNullString))
    objectType = UCase$(Trim$(Nz(Me!object_type.Value, vbNullString)))
    objectName = Trim$(Nz(Me!object_name.Value, vbNullString))
    openMode = UCase$(Trim$(Nz(Me!open_mode.Value, vbNullString)))
    parentNavigationId = modDaoHelper.NzLong(Me!parent_navigation_id.Value, 0)

    If LenB(navigationGroup) = 0 Then
        MsgBox "navigation_group ist erforderlich.", vbExclamation, MODULE_NAME
        Cancel = True
        Exit Function
    End If

    If LenB(captionKey) = 0 Then
        MsgBox "caption_key ist erforderlich.", vbExclamation, MODULE_NAME
        Cancel = True
        Exit Function
    End If

    If LenB(fallbackCaption) = 0 Then
        MsgBox "fallback_caption ist erforderlich.", vbExclamation, MODULE_NAME
        Cancel = True
        Exit Function
    End If

    Select Case objectType
        Case TYPE_GROUP, TYPE_FORM, TYPE_REPORT, TYPE_ACTION
        Case Else
            MsgBox "object_type ist ungueltig.", vbExclamation, MODULE_NAME
            Cancel = True
            Exit Function
    End Select

    ' caption_key validation for navigation entries:
    ' GROUP entries must use NAV.GROUP.*
    ' all other navigation entries must use NAV.*
    If objectType = TYPE_GROUP Then
        If Left$(captionKey, 10) <> "NAV.GROUP." Then
            MsgBox "caption_key fuer Gruppen muss mit NAV.GROUP. beginnen.", vbExclamation, MODULE_NAME
            Cancel = True
            Exit Function
        End If
    Else
        If Left$(captionKey, 4) <> "NAV." Then
            MsgBox "caption_key fuer Navigationseintraege muss mit NAV. beginnen.", vbExclamation, MODULE_NAME
            Cancel = True
            Exit Function
        End If

        If Left$(captionKey, 10) = "NAV.GROUP." Then
            MsgBox "caption_key NAV.GROUP.* ist nur fuer Gruppen erlaubt.", vbExclamation, MODULE_NAME
            Cancel = True
            Exit Function
        End If
    End If

    ' Persist normalized uppercase key.
    Me!caption_key = captionKey

    Select Case openMode
        Case OPEN_MODE_NORMAL, OPEN_MODE_ADD, OPEN_MODE_EDIT, OPEN_MODE_READONLY
        Case Else
            MsgBox "open_mode ist ungueltig.", vbExclamation, MODULE_NAME
            Cancel = True
            Exit Function
    End Select

    If objectType = TYPE_GROUP Then
        Me!parent_navigation_id = Null
        Me!object_name = Null
        If openMode <> OPEN_MODE_NORMAL Then
            Me!open_mode = OPEN_MODE_NORMAL
        End If
    Else
        If LenB(objectName) = 0 And (objectType = TYPE_FORM Or objectType = TYPE_REPORT) Then
            MsgBox "object_name ist fuer FORM und REPORT erforderlich.", vbExclamation, MODULE_NAME
            Cancel = True
            Exit Function
        End If

        If parentNavigationId > 0 Then
            If Not IsGroupEntry(parentNavigationId) Then
                MsgBox "parent_navigation_id muss auf eine GROUP zeigen.", vbExclamation, MODULE_NAME
                Cancel = True
                Exit Function
            End If
        End If

        If openMode = OPEN_MODE_ADD And objectType <> TYPE_FORM Then
            MsgBox "open_mode=ADD ist nur fuer FORM erlaubt.", vbExclamation, MODULE_NAME
            Cancel = True
            Exit Function
        End If
    End If

    ValidateCurrentEntry = True
End Function
Private Function IsGroupEntry(ByVal navigationId As Long) As Boolean
    IsGroupEntry = (StrComp( _
        Trim$(Nz(DLookup("object_type", TABLE_NAVIGATION, "navigation_id = " & CStr(navigationId)), vbNullString)), _
        TYPE_GROUP, vbTextCompare) = 0)
End Function

Private Function ResolveDefaultParentNavigationId() As Variant
    Dim currentNavigationIdValue As Long
    Dim currentObjectType As String
    Dim parentNavigationIdValue As Variant

    If Me.NewRecord Then
        ResolveDefaultParentNavigationId = Null
        Exit Function
    End If

    currentNavigationIdValue = CurrentNavigationId()
    currentObjectType = UCase$(Trim$(Nz(Me.Controls("object_type").Value, vbNullString)))

    If currentObjectType = TYPE_GROUP Then
        ResolveDefaultParentNavigationId = currentNavigationIdValue
    Else
        parentNavigationIdValue = Me.Controls("parent_navigation_id").Value

        If IsNull(parentNavigationIdValue) Or IsEmpty(parentNavigationIdValue) Then
            ResolveDefaultParentNavigationId = Null
        Else
            ResolveDefaultParentNavigationId = parentNavigationIdValue
        End If
    End If
End Function

Private Function ResolveParentNavigationGroup(ByVal parentNavigationId As Variant) As String
    If IsNull(parentNavigationId) Or IsEmpty(parentNavigationId) Then
        ResolveParentNavigationGroup = vbNullString
    Else
        ResolveParentNavigationGroup = Trim$(Nz( _
            DLookup("navigation_group", TABLE_NAVIGATION, "navigation_id = " & CStr(CLng(parentNavigationId))), _
            vbNullString))
    End If
End Function

Private Function ResolveDefaultSortOrder() As Long
    Dim currentSortOrder As Long

    If Me.NewRecord Then
        ResolveDefaultSortOrder = 10
    Else
        currentSortOrder = modDaoHelper.NzLong(Me!sort_order.Value, 0)
        If currentSortOrder <= 0 Then
            ResolveDefaultSortOrder = 10
        Else
            ResolveDefaultSortOrder = currentSortOrder + 10
        End If
    End If
End Function

Private Sub SetCurrentBooleanField(ByVal fieldName As String, ByVal fieldValue As Boolean)
    On Error GoTo ErrorHandler

    If Me.NewRecord Or CurrentNavigationId() <= 0 Then
        MsgBox "Bitte zuerst einen gespeicherten Navigationseintrag auswaehlen.", vbExclamation, MODULE_NAME
        Exit Sub
    End If

    Me.Controls(fieldName).Value = fieldValue
    If Me.Dirty Then
        Me.Dirty = False
    End If

    RefreshNavigationAdminList
    RefreshShellNavigationHost
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SetCurrentBooleanField", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub RefreshShellNavigationHost()
    On Error GoTo SafeExit

    Dim shellForm As Access.Form

    Set shellForm = ResolveShellForm()
    If shellForm Is Nothing Then
        Exit Sub
    End If

    If HasFormControl(shellForm, SHELL_NAVIGATION_HOST) Then
        shellForm.Controls(SHELL_NAVIGATION_HOST).Form.Requery
    End If

SafeExit:
End Sub

Private Sub UpdateFormTitle()
    On Error GoTo SafeExit

    If HasControl(TITLE_CONTROL) Then
        Me.Controls(TITLE_CONTROL).Value = ResolveCurrentTitleText()
    End If

SafeExit:
End Sub

Private Function ResolveCurrentTitleText() As String
    Dim fallbackCaption As String

    fallbackCaption = Trim$(Nz(Me!fallback_caption.Value, vbNullString))
    If Me.NewRecord Then
        ResolveCurrentTitleText = "Navigation verwalten"
    ElseIf LenB(fallbackCaption) > 0 Then
        ResolveCurrentTitleText = fallbackCaption
    Else
        ResolveCurrentTitleText = "Navigation verwalten"
    End If
End Function

Private Sub UpdateActionState()
    On Error GoTo SafeExit

    SetButtonEnabledIfPresent "cmdDeactivate", (Not Me.NewRecord And CurrentNavigationId() > 0)
    SetButtonEnabledIfPresent "cmdHide", (Not Me.NewRecord And CurrentNavigationId() > 0)
    SetButtonEnabledIfPresent "cmdShow", (Not Me.NewRecord And CurrentNavigationId() > 0)
    SetButtonEnabledIfPresent "cmdTestOpen", (Not Me.NewRecord And CurrentNavigationId() > 0)

SafeExit:
End Sub

Private Sub SetButtonEnabledIfPresent(ByVal ControlName As String, ByVal isEnabled As Boolean)
    On Error GoTo SafeExit

    If HasControl(ControlName) Then
        Me.Controls(ControlName).Enabled = isEnabled
    End If

SafeExit:
End Sub

Private Function CurrentNavigationId() As Long
    On Error GoTo SafeExit

    CurrentNavigationId = modDaoHelper.NzLong(Me!navigation_id.Value, 0)

SafeExit:
End Function

Private Function ResolveShellForm() As Access.Form
    On Error GoTo SafeExit

    If Not (Me.Parent Is Nothing) Then
        Set ResolveShellForm = Me.Parent
    End If

SafeExit:
End Function

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


Private Sub ConfigureObjectNameCombo()
    On Error GoTo ErrorHandler

    Dim objectType As String
    Dim rowSource As String

    If Not HasControl(OBJECT_NAME_COMBO) Then
        Exit Sub
    End If

    objectType = UCase$(Trim$(Nz(Me.Controls(OBJECT_TYPE_COMBO).Value, vbNullString)))

    Select Case objectType
        Case TYPE_FORM
            rowSource = BuildAccessObjectValueList(acForm)

        Case TYPE_REPORT
            rowSource = BuildAccessObjectValueList(acReport)

        Case Else
            rowSource = vbNullString
    End Select

    With Me.Controls(OBJECT_NAME_COMBO)
        .RowSourceType = "Value List"
        .rowSource = rowSource
        .LimitToList = False
    End With

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureObjectNameCombo", Err
End Sub
Private Function BuildAccessObjectValueList(ByVal objectType As AcObjectType) As String
    On Error GoTo ErrorHandler

    Dim accessObject As Access.accessObject
    Dim result As String

    Select Case objectType
        Case acForm
            For Each accessObject In CurrentProject.AllForms
                If LenB(result) > 0 Then result = result & ";"
                result = result & accessObject.Name
            Next accessObject

        Case acReport
            For Each accessObject In CurrentProject.AllReports
                If LenB(result) > 0 Then result = result & ";"
                result = result & accessObject.Name
            Next accessObject
    End Select

    BuildAccessObjectValueList = result
    Exit Function

ErrorHandler:
    BuildAccessObjectValueList = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "BuildAccessObjectValueList", Err
End Function
Private Sub UpdateFieldVisibility()
    On Error GoTo ErrorHandler

    Dim objectType As String
    Dim isGroup As Boolean
    Dim isForm As Boolean
    Dim isReport As Boolean
    Dim isAction As Boolean

    objectType = UCase$(Trim$(Nz(Me.Controls("object_type").Value, vbNullString)))

    isGroup = (objectType = TYPE_GROUP)
    isForm = (objectType = TYPE_FORM)
    isReport = (objectType = TYPE_REPORT)
    isAction = (objectType = TYPE_ACTION)
    
    If HasControl("txtFocusSink") Then
        Me.Controls("txtFocusSink").SetFocus
    End If

    SetControlVisibleIfPresent "parent_navigation_id", Not isGroup
    SetControlVisibleIfPresent "cboParentNavigationId", Not isGroup

    SetControlVisibleIfPresent "object_name", Not isGroup
    SetControlVisibleIfPresent "cboObjectName", Not isGroup

    SetControlVisibleIfPresent "open_mode", isForm
    SetControlVisibleIfPresent "cboOpenMode", isForm

    SetControlVisibleIfPresent "is_expanded", isGroup

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "UpdateFieldVisibility", Err
End Sub
Private Sub SetControlVisibleIfPresent(ByVal ControlName As String, ByVal isVisible As Boolean)
    On Error GoTo SafeExit

    If HasControl(ControlName) Then
        Me.Controls(ControlName).Visible = isVisible
    End If

SafeExit:
End Sub