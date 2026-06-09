VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmArticleGroupList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmArticleGroupList
' Purpose   : Workspace-aware list form for tenant Article Groups.
' Author    : Codex
' Version   : 0.1.1
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmArticleGroupList"
Private m_currentSearchText As String

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    modArticleGroupService.EnsureArticleGroupTable
    modFormRuntime.InitializeForm Me
    modFwTranslationRuntime.ApplyTranslations Me

    Me.RecordSource = modArticleGroupService.BuildArticleGroupListRowSource()
    Me.Filter = vbNullString
    Me.FilterOn = False
    m_currentSearchText = vbNullString

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "Article Group list loaded."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_GROUP_LIST_LOAD_ERROR", "Fehler beim Laden der Artikelgruppenliste: " & Err.description), vbExclamation, MODULE_NAME
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    UpdateActionState
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub cmdNew_Click()
    On Error GoTo ErrorHandler

    OpenNewArticleGroup
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdNew_Click", Err
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo ErrorHandler

    OpenCurrentArticleGroup
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdEdit_Click", Err
End Sub

Private Sub OpenNewArticleGroup()
    On Error GoTo ErrorHandler

    If Not modAppWorkspaceService.OpenWorkspaceForm(Me, modArticleGroupService.FORM_ARTICLE_GROUP_DETAIL, vbNullString, acFormAdd) Then
        Err.Raise vbObjectError + 2101, MODULE_NAME & ".OpenNewArticleGroup", "Workspace navigation failed."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "OpenNewArticleGroup", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo ErrorHandler

    RefreshListData
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdRefresh_Click", Err
End Sub

Public Sub ListSearch(ByVal searchText As String)
    On Error GoTo ErrorHandler

    m_currentSearchText = Trim$(modDaoHelper.NzString(searchText))
    ApplySearchFilter
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ListSearch", Err
End Sub

Public Sub ListClearSearch()
    On Error GoTo ErrorHandler

    ClearSearchFilter
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ListClearSearch", Err
End Sub

Public Sub ListNew()
    On Error GoTo ErrorHandler

    OpenNewArticleGroup
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ListNew", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Public Sub ListEdit()
    On Error GoTo ErrorHandler

    OpenCurrentArticleGroup
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ListEdit", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Public Sub ListRefresh()
    On Error GoTo ErrorHandler

    RefreshListData
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ListRefresh", Err
End Sub

Public Function SupportsListCommandBar() As Boolean
    SupportsListCommandBar = True
End Function

Public Function SupportsListNew() As Boolean
    SupportsListNew = True
End Function

Public Function SupportsListEdit() As Boolean
    SupportsListEdit = True
End Function

Public Function GetCurrentWorkspaceRecordId() As Long
    GetCurrentWorkspaceRecordId = GetCurrentArticleGroupId()
End Function

Public Function GetWorkspaceState() As String
    Dim stateText As String

    stateText = AddWorkspaceStatePair(stateText, "search", GetSearchTextValue())
    stateText = AddWorkspaceStatePair(stateText, "product_group_id", CStr(GetCurrentArticleGroupId()))

    GetWorkspaceState = stateText
End Function

Public Sub RestoreWorkspaceState(ByVal stateText As String)
    On Error GoTo ErrorHandler

    Dim articleGroupId As Long
    Dim searchText As String

    searchText = GetWorkspaceStateValue(stateText, "search")
    articleGroupId = modDaoHelper.NzLong(GetWorkspaceStateValue(stateText, "product_group_id"), 0)

    m_currentSearchText = Trim$(modDaoHelper.NzString(searchText))
    ApplySearchFilter

    If articleGroupId > 0 Then
        RestoreCurrentArticleGroupSelection articleGroupId
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreWorkspaceState", Err
End Sub

Private Sub ApplySearchFilter()
    On Error GoTo ErrorHandler

    If HasControl("txtFocusSink") Then
        Me!txtFocusSink.SetFocus
    End If

    Me.FilterOn = False
    If LenB(m_currentSearchText) = 0 Then
        Me.Filter = vbNullString
    Else
        Me.Filter = modArticleGroupService.BuildArticleGroupSearchFilter(m_currentSearchText)
        Me.FilterOn = True
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplySearchFilter", Err
End Sub

Private Sub ClearSearchFilter()
    On Error GoTo ErrorHandler

    m_currentSearchText = vbNullString
    Me.Filter = vbNullString
    Me.FilterOn = False
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ClearSearchFilter", Err
End Sub

Private Sub OpenCurrentArticleGroup()
    On Error GoTo ErrorHandler

    Dim articleGroupId As Long

    articleGroupId = GetCurrentArticleGroupId()
    If articleGroupId <= 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_GROUP_SELECT_FIRST", "Bitte zuerst eine Artikelgruppe auswaehlen."), vbExclamation, MODULE_NAME
        Exit Sub
    End If

    If Not modAppWorkspaceService.OpenWorkspaceForm(Me, modArticleGroupService.FORM_ARTICLE_GROUP_DETAIL, "[product_group_id]=" & CStr(articleGroupId)) Then
        Err.Raise vbObjectError + 2100, MODULE_NAME & ".OpenCurrentArticleGroup", "Workspace navigation failed."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "OpenCurrentArticleGroup", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub RefreshListData()
    On Error GoTo ErrorHandler

    Me.Requery
    UpdateActionState

    modLoggingHandler.LogInfo MODULE_NAME & ".RefreshListData", _
        "Article Group list requeried."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RefreshListData", Err
End Sub

Private Sub UpdateActionState()
    On Error GoTo SafeExit

    If HasControl("cmdNew") Then
        Me!cmdNew.Enabled = True
    End If

    If HasControl("cmdEdit") Then
        Me!cmdEdit.Enabled = (GetCurrentArticleGroupId() > 0)
    End If

SafeExit:
End Sub

Private Function GetCurrentArticleGroupId() As Long
    On Error GoTo SafeExit

    If Me.NewRecord Then
        Exit Function
    End If

    If HasControl("product_group_id") Then
        GetCurrentArticleGroupId = modDaoHelper.NzLong(Me!product_group_id.Value, 0)
    ElseIf HasControl("txtArticleGroupId") Then
        GetCurrentArticleGroupId = modDaoHelper.NzLong(Me!txtArticleGroupId.Value, 0)
    End If

SafeExit:
End Function

Private Sub RestoreCurrentArticleGroupSelection(ByVal articleGroupId As Long)
    On Error GoTo ErrorHandler

    Dim recordsetClone As DAO.Recordset

    Set recordsetClone = Me.recordsetClone
    recordsetClone.FindFirst "[product_group_id]=" & CStr(articleGroupId)

    If Not recordsetClone.NoMatch Then
        Me.Bookmark = recordsetClone.Bookmark
    End If

CleanExit:
    On Error Resume Next
    If Not recordsetClone Is Nothing Then recordsetClone.Close
    Set recordsetClone = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreCurrentArticleGroupSelection", Err
    Resume CleanExit
End Sub

Private Function GetSearchTextValue() As String
    GetSearchTextValue = m_currentSearchText
End Function

Private Function HasControl(ByVal ControlName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentControl As Control

    For Each currentControl In Me.Controls
        If StrComp(currentControl.Name, ControlName, vbTextCompare) = 0 Then
            HasControl = True
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


