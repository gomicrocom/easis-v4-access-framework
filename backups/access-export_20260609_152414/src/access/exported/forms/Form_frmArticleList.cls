Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmArticleList
' Purpose   : Workspace-aware list form for simple Article master data.
' Author    : Codex
' Version   : 0.1.2
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmArticleList"
Private m_currentSearchText As String

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    modFormRuntime.InitializeForm Me
    modFwTranslationRuntime.ApplyTranslations Me

    Me.RecordSource = modArticleService.BuildArticleListRowSource()
    Me.Filter = vbNullString
    Me.FilterOn = False
    m_currentSearchText = vbNullString

    UpdateActionState

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "Article list loaded."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_LIST_LOAD_ERROR", "Fehler beim Laden der Artikelliste: " & Err.Description), vbExclamation, MODULE_NAME
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

    OpenNewArticle
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdNew_Click", Err
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo ErrorHandler

    OpenCurrentArticle
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdEdit_Click", Err
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo ErrorHandler

    RefreshListData
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdRefresh_Click", Err
End Sub

Public Function GetCurrentWorkspaceRecordId() As Long
    GetCurrentWorkspaceRecordId = GetCurrentArticleId()
End Function

Public Function SupportsListCommandBar() As Boolean
    SupportsListCommandBar = True
End Function

Public Function SupportsListNew() As Boolean
    SupportsListNew = True
End Function

Public Function SupportsListEdit() As Boolean
    SupportsListEdit = True
End Function

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

    OpenNewArticle
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ListNew", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ListEdit()
    On Error GoTo ErrorHandler

    OpenCurrentArticle
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ListEdit", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ListRefresh()
    On Error GoTo ErrorHandler

    RefreshListData
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ListRefresh", Err
End Sub

Public Function GetWorkspaceState() As String
    Dim stateText As String

    stateText = AddWorkspaceStatePair(stateText, "search", GetSearchTextValue())
    stateText = AddWorkspaceStatePair(stateText, "article_id", CStr(GetCurrentArticleId()))

    GetWorkspaceState = stateText
End Function

Public Sub RestoreWorkspaceState(ByVal stateText As String)
    On Error GoTo ErrorHandler

    Dim articleId As Long
    Dim searchText As String

    searchText = GetWorkspaceStateValue(stateText, "search")
    articleId = modDaoHelper.NzLong(GetWorkspaceStateValue(stateText, "article_id"), 0)

    m_currentSearchText = Trim$(modDaoHelper.NzString(searchText))
    ApplySearchFilter

    If articleId > 0 Then
        RestoreCurrentArticleSelection articleId
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
        Me.Filter = modArticleService.BuildArticleSearchFilter(m_currentSearchText)
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

Private Sub OpenNewArticle()
    On Error GoTo ErrorHandler

    If Not modAppWorkspaceService.OpenWorkspaceForm(Me, modArticleService.FORM_ARTICLE_DETAIL, vbNullString, acFormAdd) Then
        Err.Raise vbObjectError + 2201, MODULE_NAME & ".OpenNewArticle", "Workspace navigation failed."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "OpenNewArticle", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub OpenCurrentArticle()
    On Error GoTo ErrorHandler

    Dim articleId As Long

    articleId = GetCurrentArticleId()
    If articleId <= 0 Then
        Exit Sub
    End If

    If Not modAppWorkspaceService.OpenWorkspaceForm(Me, modArticleService.FORM_ARTICLE_DETAIL, "[article_id]=" & CStr(articleId)) Then
        Err.Raise vbObjectError + 2200, MODULE_NAME & ".OpenCurrentArticle", "Workspace navigation failed."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "OpenCurrentArticle", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub RefreshListData()
    On Error GoTo ErrorHandler

    Me.Requery
    UpdateActionState

    modLoggingHandler.LogInfo MODULE_NAME & ".RefreshListData", _
        "Article list requeried."
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
        Me!cmdEdit.Enabled = (GetCurrentArticleId() > 0)
    End If

SafeExit:
End Sub

Private Function GetCurrentArticleId() As Long
    On Error GoTo SafeExit

    If Me.NewRecord Then
        Exit Function
    End If

    If HasControl("article_id") Then
        GetCurrentArticleId = modDaoHelper.NzLong(Me!article_id.Value, 0)
    ElseIf HasControl("txtArticleId") Then
        GetCurrentArticleId = modDaoHelper.NzLong(Me!txtArticleId.Value, 0)
    End If

SafeExit:
End Function

Private Sub RestoreCurrentArticleSelection(ByVal articleId As Long)
    On Error GoTo ErrorHandler

    Dim recordsetClone As DAO.Recordset

    Set recordsetClone = Me.RecordsetClone
    recordsetClone.FindFirst "[article_id]=" & CStr(articleId)

    If Not recordsetClone.NoMatch Then
        Me.Bookmark = recordsetClone.Bookmark
    End If

CleanExit:
    On Error Resume Next
    If Not recordsetClone Is Nothing Then recordsetClone.Close
    Set recordsetClone = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RestoreCurrentArticleSelection", Err
    Resume CleanExit
End Sub

Private Function GetSearchTextValue() As String
    GetSearchTextValue = m_currentSearchText
End Function

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
