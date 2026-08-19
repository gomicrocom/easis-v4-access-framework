Option Compare Database
Option Explicit

Private Const TABLE_TAG_HELP As String = "fw_tag_help"
Private Const FIELD_TOKEN_KEY As String = "token_key"
Private Const FIELD_CATEGORY As String = "category"
Private Const FIELD_SYNTAX_TEXT As String = "syntax_text"
Private Const FIELD_DESCRIPTION_TEXT As String = "description_text"
Private Const FIELD_EXAMPLE_TEXT As String = "example_text"
Private Const FIELD_NOTES_TEXT As String = "notes_text"
Private Const FIELD_SORT_ORDER As String = "sort_order"
Private Const FIELD_IS_ACTIVE As String = "is_active"

Private Sub Form_Load()
    LoadHelpList
End Sub

Private Sub LoadHelpList(Optional ByVal filterText As String = vbNullString)
    Dim sqlStatement As String

    sqlStatement = "SELECT " & FIELD_TOKEN_KEY & ", " & FIELD_SYNTAX_TEXT & " " & _
                   "FROM " & TABLE_TAG_HELP & " " & _
                   "WHERE " & FIELD_IS_ACTIVE & " = True "

    If Len(Trim$(filterText)) > 0 Then
        sqlStatement = sqlStatement & _
                       "AND (" & FIELD_TOKEN_KEY & " LIKE '*" & Replace(filterText, "'", "''") & "*' " & _
                       "OR " & FIELD_SYNTAX_TEXT & " LIKE '*" & Replace(filterText, "'", "''") & "*' " & _
                       "OR " & FIELD_CATEGORY & " LIKE '*" & Replace(filterText, "'", "''") & "*') "
    End If

    sqlStatement = sqlStatement & "ORDER BY " & FIELD_SORT_ORDER & ", " & FIELD_TOKEN_KEY & ";"

    Me.lstHelpTokens.RowSource = sqlStatement
    Me.lstHelpTokens.Requery
End Sub


Private Sub lstHelpTokens_AfterUpdate()

    LoadHelpDetail Nz(Me.lstHelpTokens.Column(0), vbNullString)

End Sub

Private Sub LoadHelpDetail(ByVal TokenKey As String)
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlStatement As String

    ClearDetailFields

    If Len(Trim$(TokenKey)) = 0 Then Exit Sub

    sqlStatement = "SELECT * FROM " & TABLE_TAG_HELP & " " & _
                   "WHERE " & FIELD_IS_ACTIVE & " = True " & _
                   "AND " & FIELD_TOKEN_KEY & " = '" & Replace(TokenKey, "'", "''") & "';"

    Set rs = CurrentDb.OpenRecordset(sqlStatement, dbOpenSnapshot)

    If Not rs.EOF Then
        Me.txtTokenKey = Nz(rs.Fields(FIELD_TOKEN_KEY).Value, "")
        Me.txtCategory = Nz(rs.Fields(FIELD_CATEGORY).Value, "")
        Me.txtSyntaxText = Nz(rs.Fields(FIELD_SYNTAX_TEXT).Value, "")
        Me.txtDescriptionText = Nz(rs.Fields(FIELD_DESCRIPTION_TEXT).Value, "")
        Me.txtExampleText = Nz(rs.Fields(FIELD_EXAMPLE_TEXT).Value, "")
        Me.txtNotesText = Nz(rs.Fields(FIELD_NOTES_TEXT).Value, "")
    End If

    rs.Close
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    MsgBox "Fehler beim Laden der Hilfe: " & Err.description, vbExclamation
End Sub

Private Sub ClearDetailFields()
    Me.txtTokenKey = Null
    Me.txtCategory = Null
    Me.txtSyntaxText = Null
    Me.txtDescriptionText = Null
    Me.txtExampleText = Null
    Me.txtNotesText = Null
End Sub

Private Sub txtFilter_AfterUpdate()

    LoadHelpList Nz(Me.txtFilter, vbNullString)

End Sub

Private Sub cmdClose_Click()

    DoCmd.Close acForm, Me.Name

End Sub








