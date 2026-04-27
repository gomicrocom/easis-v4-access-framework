VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTagHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
    LoadHelpList
End Sub

Private Sub LoadHelpList(Optional ByVal FilterText As String = vbNullString)
    Dim sqlText As String

    sqlText = "SELECT TokenKey, SyntaxText " & _
              "FROM tblFwTagHelp " & _
              "WHERE IsActive = True "

    If Len(Trim$(FilterText)) > 0 Then
        sqlText = sqlText & _
                  "AND (TokenKey LIKE '*" & Replace(FilterText, "'", "''") & "*' " & _
                  "OR SyntaxText LIKE '*" & Replace(FilterText, "'", "''") & "*' " & _
                  "OR Category LIKE '*" & Replace(FilterText, "'", "''") & "*') "
    End If

    sqlText = sqlText & "ORDER BY SortOrder, TokenKey;"

    Me.lstHelpTokens.RowSource = sqlText
    Me.lstHelpTokens.Requery
End Sub

Private Sub lstHelpTokens_AfterUpdate()
    LoadHelpDetail Nz(Me.lstHelpTokens.Column(0), vbNullString)
End Sub

Private Sub LoadHelpDetail(ByVal TokenKey As String)
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlText As String

    ClearDetailFields

    If Len(Trim$(TokenKey)) = 0 Then Exit Sub

    sqlText = "SELECT * FROM tblFwTagHelp " & _
              "WHERE IsActive = True " & _
              "AND TokenKey = '" & Replace(TokenKey, "'", "''") & "';"

    Set rs = CurrentDb.OpenRecordset(sqlText, dbOpenSnapshot)

    If Not rs.EOF Then
        Me.txtTokenKey = Nz(rs!TokenKey, "")
        Me.txtCategory = Nz(rs!Category, "")
        Me.txtSyntaxText = Nz(rs!SyntaxText, "")
        Me.txtDescriptionText = Nz(rs!DescriptionText, "")
        Me.txtExampleText = Nz(rs!ExampleText, "")
        Me.txtNotesText = Nz(rs!NotesText, "")
    End If

    rs.Close
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    MsgBox "Fehler beim Laden der Hilfe: " & Err.Description, vbExclamation
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

