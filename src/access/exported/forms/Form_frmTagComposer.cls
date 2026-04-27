VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTagComposer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_TargetForm As String
Private m_SessionId As String
Private m_CurrentControl As String

Public Property Get CurrentComposerControlName() As String
    CurrentComposerControlName = m_CurrentControl
End Property

Private Sub cmdHelp_Click()
    DoCmd.OpenForm "frmTagHelp", WindowMode:=acDialog
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    If Len(Nz(Me.openArgs, vbNullString)) = 0 Then
        MsgBox "Quellformular wurde nicht übergeben.", vbExclamation
        Exit Sub
    End If

    m_TargetForm = Me.openArgs
    m_SessionId = BuildSessionId()

    InitializeTempComposerData
    BindTempSubform

    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Initialisieren des Tag-Composer: " & Err.Description, vbExclamation
End Sub

Private Function BuildSessionId() As String
    BuildSessionId = "TC_" & Format(Now, "yyyymmddhhnnss") & "_" & CStr(Int((Rnd() * 90000) + 10000))
End Function

Private Sub InitializeTempComposerData()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim frm As Form
    Dim ctl As Control
    Dim RowNo As Long
    Dim controlTypeName As String

    Set db = CurrentDb
    db.Execute "DELETE FROM tblTmpTagComposer WHERE SessionId = '" & Replace(m_SessionId, "'", "''") & "'", dbFailOnError

    Set frm = Forms(m_TargetForm)
    Set rs = db.OpenRecordset("tblTmpTagComposer", dbOpenDynaset)

    RowNo = 0
    For Each ctl In frm.Controls
        RowNo = RowNo + 1
        controlTypeName = GetControlTypeName(ctl)

        rs.AddNew
        rs!SessionId = m_SessionId
        rs!RowNo = RowNo
        rs!ControlName = ctl.Name
        rs!ControlType = controlTypeName
        rs!TagValue = Nz(ctl.Tag, vbNullString)
        rs!IsDirty = False
        rs.Update
    Next ctl

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

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
        Case Else: GetControlTypeName = CStr(ctl.ControlType)
    End Select
    Exit Function

SafeExit:
    GetControlTypeName = vbNullString
End Function

Private Sub BindTempSubform()
    On Error GoTo ErrorHandler

    Dim sqlText As String

    sqlText = "SELECT SessionId, RowNo, ControlName, ControlType, TagValue, IsDirty " & _
              "FROM tblTmpTagComposer " & _
              "WHERE SessionId = '" & Replace(m_SessionId, "'", "''") & "' " & _
              "ORDER BY RowNo;"

    Me.sfrmTmpTagComposer.Form.RecordSource = sqlText
    Me.sfrmTmpTagComposer.Form.Requery

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub LoadSelectedControlFromSubform(ByVal ControlName As String, ByVal TagValue As String)
    On Error GoTo ErrorHandler

    Call SaveCurrentControlTagToTemp(False)
    
    m_CurrentControl = Nz(ControlName, vbNullString)

    ResetUI
    LoadFromTag Nz(TagValue, vbNullString)

    Exit Sub

ErrorHandler:
    MsgBox "Fehler beim Laden des Controls: " & Err.Description, vbExclamation
End Sub

Private Sub SaveCurrentControlTagToTemp(Optional ByVal RefreshSubform As Boolean = False)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim sqlText As String
    Dim newTag As String

    If Len(m_CurrentControl) = 0 Then Exit Sub

    newTag = BuildTagString()

    Set db = CurrentDb
    sqlText = "UPDATE tblTmpTagComposer " & _
              "SET TagValue = '" & Replace(newTag, "'", "''") & "', " & _
              "    IsDirty = True " & _
              "WHERE SessionId = '" & Replace(m_SessionId, "'", "''") & "' " & _
              "  AND ControlName = '" & Replace(m_CurrentControl, "'", "''") & "';"

    db.Execute sqlText, dbFailOnError

    Me.txtResult = newTag

    If RefreshSubform Then
        RefreshSubformCurrentRow m_CurrentControl
    End If

    Set db = Nothing
    Exit Sub

ErrorHandler:
    Set db = Nothing
    MsgBox "Fehler beim Zwischenspeichern des Tags: " & Err.Description, vbExclamation
End Sub
Private Sub cmdApply_Click()
    If Len(m_CurrentControl) = 0 Then
        MsgBox "Bitte zuerst ein Control in der Liste auswählen.", vbExclamation
        Exit Sub
    End If

    SaveCurrentControlTagToTemp True
    MsgBox "Tag für '" & m_CurrentControl & "' übernommen.", vbInformation
End Sub
Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim wasLoaded As Boolean
    Dim wasVisible As Boolean

    SaveCurrentControlTagToTemp

    wasLoaded = IsFormLoaded(m_TargetForm)
    wasVisible = False

    If wasLoaded Then
        wasVisible = Forms(m_TargetForm).Visible
        DoCmd.Close acForm, m_TargetForm, acSaveYes
    End If

    DoCmd.OpenForm m_TargetForm, acDesign, , , , acHidden

    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT ControlName, TagValue FROM tblTmpTagComposer " & _
        "WHERE SessionId = '" & Replace(m_SessionId, "'", "''") & "' " & _
        "ORDER BY RowNo;", dbOpenSnapshot)

    Do While Not rs.EOF
        On Error Resume Next
        Forms(m_TargetForm).Controls(Nz(rs!ControlName, vbNullString)).Tag = Nz(rs!TagValue, vbNullString)
        If Err.Number <> 0 Then
            Debug.Print "Tag write failed:", Nz(rs!ControlName, ""), Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    DoCmd.Save acForm, m_TargetForm
    DoCmd.Close acForm, m_TargetForm, acSaveYes

    If wasLoaded Then
        DoCmd.OpenForm m_TargetForm, acNormal
        If wasVisible = False Then
            Forms(m_TargetForm).Visible = False
        End If
    End If

    CleanupTempSession

    MsgBox "Tags wurden dauerhaft gespeichert.", vbInformation
    DoCmd.Close acForm, Me.Name
    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    MsgBox "Fehler beim Speichern der Tags: " & Err.Description, vbExclamation
End Sub

Private Sub cmdClose_Click()
    CleanupTempSession
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub CleanupTempSession()
    On Error Resume Next
    CurrentDb.Execute "DELETE FROM tblTmpTagComposer WHERE SessionId = '" & Replace(m_SessionId, "'", "''") & "';"
End Sub

Private Function IsFormLoaded(ByVal FormName As String) As Boolean
    On Error GoTo SafeExit
    IsFormLoaded = CurrentProject.AllForms(FormName).IsLoaded
SafeExit:
End Function

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

    ' Basis
    If Nz(Me.chkRequired, False) Then parts.Add "REQUIRED"
    If Nz(Me.chkNumeric, False) Then parts.Add "NUMERIC"
    If Nz(Me.chkInteger, False) Then parts.Add "INTEGER"
    If Nz(Me.chkDate, False) Then parts.Add "DATE"

    ' Numeric Range
    If Len(Nz(Me.txtMin, "")) > 0 Then parts.Add "MIN:" & Trim(Me.txtMin)
    If Len(Nz(Me.txtMax, "")) > 0 Then parts.Add "MAX:" & Trim(Me.txtMax)

    ' Text Length
    If Len(Nz(Me.txtMinLen, "")) > 0 Then parts.Add "MINLEN:" & Trim(Me.txtMinLen)
    If Len(Nz(Me.txtMaxLen, "")) > 0 Then parts.Add "MAXLEN:" & Trim(Me.txtMaxLen)

    ' Verhalten
    If Nz(Me.chkHidden, False) Then parts.Add "HIDDEN"
    If Nz(Me.chkSetFocus, False) Then parts.Add "SETFOCUS"
    If Nz(Me.chkReadOnly, False) Then parts.Add "READONLY"
    If Nz(Me.chkLocked, False) Then parts.Add "LOCKED"
    If Nz(Me.chkDisabled, False) Then parts.Add "DISABLED"

    ' Kontext
    If Len(Nz(Me.txtRole, "")) > 0 Then parts.Add "ROLE:" & Trim(Me.txtRole)
    If Len(Nz(Me.txtMod, "")) > 0 Then parts.Add "MOD:" & Trim(Me.txtMod)

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
    Me.txtResult = BuildTagString()
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
    If Me.chkInteger Then Me.chkNumeric = True
    UpdatePreview
End Sub

Private Sub txtMin_AfterUpdate()
    If Len(Nz(Me.txtMin, "")) > 0 Then Me.chkNumeric = True
    UpdatePreview
End Sub

Private Sub txtMax_AfterUpdate()
    If Len(Nz(Me.txtMax, "")) > 0 Then Me.chkNumeric = True
    UpdatePreview
End Sub

Public Sub LoadFromTag(ByVal TagValue As String)
    Dim tokens As Object
    Set tokens = ParseTagTokens(TagValue)

    ResetUI

    Me.chkRequired = tokens.Exists("REQUIRED")
    Me.chkNumeric = tokens.Exists("NUMERIC")
    Me.chkInteger = tokens.Exists("INTEGER")
    Me.chkDate = tokens.Exists("DATE")

    Me.chkHidden = tokens.Exists("HIDDEN")
    Me.chkSetFocus = tokens.Exists("SETFOCUS")
    Me.chkReadOnly = tokens.Exists("READONLY")
    Me.chkLocked = tokens.Exists("LOCKED")
    Me.chkDisabled = tokens.Exists("DISABLED")

    If tokens.Exists("MIN") Then Me.txtMin = tokens("MIN")
    If tokens.Exists("MAX") Then Me.txtMax = tokens("MAX")
    If tokens.Exists("MINLEN") Then Me.txtMinLen = tokens("MINLEN")
    If tokens.Exists("MAXLEN") Then Me.txtMaxLen = tokens("MAXLEN")
    If tokens.Exists("ROLE") Then Me.txtRole = tokens("ROLE")
    If tokens.Exists("MOD") Then Me.txtMod = tokens("MOD")

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

    Me.chkRequired = False
    Me.chkNumeric = False
    Me.chkInteger = False
    Me.chkDate = False

    Me.chkHidden = False
    Me.chkSetFocus = False
    Me.chkReadOnly = False
    Me.chkLocked = False
    Me.chkDisabled = False

    Me.txtMin = Null
    Me.txtMax = Null
    Me.txtMinLen = Null
    Me.txtMaxLen = Null
    Me.txtRole = Null
    Me.txtMod = Null
    Me.txtResult = Null
End Sub

Private Sub RefreshSubformCurrentRow(ByVal ControlName As String)
    On Error GoTo ErrorHandler

    Dim rsClone As DAO.Recordset

    Me.sfrmTmpTagComposer.Form.Requery

    Set rsClone = Me.sfrmTmpTagComposer.Form.RecordsetClone
    rsClone.FindFirst "ControlName = '" & Replace(ControlName, "'", "''") & "'"

    If Not rsClone.NoMatch Then
        Me.sfrmTmpTagComposer.Form.Bookmark = rsClone.Bookmark
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
    TokenKey = UCase$(Trim$(Nz(TokenKey, "")))

    Select Case TokenKey
        Case "REQUIRED", "NUMERIC", "INTEGER", "DATE", _
             "MIN", "MAX", "MINLEN", "MAXLEN", _
             "HIDDEN", "SETFOCUS", "READONLY", "LOCKED", "DISABLED", _
             "ROLE", "MOD"
            IsComposerManagedToken = True
    End Select
End Function

Private Function GetExistingTagForCurrentControl() As String
    On Error GoTo SafeExit

    Dim rs As DAO.Recordset
    Dim sqlText As String

    If Len(m_CurrentControl) = 0 Then Exit Function
    If Len(m_SessionId) = 0 Then Exit Function

    sqlText = "SELECT TagValue " & _
              "FROM tblTmpTagComposer " & _
              "WHERE SessionId = '" & Replace(m_SessionId, "'", "''") & "' " & _
              "  AND ControlName = '" & Replace(m_CurrentControl, "'", "''") & "';"

    Set rs = CurrentDb.OpenRecordset(sqlText, dbOpenSnapshot)

    If Not rs.EOF Then
        GetExistingTagForCurrentControl = Nz(rs!TagValue, vbNullString)
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

SafeExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

