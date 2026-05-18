VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPostalCodeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmPostalCodeSelect
' Purpose   : Modal selection dialog for ambiguous postal-code matches.
' Author    : Codex
' Version   : 0.1.1
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmPostalCodeSelect"
Private Const LOOKUP_MODE_ZIP As String = "ZIP"
Private Const TEMPVAR_POSTAL_SELECT_CONFIRMED As String = "fw_postal_select_confirmed"
Private Const TEMPVAR_POSTAL_SELECT_COUNTRY As String = "fw_postal_select_country_code"
Private Const TEMPVAR_POSTAL_SELECT_POSTAL As String = "fw_postal_select_postal_code"
Private Const TEMPVAR_POSTAL_SELECT_PLACE As String = "fw_postal_select_place_name"

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    ClearSelectionResult
    ConfigureListBox
    LoadMatches
    modFwTranslationRuntime.ApplyTranslations Me
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox "Fehler beim Laden der Postleitzahl-Auswahl: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo ErrorHandler

    ConfirmSelection
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdSelect_Click", Err
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo ErrorHandler

    CancelSelection
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cmdCancel_Click", Err
End Sub

Private Sub lstPostalCodes_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler

    ConfirmSelection
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "lstPostalCodes_DblClick", Err
End Sub

Private Sub ConfigureListBox()
    On Error GoTo ErrorHandler

    With Me!lstPostalCodes
        .RowSourceType = "Table/Query"
        .BoundColumn = 1
        .ColumnCount = 7
        .ColumnWidths = "0cm;1.2cm;1.8cm;3.5cm;3cm;3cm;3cm"
    End With
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureListBox", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub LoadMatches()
    On Error GoTo ErrorHandler

    Me!lstPostalCodes.RowSource = BuildRowSource()
    Me!lstPostalCodes.Requery

    If Me!lstPostalCodes.ListCount > 0 Then
        Me!lstPostalCodes = Me!lstPostalCodes.ItemData(0)
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".LoadMatches", _
        "Postal-code selection dialog loaded " & CStr(Me!lstPostalCodes.ListCount) & " row(s)."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadMatches", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Function BuildRowSource() As String
    Dim countryCode As String
    Dim postalCode As String
    Dim placeName As String
    Dim lookupMode As String
    Dim sqlStatement As String

    countryCode = UCase$(Trim$(GetOpenArgValue("country_code")))
    postalCode = Trim$(GetOpenArgValue("postal_code"))
    placeName = Trim$(GetOpenArgValue("place_name"))
    lookupMode = UCase$(Trim$(GetOpenArgValue("lookup_mode")))

    sqlStatement = "SELECT postal_code_id, country_code, postal_code, place_name, " & _
                   "state_name, province_name, community_name " & _
                   "FROM ref_postal_code " & _
                   "WHERE "

    If lookupMode = LOOKUP_MODE_ZIP Then
        sqlStatement = sqlStatement & "postal_code = " & SqlText(postalCode) & " "
    Else
        sqlStatement = sqlStatement & "UCase(place_name) = " & SqlText(UCase$(placeName)) & " "
    End If

    If LenB(countryCode) > 0 Then
        sqlStatement = sqlStatement & _
                       "AND country_code = " & SqlText(countryCode) & " "
    End If

    sqlStatement = sqlStatement & _
                   "AND is_active = True " & _
                   "ORDER BY country_code, postal_code, place_name, state_name, province_name, community_name;"

    BuildRowSource = sqlStatement
End Function

Private Sub ConfirmSelection()
    On Error GoTo ErrorHandler

    If Me!lstPostalCodes.ListIndex < 0 Then
        MsgBox "Bitte zuerst einen Ort auswaehlen.", vbExclamation, MODULE_NAME
        Exit Sub
    End If

    SetTempVar TEMPVAR_POSTAL_SELECT_CONFIRMED, True
    SetTempVar TEMPVAR_POSTAL_SELECT_COUNTRY, NzString(Me!lstPostalCodes.Column(1))
    SetTempVar TEMPVAR_POSTAL_SELECT_POSTAL, NzString(Me!lstPostalCodes.Column(2))
    SetTempVar TEMPVAR_POSTAL_SELECT_PLACE, NzString(Me!lstPostalCodes.Column(3))

    DoCmd.Close acForm, Me.Name
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfirmSelection", Err
End Sub

Private Sub CancelSelection()
    On Error GoTo ErrorHandler

    ClearSelectionResult
    SetTempVar TEMPVAR_POSTAL_SELECT_CONFIRMED, False
    DoCmd.Close acForm, Me.Name
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CancelSelection", Err
End Sub

Private Function GetOpenArgValue(ByVal KeyName As String) As String
    Dim parts() As String
    Dim pairText As Variant
    Dim separatorPosition As Long
    Dim currentKey As String

    If LenB(Trim$(Nz(Me.openArgs, vbNullString))) = 0 Then
        Exit Function
    End If

    parts = Split(CStr(Me.openArgs), ";")
    For Each pairText In parts
        separatorPosition = InStr(1, CStr(pairText), "=", vbTextCompare)
        If separatorPosition > 0 Then
            currentKey = Trim$(Left$(CStr(pairText), separatorPosition - 1))
            If StrComp(currentKey, KeyName, vbTextCompare) = 0 Then
                GetOpenArgValue = Trim$(Mid$(CStr(pairText), separatorPosition + 1))
                Exit Function
            End If
        End If
    Next pairText
End Function

Private Sub ClearSelectionResult()
    On Error Resume Next

    If TempVarExists(TEMPVAR_POSTAL_SELECT_CONFIRMED) Then TempVars.Remove TEMPVAR_POSTAL_SELECT_CONFIRMED
    If TempVarExists(TEMPVAR_POSTAL_SELECT_COUNTRY) Then TempVars.Remove TEMPVAR_POSTAL_SELECT_COUNTRY
    If TempVarExists(TEMPVAR_POSTAL_SELECT_POSTAL) Then TempVars.Remove TEMPVAR_POSTAL_SELECT_POSTAL
    If TempVarExists(TEMPVAR_POSTAL_SELECT_PLACE) Then TempVars.Remove TEMPVAR_POSTAL_SELECT_PLACE
End Sub

Private Function TempVarExists(ByVal TempVarName As String) As Boolean
    On Error GoTo SafeExit

    Dim tmpVar As TempVar

    For Each tmpVar In TempVars
        If StrComp(tmpVar.Name, TempVarName, vbTextCompare) = 0 Then
            TempVarExists = True
            Exit Function
        End If
    Next tmpVar

SafeExit:
End Function

Private Sub SetTempVar(ByVal TempVarName As String, ByVal TempVarValue As Variant)
    On Error GoTo ErrorHandler

    If TempVarExists(TempVarName) Then
        TempVars.Remove TempVarName
    End If

    TempVars.Add TempVarName, TempVarValue
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SetTempVar", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Function SqlText(ByVal Value As String) As String
    SqlText = "'" & Replace(NzString(Value), "'", "''") & "'"
End Function

Private Function NzString(ByVal Value As Variant, Optional ByVal DefaultValue As String = "") As String
    If IsNull(Value) Or IsEmpty(Value) Then
        NzString = DefaultValue
    Else
        NzString = CStr(Value)
    End If
End Function


