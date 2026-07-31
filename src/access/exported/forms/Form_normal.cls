Option Compare Database
Option Explicit

'===============================================================================
' Module    : Form_frmAddressDetail
' Purpose   : Address detail maintenance with runtime translations, validation,
'             and postal reference lookups.
' Author    : Codex
' Version   : 0.6.0
'===============================================================================

Private Const MODULE_NAME As String = "Form_frmAddressDetail"
Private Const DIALOG_FORM_POSTAL_CODE_SELECT As String = "frmPostalCodeSelect"
Private Const LOOKUP_MODE_ZIP As String = "ZIP"
Private Const LOOKUP_MODE_CITY As String = "CITY"
Private Const TEMPVAR_POSTAL_SELECT_CONFIRMED As String = "fw_postal_select_confirmed"
Private Const TEMPVAR_POSTAL_SELECT_COUNTRY As String = "fw_postal_select_country_code"
Private Const TEMPVAR_POSTAL_SELECT_POSTAL As String = "fw_postal_select_postal_code"
Private Const TEMPVAR_POSTAL_SELECT_PLACE As String = "fw_postal_select_place_name"
Private Const COUNTRY_ROW_SOURCE As String = _
    "SELECT country_code, country_name " & _
    "FROM ref_country " & _
    "ORDER BY country_name;"
Private Const LANGUAGE_ROW_SOURCE As String = _
    "de-CH;Deutsch Schweiz;" & _
    "de-DE;Deutsch Deutschland;" & _
    "fr-CH;Franzoesisch Schweiz;" & _
    "it-CH;Italienisch Schweiz;" & _
    "en-US;Englisch USA"

Private Sub cboAddressingMode_AfterUpdate()
    On Error GoTo ErrorHandler

    UpdateAddressingModeUi

    Exit Sub
    
ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboAddressingMode_AfterUpdate", Err
End Sub

Private Sub cboCity_BeforeUpdate(Cancel As Integer)
    On Error GoTo ErrorHandler

    Me!cboCity = NormalizePlaceName(Nz(Me!cboCity.text, vbNullString))

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboCity_BeforeUpdate", Err
End Sub

Private Sub Form_Current()
    On Error GoTo ErrorHandler

    UpdateAddressingModeUi
    
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Current", Err
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler

    ConfigureReferenceCombos
    modFwTranslationRuntime.ApplyTranslations Me

    modLoggingHandler.LogInfo MODULE_NAME & ".Form_Load", _
        "Address detail loaded."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_Load", Err
    MsgBox "Fehler beim Laden der Adressdetailmaske: " & Err.description, vbExclamation, MODULE_NAME
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    On Error GoTo ErrorHandler

    If Not modFwValidationRuntime.ValidateForm(Me) Then
        Cancel = True
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "Form_BeforeUpdate", Err
    Cancel = True
End Sub

Private Sub txtZipCode_AfterUpdate()
    On Error GoTo ErrorHandler

    ApplyPostalLookupByZip
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtZipCode_AfterUpdate", Err
End Sub

Private Sub txtCity_AfterUpdate()
    On Error GoTo ErrorHandler

    ApplyPostalLookupByCity
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "txtCity_AfterUpdate", Err
End Sub

Private Sub cboCity_Change()
    On Error GoTo ErrorHandler

    UpdateCitySuggestions
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboCity_Change", Err
End Sub

Private Sub cboCity_AfterUpdate()
    On Error GoTo ErrorHandler

    ApplyPostalLookupByCity
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "cboCity_AfterUpdate", Err
End Sub

Private Sub ConfigureReferenceCombos()
    On Error GoTo ErrorHandler

    ConfigureAddressTypeCombo
    ConfigureSalutationCombo
    ConfigureAddressingModeCombo
    ConfigurePreferredNameControl
    ConfigureCountryCombo
    ConfigureLanguageCombo
    ConfigureCityCombo
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureReferenceCombos", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub ConfigurePreferredNameControl()
    On Error GoTo ErrorHandler

    If Not HasControl("txtPreferredName") Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ConfigurePreferredNameControl", _
            "Control txtPreferredName not found."
        Exit Sub
    End If

    Me!txtPreferredName.ControlSource = "preferred_name"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigurePreferredNameControl", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub ConfigureAddressTypeCombo()
    On Error GoTo ErrorHandler

    Dim currentLanguageCode As String
    Dim sqlStatement As String

    If Not HasControl("cboAddressType") Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ConfigureAddressTypeCombo", _
            "Control cboAddressType not found."
        Exit Sub
    End If

    currentLanguageCode = modFwTranslationRuntime.GetCurrentLanguageCode()
    sqlStatement = BuildTranslatedReferenceRowSource("ref_address_type", "address_type_code", currentLanguageCode)
    
    With Me!cboAddressType
        .ControlSource = "address_type_code"
        .RowSourceType = "Table/Query"
        .RowSource = sqlStatement
        .BoundColumn = 1
        .ColumnCount = 2
        .ColumnWidths = "0cm;4cm"
        .LimitToList = True
        .Requery
    End With
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureAddressTypeCombo", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub ConfigureSalutationCombo()
    On Error GoTo ErrorHandler

    Dim currentLanguageCode As String
    Dim sqlStatement As String

    If Not HasControl("cboSalutation") Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ConfigureSalutationCombo", _
            "Control cboSalutation not found."
        Exit Sub
    End If

    currentLanguageCode = modFwTranslationRuntime.GetCurrentLanguageCode()
    sqlStatement = BuildTranslatedReferenceRowSource("ref_salutation", "salutation_code", currentLanguageCode)
    
    With Me!cboSalutation
        .ControlSource = "salutation_code"
        .RowSourceType = "Table/Query"
        .RowSource = sqlStatement
        .BoundColumn = 1
        .ColumnCount = 2
        .ColumnWidths = "0cm;4cm"
        .LimitToList = True
        .Requery
    End With
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureSalutationCombo", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub ConfigureAddressingModeCombo()
    On Error GoTo ErrorHandler

    Dim currentLanguageCode As String
    Dim sqlStatement As String

    If Not HasControl("cboAddressingMode") Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ConfigureAddressingModeCombo", _
            "Control cboAddressingMode not found."
        Exit Sub
    End If

    currentLanguageCode = modFwTranslationRuntime.GetCurrentLanguageCode()
    sqlStatement = BuildTranslatedReferenceRowSource("ref_addressing_mode", "addressing_mode_code", currentLanguageCode)

    With Me!cboAddressingMode
        .ControlSource = "addressing_mode_code"
        .RowSourceType = "Table/Query"
        .RowSource = sqlStatement
        .BoundColumn = 1
        .ColumnCount = 2
        .ColumnWidths = "0cm;4cm"
        .LimitToList = True
        .Requery
    End With
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureAddressingModeCombo", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub ConfigureCityCombo()
    On Error GoTo ErrorHandler

    If Not HasControl("cboCity") Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ConfigureCityCombo", _
            "Control cboCity not found."
        Exit Sub
    End If

    With Me!cboCity
        .ControlSource = "city"
        .RowSourceType = "Table/Query"
        .RowSource = "SELECT DISTINCT place_name FROM ref_postal_code WHERE 1=0;"
        .BoundColumn = 1
        .ColumnCount = 1
        .ColumnWidths = "5cm"
        .LimitToList = False
        .AutoExpand = True
    End With
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureCityCombo", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub ConfigureCountryCombo()
    On Error GoTo ErrorHandler

    If Not HasControl("cboCountryCode") Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ConfigureCountryCombo", _
            "Control cboCountryCode not found."
        Exit Sub
    End If

    With Me!cboCountryCode
        .ControlSource = "country_code"
        .RowSourceType = "Table/Query"
        .RowSource = COUNTRY_ROW_SOURCE
        .BoundColumn = 1
        .ColumnCount = 2
        .ColumnWidths = "1.5cm;5cm"
        .LimitToList = True
    End With
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureCountryCombo", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub ConfigureLanguageCombo()
    On Error GoTo ErrorHandler

    If Not HasControl("cboLanguageCode") Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ConfigureLanguageCombo", _
            "Control cboLanguageCode not found."
        Exit Sub
    End If

    With Me!cboLanguageCode
        .ControlSource = "language_code"
        .RowSourceType = "Value List"
        .RowSource = LANGUAGE_ROW_SOURCE
        .BoundColumn = 1
        .ColumnCount = 2
        .ColumnWidths = "0cm;5cm"
        .LimitToList = True
    End With
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureLanguageCombo", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub ApplyPostalLookupByZip()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim zipCode As String
    Dim cityValue As String
    Dim currentCountryCode As String
    Dim selectedCountryCode As String
    Dim selectedPostalCode As String
    Dim selectedPlaceName As String
    Dim sqlStatement As String

    zipCode = Trim$(GetControlTextValue("txtZipCode", "zip_code"))
    If LenB(zipCode) = 0 Then
        Exit Sub
    End If

    currentCountryCode = UCase$(Trim$(GetControlTextValue("cboCountryCode", "country_code")))

    sqlStatement = BuildPostalLookupSqlStatement(LOOKUP_MODE_ZIP, currentCountryCode, zipCode, vbNullString)

    Set db = CurrentDb
    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)

    If rs.EOF Then
        Exit Sub
    End If

    selectedPlaceName = NzString(rs.Fields("place_name").Value)
    selectedCountryCode = NzString(rs.Fields("country_code").Value)
    selectedPostalCode = NzString(rs.Fields("postal_code").Value)

    rs.MoveNext
    If Not rs.EOF Then
        modLoggingHandler.LogInfo MODULE_NAME & ".ApplyPostalLookupByZip", _
            "Multiple postal-code matches found for zip_code '" & zipCode & _
            IIf(LenB(currentCountryCode) > 0, "' in country '" & currentCountryCode & "'.", "'.")

        If Not PromptPostalSelection(currentCountryCode, zipCode, vbNullString, LOOKUP_MODE_ZIP, _
                                     selectedCountryCode, selectedPostalCode, selectedPlaceName) Then
            GoTo CleanExit
        End If
    End If

    ApplyPostalSelection currentCountryCode, selectedCountryCode, selectedPostalCode, selectedPlaceName

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyPostalLookupByZip", _
        "Resolved zip_code '" & zipCode & "' to city '" & selectedPlaceName & "'."

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyPostalLookupByZip", Err
    Resume CleanExit
End Sub

Private Sub ApplyPostalLookupByCity()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim cityValue As String
    Dim currentCountryCode As String
    Dim selectedCountryCode As String
    Dim selectedPostalCode As String
    Dim selectedPlaceName As String
    Dim sqlStatement As String

    cityValue = Trim$(GetControlTextValue("cboCity", "txtCity"))
    If LenB(cityValue) = 0 Then
        Exit Sub
    End If

    currentCountryCode = UCase$(Trim$(GetControlTextValue("cboCountryCode", "country_code")))

    sqlStatement = BuildPostalLookupSqlStatement(LOOKUP_MODE_CITY, currentCountryCode, vbNullString, cityValue)

    Set db = CurrentDb
    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)

    If rs.EOF Then
        Exit Sub
    End If

    selectedPlaceName = NzString(rs.Fields("place_name").Value)
    selectedCountryCode = NzString(rs.Fields("country_code").Value)
    selectedPostalCode = NzString(rs.Fields("postal_code").Value)

    rs.MoveNext
    If Not rs.EOF Then
        modLoggingHandler.LogInfo MODULE_NAME & ".ApplyPostalLookupByCity", _
            "Multiple postal-code matches found for city '" & cityValue & _
            IIf(LenB(currentCountryCode) > 0, "' in country '" & currentCountryCode & "'.", "'.")

        If Not PromptPostalSelection(currentCountryCode, vbNullString, cityValue, LOOKUP_MODE_CITY, _
                                     selectedCountryCode, selectedPostalCode, selectedPlaceName) Then
            GoTo CleanExit
        End If
    End If

    ApplyPostalSelection currentCountryCode, selectedCountryCode, selectedPostalCode, selectedPlaceName

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyPostalLookupByCity", _
        "Resolved city '" & cityValue & "' to zip_code '" & selectedPostalCode & "'."

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyPostalLookupByCity", Err
    Resume CleanExit
End Sub

Private Sub UpdateCitySuggestions()
    On Error GoTo ErrorHandler

    Dim typedValue As String
    Dim currentCountryCode As String
    Dim sqlStatement As String

    If Not HasControl("cboCity") Then
        Exit Sub
    End If

    typedValue = Trim$(NzString(Me!cboCity.text))
    If Len(typedValue) < 2 Then
        Me!cboCity.RowSource = "SELECT DISTINCT place_name FROM ref_postal_code WHERE 1=0;"
        Exit Sub
    End If

    currentCountryCode = UCase$(Trim$(GetControlTextValue("cboCountryCode", "country_code")))

    sqlStatement = "SELECT DISTINCT place_name " & _
                   "FROM ref_postal_code " & _
                   "WHERE place_name LIKE " & SqlText(typedValue & "*") & " " & _
                   "AND Nz(is_active, True) = True "

    If LenB(currentCountryCode) > 0 Then
        sqlStatement = sqlStatement & _
                       "AND country_code = " & SqlText(currentCountryCode) & " "
    End If

    sqlStatement = sqlStatement & _
                   "ORDER BY place_name;"

    Me!cboCity.RowSource = sqlStatement
    Me!cboCity.Requery
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "UpdateCitySuggestions", Err
End Sub

Private Function BuildPostalLookupSqlStatement( _
    ByVal lookupMode As String, _
    ByVal currentCountryCode As String, _
    ByVal postalCode As String, _
    ByVal placeName As String) As String

    Dim sqlStatement As String

    sqlStatement = "SELECT TOP 2 postal_code, place_name, country_code " & _
                   "FROM ref_postal_code " & _
                   "WHERE "

    Select Case UCase$(Trim$(lookupMode))
        Case LOOKUP_MODE_ZIP
            sqlStatement = sqlStatement & "postal_code = " & SqlText(Trim$(postalCode)) & " "
        Case Else
            sqlStatement = sqlStatement & "UCase(place_name) = " & SqlText(UCase$(Trim$(placeName))) & " "
    End Select

    If LenB(Trim$(currentCountryCode)) > 0 Then
        sqlStatement = sqlStatement & _
                       "AND country_code = " & SqlText(UCase$(Trim$(currentCountryCode))) & " "
    End If

    sqlStatement = sqlStatement & _
                   "AND Nz(is_active, True) = True " & _
                   "ORDER BY country_code, postal_code, place_name;"

    BuildPostalLookupSqlStatement = sqlStatement
End Function

Private Function BuildTranslatedReferenceRowSource( _
    ByVal tableName As String, _
    ByVal codeFieldName As String, _
    ByVal LanguageCode As String) As String

    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "SELECT "
    sqlStatement = sqlStatement & "r." & codeFieldName & ", "
    sqlStatement = sqlStatement & "IIf(t.translation_value Is Null, "
    sqlStatement = sqlStatement & "r." & codeFieldName & ", "
    sqlStatement = sqlStatement & "Left(t.translation_value, 255)) AS display_text "
    sqlStatement = sqlStatement & "FROM " & tableName & " AS r "
    sqlStatement = sqlStatement & "LEFT JOIN fw_translation AS t "
    sqlStatement = sqlStatement & "ON r.translation_key = t.translation_key "
    sqlStatement = sqlStatement & "WHERE Nz(r.is_active, True) = True "
    sqlStatement = sqlStatement & "AND t.language_code = " & SqlText(LanguageCode) & " "
    sqlStatement = sqlStatement & "AND Nz(t.is_active, True) = True "
    sqlStatement = sqlStatement & "ORDER BY r.sort_order, r." & codeFieldName & ";"

    BuildTranslatedReferenceRowSource = sqlStatement
End Function

Private Function PromptPostalSelection( _
    ByVal currentCountryCode As String, _
    ByVal postalCode As String, _
    ByVal placeName As String, _
    ByVal lookupMode As String, _
    ByRef selectedCountryCode As String, _
    ByRef selectedPostalCode As String, _
    ByRef selectedPlaceName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim openArgsText As String

    ClearPostalSelectionResult

    openArgsText = BuildPostalSelectionOpenArgs(currentCountryCode, postalCode, placeName, lookupMode)

    DoCmd.OpenForm _
        FormName:=DIALOG_FORM_POSTAL_CODE_SELECT, _
        View:=acNormal, _
        WindowMode:=acDialog, _
        openArgs:=openArgsText

    If Not TempVarExists(TEMPVAR_POSTAL_SELECT_CONFIRMED) Then
        Exit Function
    End If

    If Not CBool(TempVars(TEMPVAR_POSTAL_SELECT_CONFIRMED).Value) Then
        Exit Function
    End If

    selectedCountryCode = GetTempVarString(TEMPVAR_POSTAL_SELECT_COUNTRY)
    selectedPostalCode = GetTempVarString(TEMPVAR_POSTAL_SELECT_POSTAL)
    selectedPlaceName = GetTempVarString(TEMPVAR_POSTAL_SELECT_PLACE)
    PromptPostalSelection = True

CleanExit:
    ClearPostalSelectionResult
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "PromptPostalSelection", Err
    Resume CleanExit
End Function

Private Sub ApplyPostalSelection( _
    ByVal currentCountryCode As String, _
    ByVal selectedCountryCode As String, _
    ByVal selectedPostalCode As String, _
    ByVal selectedPlaceName As String)

    If LenB(selectedPlaceName) > 0 Then
        SetControlValueIfPresent "cboCity", selectedPlaceName
        SetControlValueIfPresent "txtCity", selectedPlaceName
        SetControlValueIfPresent "city", selectedPlaceName
    End If

    If LenB(selectedPostalCode) > 0 Then
        SetControlValueIfPresent "txtZipCode", selectedPostalCode
        SetControlValueIfPresent "zip_code", selectedPostalCode
    End If

    If LenB(selectedCountryCode) > 0 Then
        If LenB(currentCountryCode) = 0 Then
            SetControlValueIfPresent "cboCountryCode", selectedCountryCode
            SetControlValueIfPresent "country_code", selectedCountryCode
        End If
    End If
End Sub

Private Function BuildPostalSelectionOpenArgs( _
    ByVal currentCountryCode As String, _
    ByVal postalCode As String, _
    ByVal placeName As String, _
    ByVal lookupMode As String) As String

    BuildPostalSelectionOpenArgs = _
        "country_code=" & SanitizeOpenArgsValue(currentCountryCode) & ";" & _
        "postal_code=" & SanitizeOpenArgsValue(postalCode) & ";" & _
        "place_name=" & SanitizeOpenArgsValue(placeName) & ";" & _
        "lookup_mode=" & SanitizeOpenArgsValue(lookupMode)
End Function

Private Function SanitizeOpenArgsValue(ByVal Value As String) As String
    Value = Replace(NzString(Value), ";", " ")
    Value = Replace(Value, "=", " ")
    SanitizeOpenArgsValue = Trim$(Value)
End Function

Private Sub ClearPostalSelectionResult()
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

Private Function GetTempVarString(ByVal TempVarName As String) As String
    On Error GoTo SafeExit

    If TempVarExists(TempVarName) Then
        GetTempVarString = NzString(TempVars(TempVarName).Value)
    End If

SafeExit:
End Function

Private Function GetControlTextValue( _
    ByVal PreferredControlName As String, _
    ByVal FallbackControlName As String) As String
    On Error GoTo SafeExit

    If HasControl(PreferredControlName) Then
        GetControlTextValue = NzString(Me.Controls(PreferredControlName).Value)
        Exit Function
    End If

    If HasControl(FallbackControlName) Then
        GetControlTextValue = NzString(Me.Controls(FallbackControlName).Value)
    End If

SafeExit:
End Function

Private Sub SetControlValueIfPresent(ByVal ControlName As String, ByVal ValueToSet As Variant)
    On Error GoTo SafeExit

    If HasControl(ControlName) Then
        Me.Controls(ControlName).Value = ValueToSet
    End If

SafeExit:
End Sub

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

Private Function NormalizePlaceName(ByVal Value As String) As String
    Dim words() As String
    Dim i As Long

    Value = Trim$(LCase$(Value))

    If LenB(Value) = 0 Then
        Exit Function
    End If

    words = Split(Value, " ")

    For i = LBound(words) To UBound(words)
        If LenB(words(i)) > 0 Then
            words(i) = UCase$(Left$(words(i), 1)) & Mid$(words(i), 2)
        End If
    Next i

    NormalizePlaceName = Join(words, " ")
End Function

Private Sub UpdateAddressingModeUi()

    Dim isInformal As Boolean

    isInformal = _
        (UCase$(Nz(Me!cboAddressingMode.Value, "")) = "INFORMAL")

    Me!txtPreferredName.Visible = isInformal
    Me!lblPreferredName.Visible = isInformal

End Sub


