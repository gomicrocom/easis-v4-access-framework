Attribute VB_Name = "modFwTranslationEditService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFwTranslationEditService
' Purpose   : Focused translation edit helpers for maintaining exactly one
'             translation key across the required framework languages.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modFwTranslationEditService"

Private Const TABLE_TRANSLATION As String = "fw_translation"
Private Const TABLE_AUDIT As String = "tmp_fw_translation_audit"

Private Const FIELD_TRANSLATION_KEY As String = "translation_key"
Private Const FIELD_LANGUAGE_CODE As String = "language_code"
Private Const FIELD_TRANSLATION_VALUE As String = "translation_value"
Private Const FIELD_MODULE_CODE As String = "module_code"
Private Const FIELD_IS_ACTIVE As String = "is_active"
Private Const FIELD_CREATED_AT As String = "created_at"
Private Const FIELD_CREATED_BY As String = "created_by"
Private Const FIELD_UPDATED_AT As String = "updated_at"
Private Const FIELD_UPDATED_BY As String = "updated_by"
Private Const FIELD_SCOPE_CODE As String = "scope_code"
Private Const FIELD_AUDIT_STATUS As String = "audit_status"
Private Const FIELD_SOURCE_TYPE As String = "source_type"
Private Const FIELD_SOURCE_OBJECT As String = "source_object"
Private Const FIELD_SOURCE_CONTROL As String = "source_control"
Private Const FIELD_FALLBACK_TEXT As String = "fallback_text"

Private Const LANGUAGE_DE_CH As String = "DE-CH"
Private Const LANGUAGE_EN_US As String = "EN-US"
Private Const LANGUAGE_FR_FR As String = "FR-FR"

Public Sub LoadTranslationEditContext(ByVal targetForm As Access.Form, ByVal translationKey As String)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim metadataText As String
    Dim scopeCode As String

    translationKey = NormalizeTranslationKey(translationKey)
    If targetForm Is Nothing Then
        Exit Sub
    End If

    If LenB(translationKey) = 0 Then
        Err.Raise vbObjectError + 5600, MODULE_NAME & ".LoadTranslationEditContext", "Translation key is required."
    End If

    metadataText = ResolveAuditMetadata(translationKey)
    scopeCode = ResolveMetadataValue(metadataText, FIELD_SCOPE_CODE)
    If LenB(scopeCode) = 0 Then
        scopeCode = ResolveScopeFromTranslationKey(translationKey)
    End If

    EnsureRequiredLanguageRows translationKey, scopeCode

    Set db = CurrentDb

    SetControlValueIfPresent targetForm, "txtTranslationKey", translationKey
    SetControlValueIfPresent targetForm, "txtScopeCode", scopeCode
    SetControlValueIfPresent targetForm, "txtAuditStatus", ResolveMetadataValue(metadataText, FIELD_AUDIT_STATUS)
    SetControlValueIfPresent targetForm, "txtSourceType", ResolveMetadataValue(metadataText, FIELD_SOURCE_TYPE)
    SetControlValueIfPresent targetForm, "txtSourceObject", ResolveMetadataValue(metadataText, FIELD_SOURCE_OBJECT)
    SetControlValueIfPresent targetForm, "txtSourceControl", ResolveMetadataValue(metadataText, FIELD_SOURCE_CONTROL)
    SetControlValueIfPresent targetForm, "txtFallbackText", ResolveMetadataValue(metadataText, FIELD_FALLBACK_TEXT)
    SetControlValueIfPresent targetForm, "txtTranslationDeCh", LookupTranslationValue(db, translationKey, LANGUAGE_DE_CH)
    SetControlValueIfPresent targetForm, "txtTranslationEnUs", LookupTranslationValue(db, translationKey, LANGUAGE_EN_US)
    SetControlValueIfPresent targetForm, "txtTranslationFrFr", LookupTranslationValue(db, translationKey, LANGUAGE_FR_FR)

    modLoggingHandler.LogInfo MODULE_NAME & ".LoadTranslationEditContext", _
        "Loaded translation edit context for key '" & translationKey & "'."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadTranslationEditContext", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub EnsureRequiredLanguageRows(ByVal translationKey As String, Optional ByVal scopeCode As String = "")
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim languageCode As Variant

    translationKey = NormalizeTranslationKey(translationKey)
    If LenB(translationKey) = 0 Then
        Exit Sub
    End If

    scopeCode = UCase$(Trim$(modDaoHelper.NzString(scopeCode)))
    If LenB(scopeCode) = 0 Then
        scopeCode = ResolveScopeFromTranslationKey(translationKey)
    End If

    Set db = CurrentDb

    For Each languageCode In RequiredLanguageCodes()
        EnsureSingleLanguageRow db, translationKey, CStr(languageCode), scopeCode
    Next languageCode

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureRequiredLanguageRows", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub SaveTranslationValues(ByVal targetForm As Access.Form, ByVal translationKey As String)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim scopeCode As String

    If targetForm Is Nothing Then
        Exit Sub
    End If

    translationKey = NormalizeTranslationKey(translationKey)
    If LenB(translationKey) = 0 Then
        Err.Raise vbObjectError + 5601, MODULE_NAME & ".SaveTranslationValues", "Translation key is required."
    End If

    scopeCode = UCase$(Trim$(GetControlText(targetForm, "txtScopeCode")))
    If LenB(scopeCode) = 0 Then
        scopeCode = ResolveScopeFromTranslationKey(translationKey)
    End If

    EnsureRequiredLanguageRows translationKey, scopeCode

    Set db = CurrentDb

    SaveSingleTranslationValue db, translationKey, LANGUAGE_DE_CH, GetControlText(targetForm, "txtTranslationDeCh"), scopeCode
    SaveSingleTranslationValue db, translationKey, LANGUAGE_EN_US, GetControlText(targetForm, "txtTranslationEnUs"), scopeCode
    SaveSingleTranslationValue db, translationKey, LANGUAGE_FR_FR, GetControlText(targetForm, "txtTranslationFrFr"), scopeCode

    modLoggingHandler.LogInfo MODULE_NAME & ".SaveTranslationValues", _
        "Saved translation values for key '" & translationKey & "'."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SaveTranslationValues", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ResolveAuditMetadata(ByVal translationKey As String) As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlStatement As String
    Dim metadataText As String
    Dim scopeCode As String

    translationKey = NormalizeTranslationKey(translationKey)
    If LenB(translationKey) = 0 Then
        Exit Function
    End If

    scopeCode = ResolveScopeFromTranslationKey(translationKey)
    metadataText = AddMetadataPair(metadataText, FIELD_SCOPE_CODE, scopeCode)

    Set db = CurrentDb
    If Not TableExists(db, TABLE_AUDIT) Then
        ResolveAuditMetadata = metadataText
        Exit Function
    End If

    sqlStatement = "SELECT TOP 1 scope_code, audit_status, source_type, source_object, source_control, fallback_text " & _
                   "FROM " & TABLE_AUDIT & " " & _
                   "WHERE translation_key = " & SqlText(translationKey) & " " & _
                   "ORDER BY IIf([audit_status]='ORPHAN',1," & _
                   "IIf([audit_status]='LEGACY_KEY',2," & _
                   "IIf([audit_status]='MISSING_ROW',3," & _
                   "IIf([audit_status]='EMPTY_VALUE',4,5)))), [language_code];"

    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)
    If Not rs.EOF Then
        metadataText = AddMetadataPair(metadataText, FIELD_SCOPE_CODE, modDaoHelper.NzString(rs.Fields(FIELD_SCOPE_CODE).Value, scopeCode))
        metadataText = AddMetadataPair(metadataText, FIELD_AUDIT_STATUS, modDaoHelper.NzString(rs.Fields(FIELD_AUDIT_STATUS).Value))
        metadataText = AddMetadataPair(metadataText, FIELD_SOURCE_TYPE, modDaoHelper.NzString(rs.Fields(FIELD_SOURCE_TYPE).Value))
        metadataText = AddMetadataPair(metadataText, FIELD_SOURCE_OBJECT, modDaoHelper.NzString(rs.Fields(FIELD_SOURCE_OBJECT).Value))
        metadataText = AddMetadataPair(metadataText, FIELD_SOURCE_CONTROL, modDaoHelper.NzString(rs.Fields(FIELD_SOURCE_CONTROL).Value))
        metadataText = AddMetadataPair(metadataText, FIELD_FALLBACK_TEXT, modDaoHelper.NzString(rs.Fields(FIELD_FALLBACK_TEXT).Value))
    End If

    rs.Close
    Set rs = Nothing

    ResolveAuditMetadata = metadataText
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    ResolveAuditMetadata = AddMetadataPair(vbNullString, FIELD_SCOPE_CODE, ResolveScopeFromTranslationKey(translationKey))
    modErrorHandler.HandleError MODULE_NAME, "ResolveAuditMetadata", Err
End Function

Private Sub EnsureSingleLanguageRow(ByVal db As DAO.Database, ByVal translationKey As String, ByVal languageCode As String, ByVal scopeCode As String)
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset

    Set rs = db.OpenRecordset( _
        "SELECT * FROM " & TABLE_TRANSLATION & " " & _
        "WHERE translation_key = " & SqlText(translationKey) & " " & _
        "AND language_code = " & SqlText(languageCode) & ";", dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs.Fields(FIELD_TRANSLATION_KEY).Value = translationKey
        rs.Fields(FIELD_LANGUAGE_CODE).Value = languageCode
        rs.Fields(FIELD_MODULE_CODE).Value = scopeCode
        rs.Fields(FIELD_IS_ACTIVE).Value = True
        SetAuditRecordFields rs, True
        rs.Update
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureSingleLanguageRow", Err
    Resume CleanExit
End Sub

Private Sub SaveSingleTranslationValue( _
    ByVal db As DAO.Database, _
    ByVal translationKey As String, _
    ByVal languageCode As String, _
    ByVal translationValue As String, _
    ByVal scopeCode As String)
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset

    Set rs = db.OpenRecordset( _
        "SELECT * FROM " & TABLE_TRANSLATION & " " & _
        "WHERE translation_key = " & SqlText(translationKey) & " " & _
        "AND language_code = " & SqlText(languageCode) & ";", dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs.Fields(FIELD_TRANSLATION_KEY).Value = translationKey
        rs.Fields(FIELD_LANGUAGE_CODE).Value = languageCode
        rs.Fields(FIELD_MODULE_CODE).Value = scopeCode
        rs.Fields(FIELD_IS_ACTIVE).Value = True
        SetAuditRecordFields rs, True
    Else
        rs.Edit
        If modDaoHelper.RecordsetHasField(rs, FIELD_MODULE_CODE) Then
            rs.Fields(FIELD_MODULE_CODE).Value = scopeCode
        End If
        If modDaoHelper.RecordsetHasField(rs, FIELD_IS_ACTIVE) Then
            rs.Fields(FIELD_IS_ACTIVE).Value = True
        End If
        SetAuditRecordFields rs, False
    End If

    If LenB(Trim$(translationValue)) = 0 Then
        rs.Fields(FIELD_TRANSLATION_VALUE).Value = Null
    Else
        rs.Fields(FIELD_TRANSLATION_VALUE).Value = translationValue
    End If

    rs.Update

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SaveSingleTranslationValue", Err
    Resume CleanExit
End Sub

Private Sub SetAuditRecordFields(ByVal rs As DAO.Recordset, ByVal isInsert As Boolean)
    Dim currentUser As String
    Dim currentTimestamp As Date

    currentUser = modAuditHelper.ResolveAuditUserName()
    currentTimestamp = Now()

    If isInsert Then
        If modDaoHelper.RecordsetHasField(rs, FIELD_CREATED_AT) Then
            If IsNull(rs.Fields(FIELD_CREATED_AT).Value) Then
                rs.Fields(FIELD_CREATED_AT).Value = currentTimestamp
            End If
        End If

        If modDaoHelper.RecordsetHasField(rs, FIELD_CREATED_BY) Then
            If LenB(Trim$(modDaoHelper.NzString(rs.Fields(FIELD_CREATED_BY).Value))) = 0 Then
                rs.Fields(FIELD_CREATED_BY).Value = currentUser
            End If
        End If
    End If

    If modDaoHelper.RecordsetHasField(rs, FIELD_UPDATED_AT) Then
        rs.Fields(FIELD_UPDATED_AT).Value = currentTimestamp
    End If

    If modDaoHelper.RecordsetHasField(rs, FIELD_UPDATED_BY) Then
        rs.Fields(FIELD_UPDATED_BY).Value = currentUser
    End If
End Sub

Private Function LookupTranslationValue(ByVal db As DAO.Database, ByVal translationKey As String, ByVal languageCode As String) As String
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset

    Set rs = db.OpenRecordset( _
        "SELECT translation_value FROM " & TABLE_TRANSLATION & " " & _
        "WHERE translation_key = " & SqlText(translationKey) & " " & _
        "AND language_code = " & SqlText(languageCode) & ";", dbOpenSnapshot)

    If Not rs.EOF Then
        LookupTranslationValue = modDaoHelper.NzString(rs.Fields(FIELD_TRANSLATION_VALUE).Value)
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LookupTranslationValue", Err
    Resume CleanExit
End Function

Private Function RequiredLanguageCodes() As Variant
    RequiredLanguageCodes = Array(LANGUAGE_DE_CH, LANGUAGE_EN_US, LANGUAGE_FR_FR)
End Function

Private Function ResolveScopeFromTranslationKey(ByVal translationKey As String) As String
    Dim separatorPosition As Long

    translationKey = NormalizeTranslationKey(translationKey)
    If LenB(translationKey) = 0 Then
        Exit Function
    End If

    separatorPosition = InStr(1, translationKey, ".", vbBinaryCompare)
    If separatorPosition > 1 Then
        ResolveScopeFromTranslationKey = UCase$(Left$(translationKey, separatorPosition - 1))
    ElseIf InStr(1, translationKey, "_", vbBinaryCompare) > 1 Then
        ResolveScopeFromTranslationKey = UCase$(Left$(translationKey, InStr(1, translationKey, "_", vbBinaryCompare) - 1))
    Else
        ResolveScopeFromTranslationKey = "COMMON"
    End If
End Function

Private Function NormalizeTranslationKey(ByVal translationKey As String) As String
    translationKey = Trim$(modDaoHelper.NzString(translationKey))
    If UCase$(Left$(translationKey, 3)) = "TR:" Then
        translationKey = Mid$(translationKey, 4)
    End If
    NormalizeTranslationKey = translationKey
End Function

Private Function AddMetadataPair(ByVal metadataText As String, ByVal keyName As String, ByVal valueText As String) As String
    If LenB(metadataText) > 0 Then
        metadataText = metadataText & ";"
    End If

    AddMetadataPair = metadataText & keyName & "=" & EscapeMetadataValue(valueText)
End Function

Private Function ResolveMetadataValue(ByVal metadataText As String, ByVal keyName As String) As String
    Dim parts() As String
    Dim pairText As Variant
    Dim separatorPosition As Long
    Dim currentKey As String

    If LenB(metadataText) = 0 Then
        Exit Function
    End If

    parts = Split(metadataText, ";")
    For Each pairText In parts
        separatorPosition = InStr(1, CStr(pairText), "=", vbBinaryCompare)
        If separatorPosition > 0 Then
            currentKey = Left$(CStr(pairText), separatorPosition - 1)
            If StrComp(currentKey, keyName, vbTextCompare) = 0 Then
                ResolveMetadataValue = UnescapeMetadataValue(Mid$(CStr(pairText), separatorPosition + 1))
                Exit Function
            End If
        End If
    Next pairText
End Function

Private Function EscapeMetadataValue(ByVal valueText As String) As String
    valueText = Replace(modDaoHelper.NzString(valueText), "%", "%25")
    valueText = Replace(valueText, ";", "%3B")
    valueText = Replace(valueText, "=", "%3D")
    EscapeMetadataValue = valueText
End Function

Private Function UnescapeMetadataValue(ByVal valueText As String) As String
    valueText = Replace(valueText, "%3D", "=")
    valueText = Replace(valueText, "%3B", ";")
    valueText = Replace(valueText, "%25", "%")
    UnescapeMetadataValue = valueText
End Function

Private Function GetControlText(ByVal targetForm As Access.Form, ByVal controlName As String) As String
    On Error GoTo SafeExit

    If HasControl(targetForm, controlName) Then
        GetControlText = modDaoHelper.NzString(targetForm.Controls(controlName).Value)
    End If

SafeExit:
End Function

Private Sub SetControlValueIfPresent(ByVal targetForm As Access.Form, ByVal controlName As String, ByVal controlValue As Variant)
    On Error GoTo SafeExit

    If HasControl(targetForm, controlName) Then
        targetForm.Controls(controlName).Value = controlValue
    End If

SafeExit:
End Sub

Private Function HasControl(ByVal targetForm As Access.Form, ByVal controlName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentControl As Control

    If targetForm Is Nothing Then
        Exit Function
    End If

    For Each currentControl In targetForm.Controls
        If StrComp(currentControl.Name, controlName, vbTextCompare) = 0 Then
            HasControl = True
            Exit Function
        End If
    Next currentControl

SafeExit:
End Function

Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    On Error GoTo SafeExit

    Dim tableDefinition As DAO.TableDef

    For Each tableDefinition In db.TableDefs
        If StrComp(tableDefinition.Name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next tableDefinition

SafeExit:
    Set tableDefinition = Nothing
End Function
