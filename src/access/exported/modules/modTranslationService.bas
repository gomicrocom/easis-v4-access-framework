Option Compare Database
Option Explicit

'===============================================================================
' Module    : modTranslationService
' Purpose   : Centralized translation lookup service for runtime language support.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modTranslationService"
Private Const FALLBACK_LANGUAGE As String = "EN"
Private Const TABLE_FW_TRANSLATIONS As String = "tblFwTranslations"
Private Const FIELD_TRANSLATION_KEY As String = "TranslationKey"
Private Const FIELD_LANGUAGE_CODE As String = "LanguageCode"
Private Const FIELD_TRANSLATION_VALUE As String = "TranslationValue"
Private Const FIELD_IS_ACTIVE As String = "IsActive"

Private mTranslations As Object
Private mCurrentLanguage As String
Private mDefaultLanguage As String

Public Sub InitializeTranslations()
    On Error GoTo ErrorHandler

    Set mTranslations = CreateObject("Scripting.Dictionary")
    mTranslations.CompareMode = vbTextCompare

    mDefaultLanguage = ResolveDefaultLanguage()
    mCurrentLanguage = mDefaultLanguage

    LoadTranslations

    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeTranslations", _
        "Translations initialized for language '" & mCurrentLanguage & "'."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "InitializeTranslations", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub SetCurrentLanguage(ByVal LanguageCode As String)
    On Error GoTo ErrorHandler

    Dim normalizedLanguage As String

    If LenB(mDefaultLanguage) = 0 Then
        mDefaultLanguage = ResolveDefaultLanguage()
    End If

    normalizedLanguage = NormalizeLanguageCode(LanguageCode)
    If LenB(normalizedLanguage) = 0 Then
        mCurrentLanguage = mDefaultLanguage
    Else
        mCurrentLanguage = normalizedLanguage
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SetCurrentLanguage", Err
End Sub

Public Function GetCurrentLanguage() As String
    On Error GoTo ErrorHandler

    If LenB(mCurrentLanguage) = 0 Then
        mCurrentLanguage = ResolveDefaultLanguage()
    End If

    GetCurrentLanguage = mCurrentLanguage
    Exit Function

ErrorHandler:
    GetCurrentLanguage = FALLBACK_LANGUAGE
    modErrorHandler.HandleError MODULE_NAME, "GetCurrentLanguage", Err
End Function

Public Function T(ByVal TextKey As String, Optional ByVal Fallback As String = "") As String
    On Error GoTo ErrorHandler

    Dim normalizedKey As String
    Dim currentLanguageCode As String
    Dim defaultLanguageCode As String
    Dim translatedValue As String

    normalizedKey = NormalizeTextKey(TextKey)
    If LenB(normalizedKey) = 0 Then
        T = Fallback
        Exit Function
    End If

    EnsureTranslationStore

    currentLanguageCode = GetCurrentLanguage()
    translatedValue = LookupTranslation(currentLanguageCode, normalizedKey)
    If LenB(translatedValue) > 0 Then
        T = translatedValue
        Exit Function
    End If

    If LenB(mDefaultLanguage) = 0 Then
        mDefaultLanguage = ResolveDefaultLanguage()
    End If

    defaultLanguageCode = mDefaultLanguage
    If StrComp(currentLanguageCode, defaultLanguageCode, vbTextCompare) <> 0 Then
        translatedValue = LookupTranslation(defaultLanguageCode, normalizedKey)
        If LenB(translatedValue) > 0 Then
            T = translatedValue
            Exit Function
        End If
    End If

    If LenB(Fallback) > 0 Then
        T = Fallback
    Else
        T = TextKey
    End If
    Exit Function

ErrorHandler:
    If LenB(Fallback) > 0 Then
        T = Fallback
    Else
        T = TextKey
    End If
    modErrorHandler.HandleError MODULE_NAME, "T", Err
End Function

Private Sub EnsureTranslationStore()
    On Error GoTo ErrorHandler

    If mTranslations Is Nothing Then
        InitializeTranslations
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureTranslationStore", Err
End Sub

Private Sub LoadTranslations()
    On Error GoTo ErrorHandler

    Dim loadedCount As Long

    loadedCount = LoadTranslationsFromTable()
    If loadedCount > 0 Then
        modLoggingHandler.LogInfo MODULE_NAME & ".LoadTranslations", _
            CStr(loadedCount) & " translation(s) loaded from table '" & TABLE_FW_TRANSLATIONS & "'."
        Exit Sub
    End If

    LoadFallbackTranslations
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadTranslations", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function ResolveDefaultLanguage() As String
    On Error GoTo ErrorHandler

    Dim resolvedLanguage As String

    resolvedLanguage = vbNullString

    If LenB(resolvedLanguage) = 0 Then
        resolvedLanguage = NormalizeLanguageCode( _
            modConfigIni.GetConfigValue(CONFIG_SECTION_APP, "Language", vbNullString, ConfigFilePath))
    End If

    If LenB(resolvedLanguage) = 0 Then
        resolvedLanguage = FALLBACK_LANGUAGE
    End If

    ResolveDefaultLanguage = resolvedLanguage
    Exit Function

ErrorHandler:
    ResolveDefaultLanguage = FALLBACK_LANGUAGE
    modErrorHandler.HandleError MODULE_NAME, "ResolveDefaultLanguage", Err
End Function

Private Function LoadTranslationsFromTable() As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim SqlText As String
    Dim hasIsActiveField As Boolean

    If Not TranslationTableExists(TABLE_FW_TRANSLATIONS) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".LoadTranslationsFromTable", _
            "Translation table '" & TABLE_FW_TRANSLATIONS & "' not found. Falling back to minimal internal translations."
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Err.Raise vbObjectError + 2800, MODULE_NAME & ".LoadTranslationsFromTable", _
            "Current database could not be resolved."
    End If

    SqlText = "SELECT * FROM [" & TABLE_FW_TRANSLATIONS & "];"
    Set rs = db.OpenRecordset(SqlText, dbOpenSnapshot)

    If rs.EOF Then
        modLoggingHandler.LogWarning MODULE_NAME & ".LoadTranslationsFromTable", _
            "Translation table '" & TABLE_FW_TRANSLATIONS & "' contains no rows. Falling back to minimal internal translations."
        GoTo CleanExit
    End If

    If Not modDaoHelper.RecordsetHasField(rs, FIELD_TRANSLATION_KEY) _
        Or Not modDaoHelper.RecordsetHasField(rs, FIELD_LANGUAGE_CODE) _
        Or Not modDaoHelper.RecordsetHasField(rs, FIELD_TRANSLATION_VALUE) Then
        Err.Raise vbObjectError + 2801, MODULE_NAME & ".LoadTranslationsFromTable", _
            "Translation table '" & TABLE_FW_TRANSLATIONS & "' is missing one or more required fields."
    End If

    hasIsActiveField = modDaoHelper.RecordsetHasField(rs, FIELD_IS_ACTIVE)

    Do While Not rs.EOF
        If Not hasIsActiveField Or modDaoHelper.NzBoolean(rs.Fields(FIELD_IS_ACTIVE).Value, True) Then
            AddTranslation modDaoHelper.NzString(rs.Fields(FIELD_LANGUAGE_CODE).Value), _
                           modDaoHelper.NzString(rs.Fields(FIELD_TRANSLATION_KEY).Value), _
                           modDaoHelper.NzString(rs.Fields(FIELD_TRANSLATION_VALUE).Value)

            If LenB(NormalizeLanguageCode(modDaoHelper.NzString(rs.Fields(FIELD_LANGUAGE_CODE).Value))) > 0 _
                And LenB(NormalizeTextKey(modDaoHelper.NzString(rs.Fields(FIELD_TRANSLATION_KEY).Value))) > 0 Then
                LoadTranslationsFromTable = LoadTranslationsFromTable + 1
            End If
        End If
        rs.MoveNext
    Loop

    If LoadTranslationsFromTable = 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".LoadTranslationsFromTable", _
            "Translation table '" & TABLE_FW_TRANSLATIONS & "' contains no active translation rows. Falling back to minimal internal translations."
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
    End If
    Set rs = Nothing
    Set db = Nothing
    modErrorHandler.HandleError MODULE_NAME, "LoadTranslationsFromTable", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub LoadFallbackTranslations()
    On Error GoTo ErrorHandler

    AddTranslation "EN", "APP_TITLE", "Easis Version 4"
    AddTranslation "EN", "DOCUMENT", "Document"
    AddTranslation "EN", "CUSTOMER", "Customer"
    AddTranslation "EN", "TOTAL", "Total"

    AddTranslation "DE", "APP_TITLE", "Easis Version 4"
    AddTranslation "DE", "DOCUMENT", "Beleg"
    AddTranslation "DE", "CUSTOMER", "Kunde"
    AddTranslation "DE", "TOTAL", "Total"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadFallbackTranslations", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub AddTranslation(ByVal LanguageCode As String, ByVal TextKey As String, ByVal textValue As String)
    On Error GoTo ErrorHandler

    Dim compositeKey As String

    compositeKey = BuildTranslationKey(LanguageCode, TextKey)
    If LenB(compositeKey) = 0 Then
        Exit Sub
    End If

    EnsureTranslationStore
    mTranslations(compositeKey) = textValue
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "AddTranslation", Err
End Sub

Private Function LookupTranslation(ByVal LanguageCode As String, ByVal TextKey As String) As String
    On Error GoTo ErrorHandler

    Dim compositeKey As String

    compositeKey = BuildTranslationKey(LanguageCode, TextKey)
    If LenB(compositeKey) = 0 Then
        Exit Function
    End If

    If mTranslations.Exists(compositeKey) Then
        LookupTranslation = CStr(mTranslations(compositeKey))
    End If
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LookupTranslation", Err
End Function

Private Function BuildTranslationKey(ByVal LanguageCode As String, ByVal TextKey As String) As String
    Dim normalizedLanguage As String
    Dim normalizedTextKey As String

    normalizedLanguage = NormalizeLanguageCode(LanguageCode)
    normalizedTextKey = NormalizeTextKey(TextKey)

    If LenB(normalizedLanguage) = 0 Or LenB(normalizedTextKey) = 0 Then
        Exit Function
    End If

    BuildTranslationKey = normalizedLanguage & "|" & normalizedTextKey
End Function

Private Function NormalizeLanguageCode(ByVal LanguageCode As String) As String
    NormalizeLanguageCode = UCase$(Trim$(LanguageCode))
End Function

Private Function NormalizeTextKey(ByVal TextKey As String) As String
    NormalizeTextKey = UCase$(Trim$(TextKey))
End Function

Private Function TranslationTableExists(ByVal TableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tableDef As DAO.tableDef
    Dim normalizedTableName As String

    normalizedTableName = UCase$(Trim$(TableName))
    If LenB(normalizedTableName) = 0 Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    If db Is Nothing Then
        Err.Raise vbObjectError + 2802, MODULE_NAME & ".TranslationTableExists", _
            "Current database could not be resolved."
    End If

    For Each tableDef In db.TableDefs
        If UCase$(tableDef.Name) = normalizedTableName Then
            TranslationTableExists = True
            Exit For
        End If
    Next tableDef

    Set tableDef = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    TranslationTableExists = False
    modErrorHandler.HandleError MODULE_NAME, "TranslationTableExists", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function TEx(ByVal TextKey As String, ByVal Fallback As String, ParamArray Args() As Variant) As String
    On Error GoTo ErrorHandler

    Dim resultText As String
    Dim i As Long

    resultText = T(TextKey, Fallback)

    For i = LBound(Args) To UBound(Args)
        resultText = Replace(resultText, "{" & CStr(i) & "}", NzArgumentValue(Args(i)))
    Next i

    TEx = resultText
    Exit Function

ErrorHandler:
    TEx = T(TextKey, Fallback)
End Function

Private Function NzArgumentValue(ByVal Value As Variant) As String
    On Error GoTo SafeExit

    If IsNull(Value) Or IsEmpty(Value) Then
        NzArgumentValue = vbNullString
    Else
        NzArgumentValue = CStr(Value)
    End If
    Exit Function

SafeExit:
    NzArgumentValue = vbNullString
End Function