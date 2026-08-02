Attribute VB_Name = "modFwTranslationRuntime"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFwTranslationRuntime
' Purpose   : Resolves Tag-based TR markers at runtime for forms and provides
'             the current shared runtime base for future report-aware
'             localization.
' Author    : Codex
' Version   : 0.3.0
'===============================================================================

Private Const MODULE_NAME As String = "modFwTranslationRuntime"
Private Const TABLE_FW_TRANSLATIONS As String = "fw_translation"
Private Const FIELD_TRANSLATION_KEY As String = "translation_key"
Private Const FIELD_LANGUAGE_CODE As String = "language_code"
Private Const FIELD_TRANSLATION_VALUE As String = "translation_value"
Private Const TR_PREFIX As String = "TR:"
Private Const DEFAULT_LANGUAGE_CODE As String = "en-US"
Private Const FALLBACK_LANGUAGE_CODE As String = "en-US"
Private Const TENANT_PARAMETER_DEFAULT_LANGUAGE As String = "DEFAULT_LANGUAGE"
Private Const SUPPORTED_TRANSLATION_LANGUAGE_DE_CH As String = "de-CH"
Private Const SUPPORTED_TRANSLATION_LANGUAGE_FR_CH As String = "fr-CH"
Private Const SUPPORTED_TRANSLATION_LANGUAGE_EN_US As String = "en-US"
Private Const KNOWN_LANGUAGE_IT_CH As String = "it-CH"

Public Sub ApplyTranslations(ByVal TargetObject As Object)
    On Error GoTo ErrorHandler

    Dim languageCode As String
    Dim objectName As String
    Dim objectKind As String
    Dim resolvedCount As Long
    Dim missingCount As Long
    Dim ctl As Control

    ' Current runtime behavior is shared for forms and reports at a generic
    ' object level. A future phase may add report-aware composition and control
    ' scanning rules around REPORT.* keys, but that architecture is not fully
    ' implemented here yet.
    If TargetObject Is Nothing Then
        Exit Sub
    End If

    languageCode = GetCurrentLanguageCode()
    objectName = GetTargetObjectName(TargetObject)
    objectKind = GetTargetObjectKind(TargetObject)

    If Not ApplyTranslationToTarget(TargetObject, languageCode, resolvedCount, missingCount) Then
        If IsTransientUnavailableError(Err.Number) Then
            Err.Clear
            Exit Sub
        End If
    End If

    For Each ctl In TargetObject.Controls
        ApplyTranslationToTarget ctl, languageCode, resolvedCount, missingCount
        If IsTransientUnavailableError(Err.Number) Then
            Err.Clear
        End If
    Next ctl

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyTranslations", _
        "Translated " & objectKind & " '" & objectName & "' with " & _
        CStr(resolvedCount) & " resolved caption(s) and " & _
        CStr(missingCount) & " missing translation(s)."
    If TypeOf TargetObject Is Access.Form Then
        modFwDiagnostics.LogFormDiagnostics "AfterApplyTranslations:" & objectName, TargetObject
    End If
    Exit Sub

ErrorHandler:
    If IsTransientUnavailableError(Err.Number) Then
        Err.Clear
        Exit Sub
    End If
    modErrorHandler.HandleError MODULE_NAME, "ApplyTranslations", Err
End Sub

Public Function ResolveTranslation(ByVal translationKey As String, Optional ByVal languageCode As String = "") As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim originalValue As String
    Dim normalizedKey As String
    Dim normalizedLanguageCode As String
    Dim baseLanguageValue As String
    Dim translatedValue As String

    originalValue = NzString(translationKey)
    normalizedKey = NormalizeTranslationKey(translationKey)
    normalizedLanguageCode = NormalizeLanguageCode(languageCode)

    If LenB(normalizedKey) = 0 Then
        ResolveTranslation = originalValue
        Exit Function
    End If

    If LenB(normalizedLanguageCode) = 0 Then
        normalizedLanguageCode = GetCurrentLanguageCode()
    End If

    Set db = modDb.GetSystemDatabase()
    If db Is Nothing Then
        ResolveTranslation = originalValue
        Exit Function
    End If
    If Not modDbSchema.TableExists(db, TABLE_FW_TRANSLATIONS) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ResolveTranslation", _
            "Translation table not found: " & TABLE_FW_TRANSLATIONS & "."
        ResolveTranslation = originalValue
        Exit Function
    End If

    translatedValue = LookupTranslation(db, normalizedKey, normalizedLanguageCode)
    If LenB(translatedValue) > 0 Then
        ResolveTranslation = translatedValue
        Exit Function
    End If

    baseLanguageValue = GetBaseLanguageCode(normalizedLanguageCode)
    If LenB(baseLanguageValue) > 0 Then
        If StrComp(baseLanguageValue, normalizedLanguageCode, vbTextCompare) <> 0 Then
            translatedValue = LookupTranslation(db, normalizedKey, baseLanguageValue)
            If LenB(translatedValue) > 0 Then
                ResolveTranslation = translatedValue
                Exit Function
            End If
        End If
    End If

    If StrComp(normalizedLanguageCode, FALLBACK_LANGUAGE_CODE, vbTextCompare) <> 0 Then
        translatedValue = LookupTranslation(db, normalizedKey, FALLBACK_LANGUAGE_CODE)
        If LenB(translatedValue) > 0 Then
            ResolveTranslation = translatedValue
            Exit Function
        End If
    End If

    modLoggingHandler.LogWarning MODULE_NAME & ".ResolveTranslation", _
        "Missing translation for key '" & normalizedKey & "' in language '" & normalizedLanguageCode & "'."
    ResolveTranslation = originalValue
    Exit Function

ErrorHandler:
    ResolveTranslation = NzString(translationKey)
    modErrorHandler.HandleError MODULE_NAME, "ResolveTranslation", Err
End Function

Public Function ResolveCaptionText( _
    ByVal fallbackCaption As String, _
    ByVal tagText As String, _
    Optional ByVal languageCode As String = "") As String
    On Error GoTo ErrorHandler

    Dim translationKey As String
    Dim translatedValue As String

    fallbackCaption = NzString(fallbackCaption)
    translationKey = ResolveTranslationKeyFromMetadata(fallbackCaption, tagText)

    If LenB(translationKey) = 0 Then
        ResolveCaptionText = fallbackCaption
        Exit Function
    End If

    translatedValue = ResolveTranslation(translationKey, languageCode)
    If IsResolvedTranslationValue(translatedValue, translationKey) Then
        ResolveCaptionText = translatedValue
    Else
        ResolveCaptionText = fallbackCaption
    End If
    Exit Function

ErrorHandler:
    ResolveCaptionText = NzString(fallbackCaption)
    modErrorHandler.HandleError MODULE_NAME, "ResolveCaptionText", Err
End Function

Public Function ResolveText( _
    ByVal translationKey As String, _
    ByVal fallbackText As String, _
    Optional ByVal languageCode As String = "") As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim normalizedKey As String
    Dim normalizedLanguageCode As String
    Dim BaseLanguageCode As String
    Dim translatedValue As String

    normalizedKey = NormalizeTranslationKey(translationKey)
    normalizedLanguageCode = NormalizeLanguageCode(languageCode)

    If LenB(normalizedKey) = 0 Then
        ResolveText = NzString(fallbackText)
        Exit Function
    End If

    If LenB(normalizedLanguageCode) = 0 Then
        normalizedLanguageCode = GetCurrentLanguageCode()
    End If

    Set db = modDb.GetSystemDatabase()
    If db Is Nothing Then
        ResolveText = NzString(fallbackText)
        Exit Function
    End If
    If Not modDbSchema.TableExists(db, TABLE_FW_TRANSLATIONS) Then
        ResolveText = NzString(fallbackText)
        Exit Function
    End If

    translatedValue = LookupTranslation(db, normalizedKey, normalizedLanguageCode)
    If LenB(translatedValue) > 0 Then
        ResolveText = translatedValue
        Exit Function
    End If

    BaseLanguageCode = GetBaseLanguageCode(normalizedLanguageCode)
    If LenB(BaseLanguageCode) > 0 Then
        If StrComp(BaseLanguageCode, normalizedLanguageCode, vbTextCompare) <> 0 Then
            translatedValue = LookupTranslation(db, normalizedKey, BaseLanguageCode)
            If LenB(translatedValue) > 0 Then
                ResolveText = translatedValue
                Exit Function
            End If
        End If
    End If

    If StrComp(normalizedLanguageCode, FALLBACK_LANGUAGE_CODE, vbTextCompare) <> 0 Then
        translatedValue = LookupTranslation(db, normalizedKey, FALLBACK_LANGUAGE_CODE)
        If LenB(translatedValue) > 0 Then
            ResolveText = translatedValue
            Exit Function
        End If
    End If

    ResolveText = NzString(fallbackText)
    Exit Function

ErrorHandler:
    ResolveText = NzString(fallbackText)
    modErrorHandler.HandleError MODULE_NAME, "ResolveText", Err
End Function

Public Function GetCurrentLanguageCode() As String
    On Error GoTo ErrorHandler

    GetCurrentLanguageCode = ResolveCurrentLanguageCode()
    Exit Function

ErrorHandler:
    GetCurrentLanguageCode = DEFAULT_LANGUAGE_CODE
    modErrorHandler.HandleError MODULE_NAME, "GetCurrentLanguageCode", Err
End Function

Public Function ResolveCurrentLanguageCode() As String
    On Error GoTo ErrorHandler

    Dim resolvedLanguageCode As String

    resolvedLanguageCode = ResolveSupportedLanguageCode(ResolveCurrentUserLanguageCode())

    If LenB(resolvedLanguageCode) = 0 Then
        resolvedLanguageCode = ResolveSupportedLanguageCode( _
            modTenantRepository.GetTenantParameter(TENANT_PARAMETER_DEFAULT_LANGUAGE, vbNullString))
    End If

    If LenB(resolvedLanguageCode) = 0 Then
        resolvedLanguageCode = ResolveSupportedLanguageCode(ResolveIniLanguageCode())
    End If

    If LenB(resolvedLanguageCode) = 0 Then
        resolvedLanguageCode = DEFAULT_LANGUAGE_CODE
    End If

    ResolveCurrentLanguageCode = resolvedLanguageCode
    Exit Function

ErrorHandler:
    ResolveCurrentLanguageCode = DEFAULT_LANGUAGE_CODE
    modErrorHandler.HandleError MODULE_NAME, "ResolveCurrentLanguageCode", Err
End Function

Public Function GetTranslationKeyFromTag(ByVal tagText As String) As String
    Dim tagSegments() As String
    Dim segmentValue As Variant
    Dim normalizedSegment As String

    tagText = Trim$(NzString(tagText))
    If LenB(tagText) = 0 Then
        Exit Function
    End If

    tagSegments = Split(tagText, ";")
    For Each segmentValue In tagSegments
        normalizedSegment = Trim$(CStr(segmentValue))
        If LenB(normalizedSegment) = 0 Then
            GoTo NextSegment
        End If

        If StrComp(Left$(normalizedSegment, Len(TR_PREFIX)), TR_PREFIX, vbTextCompare) = 0 Then
            GetTranslationKeyFromTag = NormalizeTranslationKey(normalizedSegment)
            Exit Function
        End If

NextSegment:
    Next segmentValue
End Function

Public Function SetTranslationKeyInTag(ByVal tagText As String, ByVal translationKey As String) As String
    Dim cleanedTag As String
    Dim normalizedKey As String

    cleanedTag = RemoveTranslationKeyFromTag(tagText)
    normalizedKey = NormalizeTranslationKey(translationKey)

    If LenB(normalizedKey) = 0 Then
        SetTranslationKeyInTag = cleanedTag
    ElseIf LenB(cleanedTag) = 0 Then
        SetTranslationKeyInTag = TR_PREFIX & normalizedKey
    Else
        SetTranslationKeyInTag = cleanedTag & ";" & TR_PREFIX & normalizedKey
    End If
End Function

Public Function RemoveTranslationKeyFromTag(ByVal tagText As String) As String
    Dim tagSegments() As String
    Dim segmentValue As Variant
    Dim normalizedSegment As String
    Dim resultText As String

    tagText = Trim$(NzString(tagText))
    If LenB(tagText) = 0 Then
        Exit Function
    End If

    tagSegments = Split(tagText, ";")
    For Each segmentValue In tagSegments
        normalizedSegment = Trim$(CStr(segmentValue))
        If LenB(normalizedSegment) = 0 Then
            GoTo NextSegment
        End If

        If StrComp(Left$(normalizedSegment, Len(TR_PREFIX)), TR_PREFIX, vbTextCompare) = 0 Then
            GoTo NextSegment
        End If

        If LenB(resultText) > 0 Then
            resultText = resultText & ";"
        End If

        resultText = resultText & normalizedSegment

NextSegment:
    Next segmentValue

    RemoveTranslationKeyFromTag = resultText
End Function

Public Function BuildTranslatedReferenceRowSource( _
    ByVal referenceTableName As String, _
    ByVal codeFieldName As String, _
    Optional ByVal languageCode As String = "") As String
    On Error GoTo ErrorHandler

    Dim currentLanguageCode As String
    Dim sqlStatement As String

    currentLanguageCode = NormalizeLanguageCode(languageCode)
    If LenB(currentLanguageCode) = 0 Then
        currentLanguageCode = GetCurrentLanguageCode()
    End If

    sqlStatement = ""
    sqlStatement = sqlStatement & "SELECT "
    sqlStatement = sqlStatement & "r." & codeFieldName & ", "
    sqlStatement = sqlStatement & "IIf(t.translation_value Is Null, "
    sqlStatement = sqlStatement & "r." & codeFieldName & ", "
    sqlStatement = sqlStatement & "Left(t.translation_value, 255)) AS display_text "
    sqlStatement = sqlStatement & "FROM " & referenceTableName & " AS r "
    sqlStatement = sqlStatement & "LEFT JOIN fw_translation AS t "
    sqlStatement = sqlStatement & "ON r.translation_key = t.translation_key "
    sqlStatement = sqlStatement & "WHERE Nz(r.is_active, True) = True "
    sqlStatement = sqlStatement & "AND (t.language_code = " & SqlText(currentLanguageCode) & " "
    sqlStatement = sqlStatement & "OR t.language_code Is Null) "
    sqlStatement = sqlStatement & "AND (Nz(t.is_active, True) = True "
    sqlStatement = sqlStatement & "OR t.translation_key Is Null) "
    sqlStatement = sqlStatement & "ORDER BY r.sort_order, r." & codeFieldName & ";"

    BuildTranslatedReferenceRowSource = sqlStatement
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "BuildTranslatedReferenceRowSource", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Private Function LookupTranslation( _
    ByVal db As DAO.Database, _
    ByVal translationKey As String, _
    ByVal languageCode As String) As String
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlStatement As String

    If db Is Nothing Then
        Exit Function
    End If

    sqlStatement = "SELECT TOP 1 [" & FIELD_TRANSLATION_VALUE & "] " & _
                   "FROM [" & TABLE_FW_TRANSLATIONS & "] " & _
                   "WHERE [" & FIELD_TRANSLATION_KEY & "] = " & SqlText(translationKey) & " " & _
                   "AND [" & FIELD_LANGUAGE_CODE & "] = " & SqlText(languageCode) & " " & _
                   "AND Trim(Nz([" & FIELD_TRANSLATION_VALUE & "], '')) <> ''"

    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)

    If Not (rs.BOF And rs.EOF) Then
        LookupTranslation = NzString(rs.Fields(0).Value)
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LookupTranslation", Err
    Resume CleanExit
End Function

Private Function ResolveCurrentUserLanguageCode() As String
    On Error GoTo ErrorHandler

    Dim UserId As String

    If Not modSessionContext.IsSessionInitialized Then
        Exit Function
    End If

    UserId = Trim$(modSessionContext.currentUserId)
    If LenB(UserId) = 0 Then
        Exit Function
    End If

    ResolveCurrentUserLanguageCode = Trim$(modUserRepository.GetUserLanguageCode(UserId, vbNullString))
    Exit Function

ErrorHandler:
    ResolveCurrentUserLanguageCode = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "ResolveCurrentUserLanguageCode", Err
End Function

Private Function ResolveIniLanguageCode() As String
    On Error GoTo ErrorHandler

    ResolveIniLanguageCode = Trim$(modConfigIni.GetConfigValue(CONFIG_SECTION_APP, "Language", CurrentLanguage, ConfigFilePath))
    Exit Function

ErrorHandler:
    ResolveIniLanguageCode = Trim$(CurrentLanguage)
    modErrorHandler.HandleError MODULE_NAME, "ResolveIniLanguageCode", Err
End Function

Private Function ResolveTranslationKeyFromMetadata( _
    ByVal fallbackCaption As String, _
    ByVal tagText As String) As String

    ResolveTranslationKeyFromMetadata = GetTranslationKeyFromTag(tagText)
    If LenB(ResolveTranslationKeyFromMetadata) > 0 Then
        Exit Function
    End If

    If HasTranslationPrefix(fallbackCaption) Then
        ResolveTranslationKeyFromMetadata = NormalizeTranslationKey(fallbackCaption)
    End If
End Function

Private Function NormalizeTranslationKey(ByVal translationKey As String) As String
    translationKey = Trim$(NzString(translationKey))

    If LenB(translationKey) = 0 Then
        Exit Function
    End If

    If StrComp(Left$(translationKey, Len(TR_PREFIX)), TR_PREFIX, vbTextCompare) = 0 Then
        NormalizeTranslationKey = Trim$(Mid$(translationKey, Len(TR_PREFIX) + 1))
    Else
        NormalizeTranslationKey = translationKey
    End If
End Function

Private Function NormalizeLanguageCode(ByVal languageCode As String) As String
    Dim parts() As String

    languageCode = Trim$(NzString(languageCode))
    If LenB(languageCode) = 0 Then
        Exit Function
    End If

    languageCode = Replace(languageCode, "_", "-")
    parts = Split(languageCode, "-")

    If UBound(parts) = 0 Then
        NormalizeLanguageCode = NormalizeLegacyLanguageAlias(LCase$(parts(0)))
    Else
        NormalizeLanguageCode = NormalizeLegacyLanguageAlias(LCase$(parts(0)) & "-" & UCase$(parts(1)))
    End If
End Function

Private Function ResolveSupportedLanguageCode(ByVal languageCode As String) As String
    Dim normalizedLanguageCode As String

    normalizedLanguageCode = NormalizeLanguageCode(languageCode)

    Select Case normalizedLanguageCode
        Case SUPPORTED_TRANSLATION_LANGUAGE_DE_CH, SUPPORTED_TRANSLATION_LANGUAGE_EN_US, SUPPORTED_TRANSLATION_LANGUAGE_FR_CH
            ResolveSupportedLanguageCode = normalizedLanguageCode
    End Select
End Function

Public Function GetSupportedTranslationLanguages() As Variant
    GetSupportedTranslationLanguages = Array( _
        SUPPORTED_TRANSLATION_LANGUAGE_DE_CH, _
        SUPPORTED_TRANSLATION_LANGUAGE_EN_US, _
        SUPPORTED_TRANSLATION_LANGUAGE_FR_CH)
End Function

Public Function GetKnownLanguageCodes() As Variant
    GetKnownLanguageCodes = Array( _
        SUPPORTED_TRANSLATION_LANGUAGE_DE_CH, _
        SUPPORTED_TRANSLATION_LANGUAGE_FR_CH, _
        SUPPORTED_TRANSLATION_LANGUAGE_EN_US, _
        KNOWN_LANGUAGE_IT_CH)
End Function

Public Function IsSupportedTranslationLanguage(ByVal languageCode As String) As Boolean
    Select Case NormalizeLanguageCode(languageCode)
        Case SUPPORTED_TRANSLATION_LANGUAGE_DE_CH, SUPPORTED_TRANSLATION_LANGUAGE_EN_US, SUPPORTED_TRANSLATION_LANGUAGE_FR_CH
            IsSupportedTranslationLanguage = True
    End Select
End Function

Public Function NormalizeProjectLanguageCode(ByVal languageCode As String) As String
    NormalizeProjectLanguageCode = NormalizeLanguageCode(languageCode)
End Function

Private Function NormalizeLegacyLanguageAlias(ByVal normalizedLanguageCode As String) As String
    Select Case normalizedLanguageCode
        Case "de", "de-CH", "de-DE"
            If StrComp(normalizedLanguageCode, "de-DE", vbTextCompare) = 0 Then
                NormalizeLegacyLanguageAlias = "de-DE"
            Else
                NormalizeLegacyLanguageAlias = SUPPORTED_TRANSLATION_LANGUAGE_DE_CH
            End If
        Case "fr", "fr-FR", "fr-CH"
            NormalizeLegacyLanguageAlias = SUPPORTED_TRANSLATION_LANGUAGE_FR_CH
        Case "en", "en-US"
            NormalizeLegacyLanguageAlias = SUPPORTED_TRANSLATION_LANGUAGE_EN_US
        Case "it", "it-CH"
            NormalizeLegacyLanguageAlias = KNOWN_LANGUAGE_IT_CH
        Case Else
            NormalizeLegacyLanguageAlias = normalizedLanguageCode
    End Select
End Function

Private Function GetBaseLanguageCode(ByVal languageCode As String) As String
    Dim normalizedLanguageCode As String
    Dim separatorPosition As Long

    normalizedLanguageCode = NormalizeLanguageCode(languageCode)
    separatorPosition = InStr(1, normalizedLanguageCode, "-", vbBinaryCompare)

    If separatorPosition > 0 Then
        GetBaseLanguageCode = Left$(normalizedLanguageCode, separatorPosition - 1)
    Else
        GetBaseLanguageCode = normalizedLanguageCode
    End If
End Function

Private Function HasAnyTranslationMarker(ByVal CaptionValue As String, ByVal tagText As String) As Boolean
    HasAnyTranslationMarker = HasTranslationPrefix(CaptionValue) Or LenB(GetTranslationKeyFromTag(tagText)) > 0
End Function

Private Function HasTranslationPrefix(ByVal CaptionValue As String) As Boolean
    CaptionValue = Trim$(NzString(CaptionValue))
    If LenB(CaptionValue) = 0 Then
        Exit Function
    End If

    HasTranslationPrefix = (StrComp(Left$(CaptionValue, Len(TR_PREFIX)), TR_PREFIX, vbTextCompare) = 0)
End Function

Private Function IsResolvedTranslationValue( _
    ByVal translatedValue As String, _
    ByVal translationKey As String) As Boolean

    translatedValue = Trim$(NzString(translatedValue))
    translationKey = Trim$(NormalizeTranslationKey(translationKey))

    If LenB(translatedValue) = 0 Then
        Exit Function
    End If

    If LenB(translationKey) = 0 Then
        Exit Function
    End If

    If StrComp(translatedValue, translationKey, vbTextCompare) = 0 Then
        Exit Function
    End If

    If StrComp(translatedValue, TR_PREFIX & translationKey, vbTextCompare) = 0 Then
        Exit Function
    End If

    IsResolvedTranslationValue = True
End Function

Private Function GetCaptionValue(ByVal target As Object) As String
    On Error GoTo SafeExit

    GetCaptionValue = NzString(target.Properties("Caption").Value)
    Exit Function

SafeExit:
    GetCaptionValue = vbNullString
End Function

Private Function GetTagValue(ByVal target As Object) As String
    On Error GoTo SafeExit

    GetTagValue = NzString(target.Properties("Tag").Value)
    Exit Function

SafeExit:
    GetTagValue = vbNullString
End Function

Private Function SetCaptionValue(ByVal target As Object, ByVal CaptionValue As String) As Boolean
    On Error GoTo SafeExit

    target.Properties("Caption").Value = CaptionValue
    SetCaptionValue = True
    Exit Function

SafeExit:
    SetCaptionValue = False
End Function

Private Function ApplyTranslationToTarget( _
    ByVal target As Object, _
    ByVal languageCode As String, _
    ByRef resolvedCount As Long, _
    ByRef missingCount As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim rawCaption As String
    Dim translatedCaption As String
    Dim tagText As String

    rawCaption = GetCaptionValue(target)
    tagText = GetTagValue(target)
    translatedCaption = ResolveCaptionText(rawCaption, tagText, languageCode)

    If StrComp(translatedCaption, rawCaption, vbBinaryCompare) <> 0 Then
        If SetCaptionValue(target, translatedCaption) Then
            resolvedCount = resolvedCount + 1
        End If
    ElseIf HasAnyTranslationMarker(rawCaption, tagText) Then
        missingCount = missingCount + 1
    End If

    ApplyTranslationToTarget = True
    Exit Function

ErrorHandler:
    If IsTransientUnavailableError(Err.Number) Then
        Exit Function
    End If
    modErrorHandler.HandleError MODULE_NAME, "ApplyTranslationToTarget", Err
End Function

Private Function IsTransientUnavailableError(ByVal errNumber As Long) As Boolean
    Select Case errNumber
        Case 2467, 2455
            IsTransientUnavailableError = True
    End Select
End Function

Private Function GetTargetObjectName(ByVal TargetObject As Object) As String
    On Error GoTo SafeExit

    GetTargetObjectName = NzString(TargetObject.Name)
    Exit Function

SafeExit:
    GetTargetObjectName = "(unknown)"
End Function

Private Function GetTargetObjectKind(ByVal TargetObject As Object) As String
    Dim typeNameValue As String

    typeNameValue = UCase$(TypeName(TargetObject))

    Select Case typeNameValue
        Case "FORM"
            GetTargetObjectKind = "form"
        Case "REPORT"
            GetTargetObjectKind = "report"
        Case Else
            GetTargetObjectKind = LCase$(typeNameValue)
    End Select
End Function


Private Function NzString(ByVal Value As Variant, Optional ByVal defaultValue As String = "") As String
    If IsNull(Value) Or IsEmpty(Value) Then
        NzString = defaultValue
    Else
        NzString = CStr(Value)
    End If
End Function



