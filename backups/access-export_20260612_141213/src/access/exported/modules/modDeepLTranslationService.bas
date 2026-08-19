Attribute VB_Name = "modDeepLTranslationService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modDeepLTranslationService
' Purpose   : Provides DeepL-based translation suggestions for the focused
'             translation edit workflow without persisting values automatically.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modDeepLTranslationService"

Private Const DEFAULT_DEEPL_API_BASE_URL As String = "https://api-free.deepl.com"
Private Const API_TRANSLATE_PATH As String = "/v2/translate"

Private Const LANGUAGE_DE_CH As String = "DE-CH"
Private Const LANGUAGE_EN_US As String = "EN-US"
Private Const LANGUAGE_FR_FR As String = "FR-FR"
Private Const LANGUAGE_DE As String = "DE"
Private Const LANGUAGE_EN As String = "EN"
Private Const LANGUAGE_FR As String = "FR"

Public Function ResolveDeepLApiKey() As String
    On Error GoTo ErrorHandler

    ResolveDeepLApiKey = Trim$(modTenantRepository.GetTenantParameter(TENANT_PARAMETER_DEEPL_API_KEY, vbNullString))
    Exit Function

ErrorHandler:
    ResolveDeepLApiKey = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "ResolveDeepLApiKey", Err
End Function

Public Function ResolveDeepLBaseUrl() As String
    On Error GoTo ErrorHandler

    Dim baseUrl As String

    baseUrl = Trim$(modTenantRepository.GetTenantParameter(TENANT_PARAMETER_DEEPL_API_BASE_URL, DEFAULT_DEEPL_API_BASE_URL))
    If LenB(baseUrl) = 0 Then
        baseUrl = DEFAULT_DEEPL_API_BASE_URL
    End If

    If Right$(baseUrl, 1) = "/" Then
        baseUrl = Left$(baseUrl, Len(baseUrl) - 1)
    End If

    ResolveDeepLBaseUrl = baseUrl
    Exit Function

ErrorHandler:
    ResolveDeepLBaseUrl = DEFAULT_DEEPL_API_BASE_URL
    modErrorHandler.HandleError MODULE_NAME, "ResolveDeepLBaseUrl", Err
End Function

Public Function TranslateText( _
    ByVal sourceText As String, _
    ByVal sourceLang As String, _
    ByVal targetLang As String, _
    Optional ByVal contextText As String = "") As String
    On Error GoTo ErrorHandler

    Dim apiKey As String
    Dim baseUrl As String
    Dim requestUrl As String
    Dim requestBody As String
    Dim http As Object
    Dim mappedSourceLang As String
    Dim mappedTargetLang As String
    Dim jsonText As String

    sourceText = Trim$(modDaoHelper.NzString(sourceText))
    mappedSourceLang = ResolveDeepLLanguageCode(sourceLang, False)
    mappedTargetLang = ResolveDeepLLanguageCode(targetLang, True)

    If LenB(sourceText) = 0 Then
        Exit Function
    End If

    apiKey = ResolveDeepLApiKey()
    If LenB(apiKey) = 0 Then
        Err.Raise vbObjectError + 5900, MODULE_NAME & ".TranslateText", "DEEPL_API_KEY is not configured."
    End If

    If LenB(mappedSourceLang) = 0 Then
        Err.Raise vbObjectError + 5901, MODULE_NAME & ".TranslateText", "Unsupported DeepL source language: " & sourceLang
    End If

    If LenB(mappedTargetLang) = 0 Then
        Err.Raise vbObjectError + 5902, MODULE_NAME & ".TranslateText", "Unsupported DeepL target language: " & targetLang
    End If

    baseUrl = ResolveDeepLBaseUrl()
    requestUrl = baseUrl & API_TRANSLATE_PATH
    requestBody = BuildTranslateRequestBody(sourceText, mappedSourceLang, mappedTargetLang, contextText)

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", requestUrl, False
    http.SetTimeouts 10000, 10000, 30000, 30000
    http.SetRequestHeader "Authorization", "DeepL-Auth-Key " & apiKey
    http.SetRequestHeader "Content-Type", "application/json"
    http.Send requestBody

    jsonText = DecodeUtf8ResponseBody(http.responseBody)

    If CLng(http.status) < 200 Or CLng(http.status) >= 300 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".TranslateText", _
            "DeepL request failed. status=" & CStr(http.status) & _
            "; url=" & requestUrl & "; response=" & jsonText
        Err.Raise vbObjectError + 5903, MODULE_NAME & ".TranslateText", _
            "DeepL request failed with status " & CStr(http.status) & "."
    End If

    TranslateText = ExtractTranslatedText(jsonText)

    modLoggingHandler.LogInfo MODULE_NAME & ".TranslateText", _
        "DeepL suggestion created for " & mappedSourceLang & " -> " & mappedTargetLang & "."

CleanExit:
    Set http = Nothing
    Exit Function

ErrorHandler:
    Set http = Nothing
    modErrorHandler.HandleError MODULE_NAME, "TranslateText", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Public Function SuggestMissingTranslations( _
    ByVal sourceText As String, _
    Optional ByVal sourceLang As String = LANGUAGE_DE_CH, _
    Optional ByVal existingEnUs As String = "", _
    Optional ByVal existingFrFr As String = "", _
    Optional ByVal contextText As String = "", _
    Optional ByVal overwriteFilledTargets As Boolean = False) As Object
    On Error GoTo ErrorHandler

    Dim suggestions As Object

    Set suggestions = CreateObject("Scripting.Dictionary")
    sourceText = Trim$(modDaoHelper.NzString(sourceText))

    If LenB(sourceText) = 0 Then
        Set SuggestMissingTranslations = suggestions
        Exit Function
    End If

    If ShouldSuggestTarget(existingEnUs, overwriteFilledTargets) Then
        suggestions.Add LANGUAGE_EN_US, TranslateText(sourceText, sourceLang, LANGUAGE_EN_US, contextText)
    End If

    If ShouldSuggestTarget(existingFrFr, overwriteFilledTargets) Then
        suggestions.Add LANGUAGE_FR_FR, TranslateText(sourceText, sourceLang, LANGUAGE_FR_FR, contextText)
    End If

    Set SuggestMissingTranslations = suggestions
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SuggestMissingTranslations", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Private Function BuildTranslateRequestBody( _
    ByVal sourceText As String, _
    ByVal sourceLang As String, _
    ByVal targetLang As String, _
    ByVal contextText As String) As String

    BuildTranslateRequestBody = "{""text"":[""" & JsonEscape(sourceText) & """]"
    BuildTranslateRequestBody = BuildTranslateRequestBody & ",""source_lang"":""" & JsonEscape(sourceLang) & """"
    BuildTranslateRequestBody = BuildTranslateRequestBody & ",""target_lang"":""" & JsonEscape(targetLang) & """"

    contextText = Trim$(modDaoHelper.NzString(contextText))
    If LenB(contextText) > 0 Then
        BuildTranslateRequestBody = BuildTranslateRequestBody & ",""context"":""" & JsonEscape(contextText) & """"
    End If

    BuildTranslateRequestBody = BuildTranslateRequestBody & "}"
End Function

Private Function ResolveDeepLLanguageCode(ByVal languageCode As String, ByVal isTargetLanguage As Boolean) As String
    languageCode = UCase$(Trim$(modDaoHelper.NzString(languageCode)))

    Select Case languageCode
        Case LANGUAGE_DE_CH, "DE-DE", LANGUAGE_DE
            ResolveDeepLLanguageCode = LANGUAGE_DE
        Case LANGUAGE_EN_US
            If isTargetLanguage Then
                ResolveDeepLLanguageCode = LANGUAGE_EN_US
            Else
                ResolveDeepLLanguageCode = LANGUAGE_EN
            End If
        Case "EN-GB", LANGUAGE_EN
            ResolveDeepLLanguageCode = LANGUAGE_EN
        Case LANGUAGE_FR_FR, "FR-CH", LANGUAGE_FR
            ResolveDeepLLanguageCode = LANGUAGE_FR
        Case Else
            ResolveDeepLLanguageCode = vbNullString
    End Select
End Function

Private Function ShouldSuggestTarget(ByVal existingValue As String, ByVal overwriteFilledTargets As Boolean) As Boolean
    existingValue = Trim$(modDaoHelper.NzString(existingValue))
    ShouldSuggestTarget = overwriteFilledTargets Or (LenB(existingValue) = 0)
End Function

Private Function ExtractTranslatedText(ByVal responseText As String) As String
    On Error GoTo ErrorHandler

    Dim regEx As Object
    Dim matches As Object

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = False
    regEx.IgnoreCase = True
    regEx.MultiLine = True
    regEx.Pattern = """text""\s*:\s*""((?:\\.|[^""\\])*)"""

    Set matches = regEx.Execute(responseText)
    If matches.count = 0 Then
        Err.Raise vbObjectError + 5904, MODULE_NAME & ".ExtractTranslatedText", "DeepL response did not contain a translation text."
    End If

    ExtractTranslatedText = JsonUnescape(matches(0).SubMatches(0))

CleanExit:
    Set matches = Nothing
    Set regEx = Nothing
    Exit Function

ErrorHandler:
    Set matches = Nothing
    Set regEx = Nothing
    modErrorHandler.HandleError MODULE_NAME, "ExtractTranslatedText", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Private Function DecodeUtf8ResponseBody(ByVal responseBody As Variant) As String
    On Error GoTo ErrorHandler

    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write responseBody
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "utf-8"
    DecodeUtf8ResponseBody = stream.ReadText

CleanExit:
    On Error Resume Next
    If Not stream Is Nothing Then
        If stream.State <> 0 Then
            stream.Close
        End If
    End If
    Set stream = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "DecodeUtf8ResponseBody", Err
    Resume CleanExit
End Function

Private Function JsonEscape(ByVal valueText As String) As String
    valueText = modDaoHelper.NzString(valueText)
    valueText = Replace(valueText, "\", "\\")
    valueText = Replace(valueText, """", Chr$(92) & Chr$(34))
    valueText = Replace(valueText, "/", "\/")
    valueText = Replace(valueText, vbBack, "\b")
    valueText = Replace(valueText, Chr$(12), "\f")
    valueText = Replace(valueText, vbCrLf, "\n")
    valueText = Replace(valueText, vbCr, "\n")
    valueText = Replace(valueText, vbLf, "\n")
    valueText = Replace(valueText, vbTab, "\t")
    JsonEscape = valueText
End Function

Private Function JsonUnescape(ByVal valueText As String) As String
    valueText = Replace(valueText, "\/", "/")
    valueText = Replace(valueText, Chr$(92) & Chr$(34), Chr$(34))
    valueText = Replace(valueText, "\\", "\")
    valueText = Replace(valueText, "\r\n", vbCrLf)
    valueText = Replace(valueText, "\n", vbCrLf)
    valueText = Replace(valueText, "\r", vbCr)
    valueText = Replace(valueText, "\t", vbTab)
    valueText = Replace(valueText, "\b", vbBack)
    valueText = Replace(valueText, "\f", Chr$(12))
    JsonUnescape = valueText
End Function