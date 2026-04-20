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

Private mTranslations As Object
Private mCurrentLanguage As String
Private mDefaultLanguage As String

Public Sub InitializeTranslations()
    On Error GoTo ErrorHandler

    Set mTranslations = CreateObject("Scripting.Dictionary")
    mTranslations.CompareMode = vbTextCompare

    mDefaultLanguage = ResolveDefaultLanguage()
    mCurrentLanguage = mDefaultLanguage

    LoadStubTranslations

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

    defaultLanguageCode = ResolveDefaultLanguage()
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

Private Sub LoadStubTranslations()
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
    modErrorHandler.HandleError MODULE_NAME, "LoadStubTranslations", Err
End Sub

Private Sub AddTranslation(ByVal LanguageCode As String, ByVal TextKey As String, ByVal TextValue As String)
    On Error GoTo ErrorHandler

    Dim compositeKey As String

    compositeKey = BuildTranslationKey(LanguageCode, TextKey)
    If LenB(compositeKey) = 0 Then
        Exit Sub
    End If

    EnsureTranslationStore
    mTranslations(compositeKey) = TextValue
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
