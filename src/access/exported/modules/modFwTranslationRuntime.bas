Attribute VB_Name = "modFwTranslationRuntime"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFwTranslationRuntime
' Purpose   : Resolves TR:* captions at runtime for forms and reports.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modFwTranslationRuntime"
Private Const TABLE_FW_TRANSLATIONS As String = "tblFwTranslations"
Private Const FIELD_TRANSLATION_KEY As String = "TranslationKey"
Private Const FIELD_LANGUAGE_CODE As String = "LanguageCode"
Private Const FIELD_TRANSLATION_VALUE As String = "TranslationValue"
Private Const TR_PREFIX As String = "TR:"
Private Const DEFAULT_LANGUAGE_CODE As String = "DE-CH"
Private Const FALLBACK_LANGUAGE_CODE As String = "EN"

Public Sub ApplyTranslations(ByVal TargetObject As Object)
    On Error GoTo ErrorHandler

    Dim languageCode As String
    Dim objectName As String
    Dim objectKind As String
    Dim resolvedCount As Long
    Dim missingCount As Long
    Dim rawCaption As String
    Dim translatedCaption As String
    Dim ctl As Control

    If TargetObject Is Nothing Then
        Exit Sub
    End If

    languageCode = GetCurrentLanguageCode()
    objectName = GetTargetObjectName(TargetObject)
    objectKind = GetTargetObjectKind(TargetObject)

    rawCaption = GetCaptionValue(TargetObject)
    If HasTranslationPrefix(rawCaption) Then
        translatedCaption = ResolveTranslation(rawCaption, languageCode)
        If StrComp(translatedCaption, rawCaption, vbBinaryCompare) <> 0 Then
            If SetCaptionValue(TargetObject, translatedCaption) Then
                resolvedCount = resolvedCount + 1
            End If
        Else
            missingCount = missingCount + 1
        End If
    End If

    For Each ctl In TargetObject.Controls
        rawCaption = GetCaptionValue(ctl)

        If Not HasTranslationPrefix(rawCaption) Then
            GoTo NextControl
        End If

        translatedCaption = ResolveTranslation(rawCaption, languageCode)
        If StrComp(translatedCaption, rawCaption, vbBinaryCompare) <> 0 Then
            If SetCaptionValue(ctl, translatedCaption) Then
                resolvedCount = resolvedCount + 1
            End If
        Else
            missingCount = missingCount + 1
        End If

NextControl:
    Next ctl

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyTranslations", _
        "Translated " & objectKind & " '" & objectName & "' with " & _
        CStr(resolvedCount) & " resolved caption(s) and " & _
        CStr(missingCount) & " missing translation(s)."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyTranslations", Err
End Sub

Public Function ResolveTranslation(ByVal TranslationKey As String, Optional ByVal LanguageCode As String = "") As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim originalValue As String
    Dim normalizedKey As String
    Dim normalizedLanguageCode As String
    Dim baseLanguageCode As String
    Dim translatedValue As String

    originalValue = NzString(TranslationKey)
    normalizedKey = NormalizeTranslationKey(TranslationKey)
    normalizedLanguageCode = NormalizeLanguageCode(LanguageCode)

    If LenB(normalizedKey) = 0 Then
        ResolveTranslation = originalValue
        Exit Function
    End If

    If LenB(normalizedLanguageCode) = 0 Then
        normalizedLanguageCode = GetCurrentLanguageCode()
    End If

    Set db = CurrentDb
    If Not TableExists(db, TABLE_FW_TRANSLATIONS) Then
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

    baseLanguageCode = BaseLanguageCode(normalizedLanguageCode)
    If LenB(baseLanguageCode) > 0 Then
        If StrComp(baseLanguageCode, normalizedLanguageCode, vbTextCompare) <> 0 Then
            translatedValue = LookupTranslation(db, normalizedKey, baseLanguageCode)
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
    ResolveTranslation = NzString(TranslationKey)
    modErrorHandler.HandleError MODULE_NAME, "ResolveTranslation", Err
End Function

Public Function GetCurrentLanguageCode() As String
    On Error GoTo ErrorHandler

    GetCurrentLanguageCode = DEFAULT_LANGUAGE_CODE
    Exit Function

ErrorHandler:
    GetCurrentLanguageCode = DEFAULT_LANGUAGE_CODE
    modErrorHandler.HandleError MODULE_NAME, "GetCurrentLanguageCode", Err
End Function

Private Function LookupTranslation( _
    ByVal db As DAO.Database, _
    ByVal TranslationKey As String, _
    ByVal LanguageCode As String) As String
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlText As String

    If db Is Nothing Then
        Exit Function
    End If

    sqlText = "SELECT TOP 1 [" & FIELD_TRANSLATION_VALUE & "] " & _
              "FROM [" & TABLE_FW_TRANSLATIONS & "] " & _
              "WHERE [" & FIELD_TRANSLATION_KEY & "] = " & SqlText(TranslationKey) & " " & _
              "AND [" & FIELD_LANGUAGE_CODE & "] = " & SqlText(LanguageCode) & " " & _
              "AND Trim(Nz([" & FIELD_TRANSLATION_VALUE & "], '')) <> ''"

    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
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

Private Function NormalizeTranslationKey(ByVal TranslationKey As String) As String
    TranslationKey = Trim$(NzString(TranslationKey))

    If LenB(TranslationKey) = 0 Then
        Exit Function
    End If

    If StrComp(Left$(TranslationKey, Len(TR_PREFIX)), TR_PREFIX, vbTextCompare) = 0 Then
        NormalizeTranslationKey = Trim$(Mid$(TranslationKey, Len(TR_PREFIX) + 1))
    Else
        NormalizeTranslationKey = TranslationKey
    End If
End Function

Private Function NormalizeLanguageCode(ByVal LanguageCode As String) As String
    Dim parts() As String

    LanguageCode = Trim$(NzString(LanguageCode))
    If LenB(LanguageCode) = 0 Then
        Exit Function
    End If

    LanguageCode = Replace(LanguageCode, "_", "-")
    parts = Split(LanguageCode, "-")

    If UBound(parts) = 0 Then
        NormalizeLanguageCode = UCase$(parts(0))
    Else
        NormalizeLanguageCode = UCase$(parts(0)) & "-" & UCase$(parts(1))
    End If
End Function

Private Function BaseLanguageCode(ByVal LanguageCode As String) As String
    Dim normalizedLanguageCode As String
    Dim separatorPosition As Long

    normalizedLanguageCode = NormalizeLanguageCode(LanguageCode)
    separatorPosition = InStr(1, normalizedLanguageCode, "-", vbBinaryCompare)

    If separatorPosition > 0 Then
        BaseLanguageCode = Left$(normalizedLanguageCode, separatorPosition - 1)
    Else
        BaseLanguageCode = normalizedLanguageCode
    End If
End Function

Private Function HasTranslationPrefix(ByVal CaptionValue As String) As Boolean
    CaptionValue = Trim$(NzString(CaptionValue))
    If LenB(CaptionValue) = 0 Then
        Exit Function
    End If

    HasTranslationPrefix = (StrComp(Left$(CaptionValue, Len(TR_PREFIX)), TR_PREFIX, vbTextCompare) = 0)
End Function

Private Function GetCaptionValue(ByVal Target As Object) As String
    On Error GoTo SafeExit

    GetCaptionValue = NzString(Target.Properties("Caption").Value)
    Exit Function

SafeExit:
    GetCaptionValue = vbNullString
End Function

Private Function SetCaptionValue(ByVal Target As Object, ByVal CaptionValue As String) As Boolean
    On Error GoTo SafeExit

    Target.Properties("Caption").Value = CaptionValue
    SetCaptionValue = True
    Exit Function

SafeExit:
    SetCaptionValue = False
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

Private Function TableExists(ByVal db As DAO.Database, ByVal TableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tdf As DAO.TableDef

    If db Is Nothing Then
        Exit Function
    End If

    For Each tdf In db.TableDefs
        If StrComp(tdf.Name, TableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next tdf

    Exit Function

ErrorHandler:
    TableExists = False
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
