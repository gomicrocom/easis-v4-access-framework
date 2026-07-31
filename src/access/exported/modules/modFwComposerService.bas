Attribute VB_Name = "modFwComposerService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFwComposerService
' Purpose   : Provides object and control inspection helpers for translation
'             maintenance and Tag-based translation key management.
' Author    : Codex
' Version   : 0.2.0
'===============================================================================

Private Const MODULE_NAME As String = "modFwComposerService"
Private Const TRANSLATION_TABLE_NAME As String = "fw_translation"
Private Const PLACEHOLDER_TRANSLATION_VALUE As String = "<neu>"
Private Const FIELD_TRANSLATION_KEY As String = "translation_key"
Private Const FIELD_LANGUAGE_CODE As String = "language_code"
Private Const FIELD_TRANSLATION_VALUE As String = "translation_value"
Private Const FIELD_IS_ACTIVE As String = "is_active"
Private Const FIELD_MODULE_CODE As String = "module_code"
Private Const FIELD_UPDATED_AT As String = "updated_at"

Public Const COMPOSER_MODE_TAGS As String = "TAGS"
Public Const COMPOSER_MODE_TRANSLATIONS As String = "TRANSLATIONS"
Public Const OBJECT_TYPE_FORM As String = "FORM"
Public Const OBJECT_TYPE_REPORT As String = "REPORT"

Public Function GetComposerObjectList(ByVal objectType As String) As Collection
    On Error GoTo ErrorHandler

    Dim normalizedObjectType As String
    Dim result As Collection
    Dim accessObject As Access.accessObject

    normalizedObjectType = NormalizeObjectType(objectType)
    Set result = New Collection

    Select Case normalizedObjectType
        Case OBJECT_TYPE_FORM
            For Each accessObject In CurrentProject.AllForms
                If Left$(accessObject.Name, 1) <> "~" Then
                    If Not IsComposerInternalObject(accessObject.Name) Then
                        CollectionAddSorted result, accessObject.Name
                    End If
                End If
            Next accessObject

        Case OBJECT_TYPE_REPORT
            For Each accessObject In CurrentProject.AllReports
                If Left$(accessObject.Name, 1) <> "~" Then
                    If Not IsComposerInternalObject(accessObject.Name) Then
                        CollectionAddSorted result, accessObject.Name
                    End If
                End If
            Next accessObject
    End Select

    Set GetComposerObjectList = result
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetComposerObjectList", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Public Function GetComposerControlList( _
    ByVal objectType As String, _
    ByVal objectName As String, _
    Optional ByVal OnlyNamedPrefix As String = "" _
) As Collection
    On Error GoTo ErrorHandler

    Dim normalizedObjectType As String
    Dim normalizedPrefix As String
    Dim result As Collection
    Dim ctl As Control
    Dim wasOpenedByService As Boolean

    normalizedObjectType = NormalizeObjectType(objectType)
    normalizedPrefix = Trim$(OnlyNamedPrefix)
    Set result = New Collection

    If IsComposerInternalObject(objectName) Then
        Set GetComposerControlList = result
        Exit Function
    End If

    wasOpenedByService = OpenObjectHiddenDesign(normalizedObjectType, objectName)

    Select Case normalizedObjectType
        Case OBJECT_TYPE_FORM
            For Each ctl In Forms(objectName).Controls
                If LenB(normalizedPrefix) = 0 _
                    Or StrComp(Left$(ctl.Name, Len(normalizedPrefix)), normalizedPrefix, vbTextCompare) = 0 Then
                    CollectionAddSorted result, ctl.Name
                End If
            Next ctl

        Case OBJECT_TYPE_REPORT
            For Each ctl In Reports(objectName).Controls
                If LenB(normalizedPrefix) = 0 _
                    Or StrComp(Left$(ctl.Name, Len(normalizedPrefix)), normalizedPrefix, vbTextCompare) = 0 Then
                    CollectionAddSorted result, ctl.Name
                End If
            Next ctl
    End Select

CleanExit:
    If wasOpenedByService Then
        CloseObjectNoSave normalizedObjectType, objectName
    End If

    Set GetComposerControlList = result
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetComposerControlList", Err
    Resume CleanExit
End Function

Public Function GetFormControlMetadata( _
    ByVal objectName As String, _
    Optional ByVal includeHidden As Boolean = False) As Collection
    On Error GoTo ErrorHandler

    Dim result As Collection
    Dim ctl As Control
    Dim metadata As Object
    Dim openedByService As Boolean
    Dim isVisible As Boolean

    Set result = New Collection

    If IsComposerInternalObject(objectName) Then
        Set GetFormControlMetadata = result
        Exit Function
    End If

    openedByService = OpenObjectHiddenDesign(OBJECT_TYPE_FORM, objectName)

    For Each ctl In Forms(objectName).Controls
        isVisible = GetControlVisibleSafely(ctl)

        If includeHidden Or isVisible Then
            Set metadata = CreateObject("Scripting.Dictionary")
            metadata.CompareMode = vbTextCompare
            metadata("control_name") = ctl.Name
            metadata("control_type_id") = ctl.ControlType
            metadata("control_type") = ResolveControlTypeName(ctl.ControlType)
            metadata("caption_value") = GetObjectPropertySafely(ctl, "Caption")
            metadata("attached_label_caption") = GetAttachedLabelCaptionSafely(ctl)
            metadata("source_text") = ResolveControlSourceText(ctl)
            metadata("current_tag") = GetControlTagSafely(ctl)
            metadata("current_translation_key") = modFwTranslationRuntime.GetTranslationKeyFromTag(CStr(metadata("current_tag")))
            metadata("has_translation_marker") = (InStr(1, CStr(metadata("current_tag")), "TR:", vbTextCompare) > 0)
            metadata("is_visible") = isVisible
            result.Add metadata
        End If
    Next ctl

CleanExit:
    If openedByService Then
        CloseObjectNoSave OBJECT_TYPE_FORM, objectName
    End If

    Set GetFormControlMetadata = result
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetFormControlMetadata", Err
    Resume CleanExit
End Function

Public Function SaveControlTagsToObject( _
    ByVal objectType As String, _
    ByVal objectName As String, _
    ByVal controlTagMap As Object, _
    Optional ByRef updatedCount As Long = 0) As Boolean
    On Error GoTo ErrorHandler

    Dim normalizedObjectType As String
    Dim ControlName As Variant
    Dim WasLoaded As Boolean
    Dim WasVisible As Boolean
    Dim openedByService As Boolean
    Dim saveSucceeded As Boolean

    normalizedObjectType = NormalizeObjectType(objectType)
    objectName = Trim$(objectName)

    If LenB(objectName) = 0 Then
        Exit Function
    End If

    If controlTagMap Is Nothing Then
        SaveControlTagsToObject = True
        Exit Function
    End If

    WasLoaded = IsObjectLoaded(normalizedObjectType, objectName)
    WasVisible = GetObjectVisible(normalizedObjectType, objectName)
    openedByService = OpenObjectForUpdate(normalizedObjectType, objectName)

    For Each ControlName In controlTagMap.Keys
        If SetControlTagForLoadedObject(normalizedObjectType, objectName, CStr(ControlName), modDaoHelper.NzString(controlTagMap(ControlName))) Then
            updatedCount = updatedCount + 1
        Else
            modLoggingHandler.LogWarning MODULE_NAME & ".SaveControlTagsToObject", _
                "Tag write skipped for " & normalizedObjectType & "." & objectName & "." & CStr(ControlName) & "."
        End If
    Next ControlName

    If updatedCount > 0 Then
        SaveOpenedObject normalizedObjectType, objectName
    End If

    saveSucceeded = True

CleanExit:
    On Error Resume Next
    If openedByService Then
        CloseObjectNoSave normalizedObjectType, objectName
    End If
    RestoreObjectState normalizedObjectType, objectName, WasLoaded, WasVisible
    SaveControlTagsToObject = saveSucceeded
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SaveControlTagsToObject", Err
    Resume CleanExit
End Function

Public Function SuggestTranslationKey( _
    ByVal objectType As String, _
    ByVal objectName As String, _
    ByVal ControlName As String _
) As String
    On Error GoTo ErrorHandler

    Dim normalizedObjectType As String
    Dim normalizedObjectName As String
    Dim normalizedControlName As String

    normalizedObjectType = NormalizeObjectType(objectType)
    normalizedObjectName = UCase$(Trim$(objectName))
    normalizedControlName = NormalizeTranslationName(StripPrefix(ControlName, "lbl"))

    Select Case normalizedObjectType
        Case OBJECT_TYPE_REPORT
            SuggestTranslationKey = "REPORT." & normalizedControlName
        Case OBJECT_TYPE_FORM
            SuggestTranslationKey = "FORM." & normalizedObjectName & "." & normalizedControlName
    End Select
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SuggestTranslationKey", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Public Function GetControlTagValue( _
    ByVal objectType As String, _
    ByVal objectName As String, _
    ByVal ControlName As String _
) As String
    On Error GoTo ErrorHandler

    GetControlTagValue = GetControlStringProperty(objectType, objectName, ControlName, "Tag")
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetControlTagValue", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Public Function GetControlCaptionValue( _
    ByVal objectType As String, _
    ByVal objectName As String, _
    ByVal ControlName As String _
) As String
    On Error GoTo ErrorHandler

    GetControlCaptionValue = GetControlStringProperty(objectType, objectName, ControlName, "Caption")
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetControlCaptionValue", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Public Function GetControlTranslationKeyValue( _
    ByVal objectType As String, _
    ByVal objectName As String, _
    ByVal ControlName As String _
) As String
    On Error GoTo ErrorHandler

    GetControlTranslationKeyValue = modFwTranslationRuntime.GetTranslationKeyFromTag( _
        GetControlTagValue(objectType, objectName, ControlName))
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetControlTranslationKeyValue", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Public Function EnsureTranslationPlaceholders(ByVal translationKey As String) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim languageCodes As Variant
    Dim languageCode As Variant
    Dim insertedCount As Long

    translationKey = Trim$(translationKey)
    If LenB(translationKey) = 0 Then
        Exit Function
    End If

    Set db = currentDb
    If Not modDbSchema.TableExists(db, TRANSLATION_TABLE_NAME) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".EnsureTranslationPlaceholders", _
            "Translation table not found: " & TRANSLATION_TABLE_NAME & "."
        Exit Function
    End If

    languageCodes = Array("de-CH", "fr-CH", "en-US")

    For Each languageCode In languageCodes
        If EnsureTranslationPlaceholderRow(db, translationKey, CStr(languageCode)) Then
            insertedCount = insertedCount + 1
        End If
    Next languageCode

    If insertedCount > 0 Then
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureTranslationPlaceholders", _
            "Inserted " & CStr(insertedCount) & " placeholder translation row(s) for " & translationKey & "."
    End If

    EnsureTranslationPlaceholders = insertedCount
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureTranslationPlaceholders", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Public Function ValidateTranslationsReady( _
    Optional ByVal objectType As String = "", _
    Optional ByVal objectName As String = "" _
) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim missingCount As Long
    Dim whereClause As String
    Dim scopeDescription As String

    Set db = currentDb

    If Not modDbSchema.TableExists(db, TRANSLATION_TABLE_NAME) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ValidateTranslationsReady", _
            "Translation table not found: " & TRANSLATION_TABLE_NAME & "."
        Exit Function
    End If

    whereClause = BuildMissingTranslationWhereClause(objectType, objectName)
    missingCount = CountTranslationsByWhereClause(db, whereClause)
    scopeDescription = BuildValidationScopeDescription(objectType, objectName)

    modLoggingHandler.LogInfo MODULE_NAME & ".ValidateTranslationsReady", _
        "Missing translations found: " & CStr(missingCount) & " (" & scopeDescription & ")."

    ValidateTranslationsReady = (missingCount = 0)
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ValidateTranslationsReady", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Public Function ApplyTranslationTagsToObject( _
    ByVal objectType As String, _
    ByVal objectName As String, _
    Optional ByRef AppliedCount As Long = 0, _
    Optional ByRef SkippedCount As Long = 0 _
) As Boolean
    On Error GoTo ErrorHandler

    Dim normalizedObjectType As String
    Dim controlList As Collection
    Dim ControlName As Variant
    Dim translationKey As String
    Dim currentCaption As String
    Dim currentTag As String
    Dim updatedTag As String
    Dim WasLoaded As Boolean
    Dim WasVisible As Boolean
    Dim openedByService As Boolean
    Dim updatedObject As Boolean

    normalizedObjectType = NormalizeObjectType(objectType)
    objectName = Trim$(objectName)

    If LenB(objectName) = 0 Then
        Exit Function
    End If

    Set controlList = GetComposerControlList(normalizedObjectType, objectName, "lbl")
    If controlList Is Nothing Or controlList.count = 0 Then
        modLoggingHandler.LogInfo MODULE_NAME & ".ApplyTranslationTagsToObject", _
            "No translation-relevant controls found for " & normalizedObjectType & "." & objectName & "."
        ApplyTranslationTagsToObject = True
        Exit Function
    End If

    WasLoaded = IsObjectLoaded(normalizedObjectType, objectName)
    WasVisible = GetObjectVisible(normalizedObjectType, objectName)
    openedByService = OpenObjectForUpdate(normalizedObjectType, objectName)

    For Each ControlName In controlList
        translationKey = SuggestTranslationKey(normalizedObjectType, objectName, CStr(ControlName))

        If LenB(translationKey) = 0 Then
            SkippedCount = SkippedCount + 1
            GoTo NextControl
        End If

        currentCaption = GetControlCaptionForLoadedObject(normalizedObjectType, objectName, CStr(ControlName))
        currentTag = GetControlTagForLoadedObject(normalizedObjectType, objectName, CStr(ControlName))

        If LenB(currentCaption) = 0 Then
            SkippedCount = SkippedCount + 1
            GoTo NextControl
        End If

        updatedTag = modFwTranslationRuntime.SetTranslationKeyInTag(currentTag, translationKey)
        If StrComp(updatedTag, currentTag, vbBinaryCompare) = 0 And _
           LenB(modFwTranslationRuntime.GetTranslationKeyFromTag(currentTag)) > 0 Then
            SkippedCount = SkippedCount + 1
            GoTo NextControl
        End If

        If SetControlTagForLoadedObject(normalizedObjectType, objectName, CStr(ControlName), updatedTag) Then
            If StrComp(Left$(Trim$(currentCaption), 3), "TR:", vbTextCompare) = 0 Then
                SetControlCaptionForLoadedObject normalizedObjectType, objectName, CStr(ControlName), _
                    ResolveReadableFallbackCaption(translationKey, currentCaption, CStr(ControlName))
            End If

            AppliedCount = AppliedCount + 1
            updatedObject = True
        Else
            SkippedCount = SkippedCount + 1
        End If

NextControl:
    Next ControlName

    If updatedObject Then
        SaveOpenedObject normalizedObjectType, objectName
    End If

    If openedByService Then
        CloseObjectNoSave normalizedObjectType, objectName
    End If

    RestoreObjectState normalizedObjectType, objectName, WasLoaded, WasVisible

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyTranslationTagsToObject", _
        "Updated " & normalizedObjectType & "." & objectName & _
        " | applied=" & CStr(AppliedCount) & _
        " | skipped=" & CStr(SkippedCount) & "."

    ApplyTranslationTagsToObject = True
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyTranslationTagsToObject", Err
    On Error Resume Next
    If openedByService Then
        CloseObjectNoSave normalizedObjectType, objectName
    End If
    RestoreObjectState normalizedObjectType, objectName, WasLoaded, WasVisible
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Private Function GetControlStringProperty( _
    ByVal objectType As String, _
    ByVal objectName As String, _
    ByVal ControlName As String, _
    ByVal propertyName As String _
) As String
    On Error GoTo ErrorHandler

    Dim normalizedObjectType As String
    Dim wasOpenedByService As Boolean

    normalizedObjectType = NormalizeObjectType(objectType)

    If IsComposerInternalObject(objectName) Then
        GetControlStringProperty = vbNullString
        Exit Function
    End If

    wasOpenedByService = OpenObjectHiddenDesign(normalizedObjectType, objectName)

    If Not ControlExists(normalizedObjectType, objectName, ControlName) Then
        GetControlStringProperty = vbNullString
        GoTo CleanExit
    End If

    Select Case normalizedObjectType
        Case OBJECT_TYPE_FORM
            GetControlStringProperty = GetObjectPropertySafely(Forms(objectName).Controls(ControlName), propertyName)
        Case OBJECT_TYPE_REPORT
            GetControlStringProperty = GetObjectPropertySafely(Reports(objectName).Controls(ControlName), propertyName)
    End Select

CleanExit:
    If wasOpenedByService Then
        CloseObjectNoSave normalizedObjectType, objectName
    End If
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetControlStringProperty", Err
    Resume CleanExit
End Function

Private Function NormalizeObjectType(ByVal objectType As String) As String
    Dim normalizedType As String

    normalizedType = UCase$(Trim$(objectType))

    Select Case normalizedType
        Case OBJECT_TYPE_FORM, OBJECT_TYPE_REPORT
            NormalizeObjectType = normalizedType
        Case Else
            Err.Raise vbObjectError + 3201, MODULE_NAME & ".NormalizeObjectType", _
                "Unsupported object type: " & objectType
    End Select
End Function

Private Function IsComposerInternalObject(ByVal objectName As String) As Boolean
    Select Case UCase$(Trim$(objectName))
        Case "FRMFWCOMPOSER", _
             "FRMFWTRANSLATIONS", _
             "FRMFWTRANSLATIONLIST"
            IsComposerInternalObject = True
        Case Else
            IsComposerInternalObject = False
    End Select
End Function

Private Function NormalizeTranslationName(ByVal Value As String) As String
    Dim i As Long
    Dim currentChar As String
    Dim previousChar As String
    Dim nextChar As String
    Dim result As String

    Value = Trim$(Value)
    If LenB(Value) = 0 Then
        Exit Function
    End If

    For i = 1 To Len(Value)
        currentChar = Mid$(Value, i, 1)

        If IsAsciiLetterOrDigit(currentChar) Then
            previousChar = vbNullString
            nextChar = vbNullString

            If i > 1 Then
                previousChar = Mid$(Value, i - 1, 1)
            End If

            If i < Len(Value) Then
                nextChar = Mid$(Value, i + 1, 1)
            End If

            If LenB(result) > 0 And Right$(result, 1) <> "_" Then
                If IsAsciiUpper(currentChar) Then
                    If IsAsciiLower(previousChar) Or IsAsciiDigit(previousChar) Then
                        result = result & "_"
                    ElseIf IsAsciiUpper(previousChar) And IsAsciiLower(nextChar) Then
                        result = result & "_"
                    End If
                ElseIf IsAsciiDigit(currentChar) Then
                    If IsAsciiLetter(previousChar) Then
                        result = result & "_"
                    End If
                ElseIf IsAsciiDigit(previousChar) Then
                    result = result & "_"
                End If
            End If

            result = result & UCase$(currentChar)
        Else
            If LenB(result) > 0 And Right$(result, 1) <> "_" Then
                result = result & "_"
            End If
        End If
    Next i

    Do While Len(result) > 0 And Left$(result, 1) = "_"
        result = Mid$(result, 2)
    Loop

    Do While Len(result) > 0 And Right$(result, 1) = "_"
        result = Left$(result, Len(result) - 1)
    Loop

    Do While InStr(1, result, "__", vbBinaryCompare) > 0
        result = Replace(result, "__", "_")
    Loop

    NormalizeTranslationName = result
End Function

Private Function IsAsciiUpper(ByVal Value As String) As Boolean
    Dim charCode As Long

    If LenB(Value) = 0 Then
        Exit Function
    End If

    charCode = AscW(Left$(Value, 1))
    IsAsciiUpper = (charCode >= 65 And charCode <= 90)
End Function

Private Function IsAsciiLower(ByVal Value As String) As Boolean
    Dim charCode As Long

    If LenB(Value) = 0 Then
        Exit Function
    End If

    charCode = AscW(Left$(Value, 1))
    IsAsciiLower = (charCode >= 97 And charCode <= 122)
End Function

Private Function IsAsciiDigit(ByVal Value As String) As Boolean
    Dim charCode As Long

    If LenB(Value) = 0 Then
        Exit Function
    End If

    charCode = AscW(Left$(Value, 1))
    IsAsciiDigit = (charCode >= 48 And charCode <= 57)
End Function

Private Function IsAsciiLetter(ByVal Value As String) As Boolean
    IsAsciiLetter = IsAsciiUpper(Value) Or IsAsciiLower(Value)
End Function

Private Function IsAsciiLetterOrDigit(ByVal Value As String) As Boolean
    IsAsciiLetterOrDigit = IsAsciiLetter(Value) Or IsAsciiDigit(Value)
End Function

' Immediate Window examples:
' ? NormalizeTranslationName("PaymentTerms")
' PAYMENT_TERMS
' ? NormalizeTranslationName("VATAmount")
' VAT_AMOUNT
' ? NormalizeTranslationName("URLValue")
' URL_VALUE

Private Function StripPrefix(ByVal Value As String, ByVal Prefix As String) As String
    If LenB(Prefix) > 0 And StrComp(Left$(Value, Len(Prefix)), Prefix, vbTextCompare) = 0 Then
        StripPrefix = Mid$(Value, Len(Prefix) + 1)
    Else
        StripPrefix = Value
    End If
End Function

Private Function ObjectExists(ByVal objectType As String, ByVal objectName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim accessObject As Access.accessObject
    Dim normalizedObjectType As String

    normalizedObjectType = NormalizeObjectType(objectType)

    Select Case normalizedObjectType
        Case OBJECT_TYPE_FORM
            For Each accessObject In CurrentProject.AllForms
                If StrComp(accessObject.Name, objectName, vbTextCompare) = 0 Then
                    ObjectExists = True
                    Exit Function
                End If
            Next accessObject

        Case OBJECT_TYPE_REPORT
            For Each accessObject In CurrentProject.AllReports
                If StrComp(accessObject.Name, objectName, vbTextCompare) = 0 Then
                    ObjectExists = True
                    Exit Function
                End If
            Next accessObject
    End Select

    Exit Function

ErrorHandler:
    ObjectExists = False
End Function

Private Function OpenObjectHiddenDesign(ByVal objectType As String, ByVal objectName As String) As Boolean
    On Error GoTo ErrorHandler

    OpenObjectHiddenDesign = False

    If Not ObjectExists(objectType, objectName) Then
        Err.Raise vbObjectError + 3202, MODULE_NAME & ".OpenObjectHiddenDesign", _
            objectType & " '" & objectName & "' does not exist."
    End If

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            If Not CurrentProject.AllForms(objectName).IsLoaded Then
                DoCmd.openForm objectName, acDesign, , , , acHidden
                OpenObjectHiddenDesign = True
            End If

        Case OBJECT_TYPE_REPORT
            If Not CurrentProject.AllReports(objectName).IsLoaded Then
                DoCmd.OpenReport objectName, acViewDesign, , , acHidden
                OpenObjectHiddenDesign = True
            End If
    End Select
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "OpenObjectHiddenDesign", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Private Sub CloseObjectNoSave(ByVal objectType As String, ByVal objectName As String)
    On Error Resume Next

    Select Case UCase$(Trim$(objectType))
        Case OBJECT_TYPE_FORM
            If ObjectExists(OBJECT_TYPE_FORM, objectName) Then
                If CurrentProject.AllForms(objectName).IsLoaded Then
                    DoCmd.Close acForm, objectName, acSaveNo
                End If
            End If

        Case OBJECT_TYPE_REPORT
            If ObjectExists(OBJECT_TYPE_REPORT, objectName) Then
                If CurrentProject.AllReports(objectName).IsLoaded Then
                    DoCmd.Close acReport, objectName, acSaveNo
                End If
            End If
    End Select
End Sub

Private Function IsObjectLoaded(ByVal objectType As String, ByVal objectName As String) As Boolean
    On Error GoTo SafeExit

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            IsObjectLoaded = CurrentProject.AllForms(objectName).IsLoaded
        Case OBJECT_TYPE_REPORT
            IsObjectLoaded = CurrentProject.AllReports(objectName).IsLoaded
    End Select

SafeExit:
End Function

Private Function GetObjectVisible(ByVal objectType As String, ByVal objectName As String) As Boolean
    On Error GoTo SafeExit

    If NormalizeObjectType(objectType) = OBJECT_TYPE_FORM Then
        If CurrentProject.AllForms(objectName).IsLoaded Then
            GetObjectVisible = Forms(objectName).Visible
        End If
    End If

SafeExit:
End Function

Private Function OpenObjectForUpdate(ByVal objectType As String, ByVal objectName As String) As Boolean
    On Error GoTo ErrorHandler

    OpenObjectForUpdate = False

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            If CurrentProject.AllForms(objectName).IsLoaded Then
                DoCmd.Close acForm, objectName, acSaveYes
            End If
            DoCmd.openForm objectName, acDesign, , , , acHidden
            OpenObjectForUpdate = True

        Case OBJECT_TYPE_REPORT
            If CurrentProject.AllReports(objectName).IsLoaded Then
                DoCmd.Close acReport, objectName, acSaveYes
            End If
            DoCmd.OpenReport objectName, acViewDesign, , , acHidden
            OpenObjectForUpdate = True
    End Select
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "OpenObjectForUpdate", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Private Sub SaveOpenedObject(ByVal objectType As String, ByVal objectName As String)
    On Error GoTo ErrorHandler

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            DoCmd.Save acForm, objectName
        Case OBJECT_TYPE_REPORT
            DoCmd.Save acReport, objectName
    End Select
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SaveOpenedObject", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub RestoreObjectState( _
    ByVal objectType As String, _
    ByVal objectName As String, _
    ByVal WasLoaded As Boolean, _
    ByVal WasVisible As Boolean)
    On Error GoTo SafeExit

    If Not WasLoaded Then
        Exit Sub
    End If

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            If Not CurrentProject.AllForms(objectName).IsLoaded Then
                DoCmd.openForm objectName, acNormal
            End If
            If Not WasVisible Then
                Forms(objectName).Visible = False
            End If

        Case OBJECT_TYPE_REPORT
            If Not CurrentProject.AllReports(objectName).IsLoaded Then
                DoCmd.OpenReport objectName, acViewPreview
            End If
    End Select

SafeExit:
End Sub

Private Sub CollectionAddSorted(ByVal target As Collection, ByVal itemValue As String)
    On Error GoTo ErrorHandler

    Dim tempValues() As String
    Dim i As Long
    Dim j As Long
    Dim inserted As Boolean
    Dim currentValue As String

    If target Is Nothing Then
        Exit Sub
    End If

    If target.count = 0 Then
        target.Add itemValue
        Exit Sub
    End If

    ReDim tempValues(1 To target.count + 1)

    For i = 1 To target.count
        tempValues(i) = CStr(target(i))
    Next i

    tempValues(target.count + 1) = itemValue

    For i = LBound(tempValues) To UBound(tempValues) - 1
        For j = i + 1 To UBound(tempValues)
            If StrComp(tempValues(i), tempValues(j), vbTextCompare) > 0 Then
                currentValue = tempValues(i)
                tempValues(i) = tempValues(j)
                tempValues(j) = currentValue
            End If
        Next j
    Next i

    Do While target.count > 0
        target.Remove 1
    Loop

    inserted = False
    currentValue = vbNullString

    For i = LBound(tempValues) To UBound(tempValues)
        If Not inserted Or StrComp(currentValue, tempValues(i), vbTextCompare) <> 0 Then
            target.Add tempValues(i)
            inserted = True
        End If
        currentValue = tempValues(i)
    Next i

    Exit Sub

ErrorHandler:
    target.Add itemValue
End Sub

Private Function ControlExists(ByVal objectType As String, ByVal objectName As String, ByVal ControlName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ctl As Control

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            For Each ctl In Forms(objectName).Controls
                If StrComp(ctl.Name, ControlName, vbTextCompare) = 0 Then
                    ControlExists = True
                    Exit Function
                End If
            Next ctl

        Case OBJECT_TYPE_REPORT
            For Each ctl In Reports(objectName).Controls
                If StrComp(ctl.Name, ControlName, vbTextCompare) = 0 Then
                    ControlExists = True
                    Exit Function
                End If
            Next ctl
    End Select

    Exit Function

ErrorHandler:
    ControlExists = False
End Function

Private Function GetObjectPropertySafely(ByVal Obj As Object, ByVal propertyName As String) As String
    On Error GoTo SafeExit

    GetObjectPropertySafely = Nz(Obj.Properties(propertyName).Value, vbNullString)
    Exit Function

SafeExit:
    GetObjectPropertySafely = vbNullString
End Function

Private Function GetControlCaptionForLoadedObject( _
    ByVal objectType As String, _
    ByVal objectName As String, _
    ByVal ControlName As String _
) As String
    On Error GoTo SafeExit

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            GetControlCaptionForLoadedObject = GetObjectPropertySafely(Forms(objectName).Controls(ControlName), "Caption")
        Case OBJECT_TYPE_REPORT
            GetControlCaptionForLoadedObject = GetObjectPropertySafely(Reports(objectName).Controls(ControlName), "Caption")
    End Select
    Exit Function

SafeExit:
    GetControlCaptionForLoadedObject = vbNullString
End Function

Private Function GetAttachedLabelCaptionSafely(ByVal ctl As Object) As String
    On Error GoTo SafeExit

    Dim attachedLabel As Object

    If ctl Is Nothing Then
        Exit Function
    End If

    Set attachedLabel = CallByName(ctl, "AttachedLabel", VbGet)
    If attachedLabel Is Nothing Then
        Exit Function
    End If

    GetAttachedLabelCaptionSafely = GetObjectPropertySafely(attachedLabel, "Caption")
    Exit Function

SafeExit:
    GetAttachedLabelCaptionSafely = vbNullString
End Function

Private Function GetControlVisibleSafely(ByVal ctl As Control) As Boolean
    On Error GoTo SafeExit

    If ctl Is Nothing Then
        Exit Function
    End If

    GetControlVisibleSafely = CBool(ctl.Properties("Visible").Value)
    Exit Function

SafeExit:
    GetControlVisibleSafely = True
End Function

Private Function GetControlTagSafely(ByVal ctl As Control) As String
    On Error GoTo SafeExit

    If ctl Is Nothing Then
        Exit Function
    End If

    GetControlTagSafely = GetObjectPropertySafely(ctl, "Tag")
    Exit Function

SafeExit:
    GetControlTagSafely = vbNullString
End Function

Private Function ResolveControlSourceText(ByVal ctl As Control) As String
    Dim CaptionValue As String
    Dim attachedLabelCaption As String

    CaptionValue = GetObjectPropertySafely(ctl, "Caption")
    If LenB(Trim$(CaptionValue)) > 0 Then
        ResolveControlSourceText = CaptionValue
        Exit Function
    End If

    attachedLabelCaption = GetAttachedLabelCaptionSafely(ctl)
    If LenB(Trim$(attachedLabelCaption)) > 0 Then
        ResolveControlSourceText = attachedLabelCaption
        Exit Function
    End If

    ResolveControlSourceText = ctl.Name
End Function

Private Function ResolveControlTypeName(ByVal controlTypeId As Long) As String
    Select Case controlTypeId
        Case acLabel
            ResolveControlTypeName = "Label"
        Case acCommandButton
            ResolveControlTypeName = "CommandButton"
        Case acPage
            ResolveControlTypeName = "TabPage"
        Case acOptionButton
            ResolveControlTypeName = "OptionButton"
        Case acCheckBox
            ResolveControlTypeName = "CheckBox"
        Case acToggleButton
            ResolveControlTypeName = "ToggleButton"
        Case acComboBox
            ResolveControlTypeName = "ComboBox"
        Case acListBox
            ResolveControlTypeName = "ListBox"
        Case acTextBox
            ResolveControlTypeName = "TextBox"
        Case acSubform
            ResolveControlTypeName = "Subform"
        Case Else
            ResolveControlTypeName = "ControlType" & CStr(controlTypeId)
    End Select
End Function

Private Function GetControlTagForLoadedObject( _
    ByVal objectType As String, _
    ByVal objectName As String, _
    ByVal ControlName As String _
) As String
    On Error GoTo SafeExit

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            GetControlTagForLoadedObject = GetObjectPropertySafely(Forms(objectName).Controls(ControlName), "Tag")
        Case OBJECT_TYPE_REPORT
            GetControlTagForLoadedObject = GetObjectPropertySafely(Reports(objectName).Controls(ControlName), "Tag")
    End Select
    Exit Function

SafeExit:
    GetControlTagForLoadedObject = vbNullString
End Function

Private Function SetControlCaptionForLoadedObject( _
    ByVal objectType As String, _
    ByVal objectName As String, _
    ByVal ControlName As String, _
    ByVal CaptionValue As String _
) As Boolean
    On Error GoTo SafeExit

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            Forms(objectName).Controls(ControlName).Properties("Caption").Value = CaptionValue
        Case OBJECT_TYPE_REPORT
            Reports(objectName).Controls(ControlName).Properties("Caption").Value = CaptionValue
    End Select

    SetControlCaptionForLoadedObject = True
    Exit Function

SafeExit:
    SetControlCaptionForLoadedObject = False
End Function

Private Function SetControlTagForLoadedObject( _
    ByVal objectType As String, _
    ByVal objectName As String, _
    ByVal ControlName As String, _
    ByVal TagValue As String _
) As Boolean
    On Error GoTo SafeExit

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            Forms(objectName).Controls(ControlName).Properties("Tag").Value = TagValue
        Case OBJECT_TYPE_REPORT
            Reports(objectName).Controls(ControlName).Properties("Tag").Value = TagValue
    End Select

    SetControlTagForLoadedObject = True
    Exit Function

SafeExit:
    SetControlTagForLoadedObject = False
End Function

Private Function ResolveReadableFallbackCaption( _
    ByVal translationKey As String, _
    ByVal currentCaption As String, _
    ByVal ControlName As String) As String

    ResolveReadableFallbackCaption = modFwTranslationRuntime.ResolveCaptionText(vbNullString, "TR:" & translationKey)

    If LenB(Trim$(ResolveReadableFallbackCaption)) = 0 Then
        ResolveReadableFallbackCaption = BuildReadableCaptionFromControlName(ControlName)
    ElseIf StrComp(Trim$(ResolveReadableFallbackCaption), Trim$(translationKey), vbTextCompare) = 0 Then
        ResolveReadableFallbackCaption = BuildReadableCaptionFromControlName(ControlName)
    ElseIf StrComp(Trim$(ResolveReadableFallbackCaption), Trim$(currentCaption), vbTextCompare) = 0 Then
        ResolveReadableFallbackCaption = BuildReadableCaptionFromControlName(ControlName)
    End If
End Function

Private Function BuildReadableCaptionFromControlName(ByVal ControlName As String) As String
    Dim fallbackCaption As String

    fallbackCaption = NormalizeTranslationName(StripPrefix(ControlName, "lbl"))
    fallbackCaption = Replace(fallbackCaption, "_", " ")

    If LenB(Trim$(fallbackCaption)) = 0 Then
        BuildReadableCaptionFromControlName = ControlName
    Else
        BuildReadableCaptionFromControlName = fallbackCaption
    End If
End Function

Private Function EnsureTranslationPlaceholderRow( _
    ByVal db As DAO.Database, _
    ByVal translationKey As String, _
    ByVal languageCode As String _
) As Boolean
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset

    If db Is Nothing Then
        Exit Function
    End If

    If TranslationRowExists(db, translationKey, languageCode) Then
        Exit Function
    End If

    Set rs = db.OpenRecordset(TRANSLATION_TABLE_NAME, dbOpenDynaset)

    rs.AddNew
    rs.Fields(FIELD_TRANSLATION_KEY).Value = translationKey
    rs.Fields(FIELD_LANGUAGE_CODE).Value = languageCode
    rs.Fields(FIELD_TRANSLATION_VALUE).Value = PLACEHOLDER_TRANSLATION_VALUE
    SetRecordsetFieldIfExists rs, FIELD_IS_ACTIVE, True
    SetRecordsetFieldIfExists rs, FIELD_MODULE_CODE, GetTranslationModuleCode(translationKey)
    SetRecordsetFieldIfExists rs, FIELD_UPDATED_AT, Now
    rs.Update

    EnsureTranslationPlaceholderRow = True

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureTranslationPlaceholderRow", Err
    Resume CleanExit
End Function

Private Function TranslationRowExists( _
    ByVal db As DAO.Database, _
    ByVal translationKey As String, _
    ByVal languageCode As String _
) As Boolean
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlStatement As String

    sqlStatement = "SELECT TOP 1 " & FIELD_TRANSLATION_KEY & _
                   " FROM " & TRANSLATION_TABLE_NAME & _
                   " WHERE " & FIELD_TRANSLATION_KEY & " = " & SqlText(translationKey) & _
                   " AND " & FIELD_LANGUAGE_CODE & " = " & SqlText(languageCode)

    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)
    TranslationRowExists = Not (rs.BOF And rs.EOF)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "TranslationRowExists", Err
    Resume CleanExit
End Function

Private Sub SetRecordsetFieldIfExists( _
    ByVal rs As DAO.Recordset, _
    ByVal fieldName As String, _
    ByVal fieldValue As Variant _
)
    On Error GoTo SafeExit

    Dim fld As DAO.Field

    If rs Is Nothing Then
        Exit Sub
    End If

    For Each fld In rs.Fields
        If StrComp(fld.Name, fieldName, vbTextCompare) = 0 Then
            fld.Value = fieldValue
            Exit For
        End If
    Next fld

SafeExit:
End Sub

Private Function GetTranslationModuleCode(ByVal translationKey As String) As String
    translationKey = UCase$(Trim$(translationKey))

    If Left$(translationKey, 7) = "REPORT." Then
        GetTranslationModuleCode = "REPORT"
    ElseIf Left$(translationKey, 5) = "FORM." Then
        GetTranslationModuleCode = "FORM"
    Else
        GetTranslationModuleCode = "FRAMEWORK"
    End If
End Function

Private Function BuildMissingTranslationWhereClause( _
    ByVal objectType As String, _
    ByVal objectName As String _
) As String
    Dim scopeClause As String

    scopeClause = BuildTranslationScopeWhereClause(objectType, objectName)
    BuildMissingTranslationWhereClause = BuildMissingTranslationValueWhereClause()

    If LenB(scopeClause) > 0 Then
        BuildMissingTranslationWhereClause = "(" & scopeClause & ") AND (" & _
                                             BuildMissingTranslationWhereClause & ")"
    End If
End Function

Private Function BuildMissingTranslationValueWhereClause() As String
    BuildMissingTranslationValueWhereClause = _
        "(" & FIELD_TRANSLATION_VALUE & " = " & SqlText(PLACEHOLDER_TRANSLATION_VALUE) & _
        " OR " & FIELD_TRANSLATION_VALUE & " Is Null" & _
        " OR Trim(Nz([" & FIELD_TRANSLATION_VALUE & "], '')) = '')"
End Function

Private Function BuildTranslationScopeWhereClause( _
    ByVal objectType As String, _
    ByVal objectName As String _
) As String
    Dim normalizedObjectType As String
    Dim normalizedObjectName As String

    normalizedObjectType = vbNullString
    normalizedObjectName = UCase$(Trim$(objectName))

    If LenB(Trim$(objectType)) > 0 Then
        normalizedObjectType = NormalizeObjectType(objectType)
    End If

    Select Case normalizedObjectType
        Case OBJECT_TYPE_FORM
            If LenB(normalizedObjectName) > 0 Then
                BuildTranslationScopeWhereClause = "[" & FIELD_TRANSLATION_KEY & "] Like " & _
                    SqlText("FORM." & normalizedObjectName & ".*")
            Else
                BuildTranslationScopeWhereClause = "[" & FIELD_TRANSLATION_KEY & "] Like " & _
                    SqlText("FORM.*")
            End If

        Case OBJECT_TYPE_REPORT
            BuildTranslationScopeWhereClause = "[" & FIELD_TRANSLATION_KEY & "] Like " & _
                SqlText("REPORT.*")

        Case Else
            BuildTranslationScopeWhereClause = vbNullString
    End Select
End Function

Private Function CountTranslationsByWhereClause( _
    ByVal db As DAO.Database, _
    ByVal whereClause As String _
) As Long
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sqlStatement As String

    sqlStatement = "SELECT Count(*) AS MissingCount FROM " & TRANSLATION_TABLE_NAME
    If LenB(Trim$(whereClause)) > 0 Then
        sqlStatement = sqlStatement & " WHERE " & whereClause
    End If

    Set rs = db.OpenRecordset(sqlStatement, dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then
        CountTranslationsByWhereClause = Nz(rs.Fields(0).Value, 0)
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "CountTranslationsByWhereClause", Err
    Resume CleanExit
End Function

Private Function BuildValidationScopeDescription( _
    ByVal objectType As String, _
    ByVal objectName As String _
) As String
    Dim normalizedObjectType As String
    Dim normalizedObjectName As String

    normalizedObjectType = UCase$(Trim$(objectType))
    normalizedObjectName = Trim$(objectName)

    If LenB(normalizedObjectType) = 0 Then
        BuildValidationScopeDescription = "all translations"
    ElseIf LenB(normalizedObjectName) = 0 Then
        BuildValidationScopeDescription = normalizedObjectType
    ElseIf normalizedObjectType = OBJECT_TYPE_REPORT Then
        BuildValidationScopeDescription = normalizedObjectType & " " & normalizedObjectName & " (report scope uses REPORT.*)"
    Else
        BuildValidationScopeDescription = normalizedObjectType & " " & normalizedObjectName
    End If
End Function




