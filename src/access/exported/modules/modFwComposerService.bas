Attribute VB_Name = "modFwComposerService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFwComposerService
' Purpose   : Provides object and control inspection helpers for a future
'             framework composer form.
' Author    : Codex
' Version   : 0.1.3
'===============================================================================

Private Const MODULE_NAME As String = "modFwComposerService"
Private Const TRANSLATION_TABLE_NAME As String = "fw_translation"
Private Const PLACEHOLDER_TRANSLATION_VALUE As String = "<neu>"
Private Const FIELD_TRANSLATION_KEY As String = "TranslationKey"
Private Const FIELD_LANGUAGE_CODE As String = "LanguageCode"
Private Const FIELD_TRANSLATION_VALUE As String = "TranslationValue"
Private Const FIELD_IS_ACTIVE As String = "IsActive"
Private Const FIELD_MODULE_CODE As String = "ModuleCode"
Private Const FIELD_UPDATED_AT As String = "UpdatedAt"

Public Const COMPOSER_MODE_TAGS As String = "TAGS"
Public Const COMPOSER_MODE_TRANSLATIONS As String = "TRANSLATIONS"
Public Const OBJECT_TYPE_FORM As String = "FORM"
Public Const OBJECT_TYPE_REPORT As String = "REPORT"

Public Function GetComposerObjectList(ByVal ObjectType As String) As Collection
    On Error GoTo ErrorHandler

    Dim normalizedObjectType As String
    Dim result As Collection
    Dim accessObject As Access.accessObject

    normalizedObjectType = NormalizeObjectType(ObjectType)
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
    ByVal ObjectType As String, _
    ByVal ObjectName As String, _
    Optional ByVal OnlyNamedPrefix As String = "" _
) As Collection
    On Error GoTo ErrorHandler

    Dim normalizedObjectType As String
    Dim normalizedPrefix As String
    Dim result As Collection
    Dim ctl As Control
    Dim wasOpenedByService As Boolean

    normalizedObjectType = NormalizeObjectType(ObjectType)
    normalizedPrefix = Trim$(OnlyNamedPrefix)
    Set result = New Collection

    If IsComposerInternalObject(ObjectName) Then
        Set GetComposerControlList = result
        Exit Function
    End If

    wasOpenedByService = OpenObjectHiddenDesign(normalizedObjectType, ObjectName)

    Select Case normalizedObjectType
        Case OBJECT_TYPE_FORM
            For Each ctl In Forms(ObjectName).Controls
                If LenB(normalizedPrefix) = 0 _
                    Or StrComp(Left$(ctl.Name, Len(normalizedPrefix)), normalizedPrefix, vbTextCompare) = 0 Then
                    CollectionAddSorted result, ctl.Name
                End If
            Next ctl

        Case OBJECT_TYPE_REPORT
            For Each ctl In Reports(ObjectName).Controls
                If LenB(normalizedPrefix) = 0 _
                    Or StrComp(Left$(ctl.Name, Len(normalizedPrefix)), normalizedPrefix, vbTextCompare) = 0 Then
                    CollectionAddSorted result, ctl.Name
                End If
            Next ctl
    End Select

CleanExit:
    If wasOpenedByService Then
        CloseObjectNoSave normalizedObjectType, ObjectName
    End If

    Set GetComposerControlList = result
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetComposerControlList", Err
    Resume CleanExit
End Function

Public Function SuggestTranslationKey( _
    ByVal ObjectType As String, _
    ByVal ObjectName As String, _
    ByVal ControlName As String _
) As String
    On Error GoTo ErrorHandler

    Dim normalizedObjectType As String
    Dim normalizedObjectName As String
    Dim normalizedControlName As String

    normalizedObjectType = NormalizeObjectType(ObjectType)
    normalizedObjectName = UCase$(Trim$(ObjectName))
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
    ByVal ObjectType As String, _
    ByVal ObjectName As String, _
    ByVal ControlName As String _
) As String
    On Error GoTo ErrorHandler

    GetControlTagValue = GetControlStringProperty(ObjectType, ObjectName, ControlName, "Tag")
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetControlTagValue", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Public Function GetControlCaptionValue( _
    ByVal ObjectType As String, _
    ByVal ObjectName As String, _
    ByVal ControlName As String _
) As String
    On Error GoTo ErrorHandler

    GetControlCaptionValue = GetControlStringProperty(ObjectType, ObjectName, ControlName, "Caption")
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetControlCaptionValue", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Public Function EnsureTranslationPlaceholders(ByVal TranslationKey As String) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim languageCodes As Variant
    Dim LanguageCode As Variant
    Dim insertedCount As Long

    TranslationKey = Trim$(TranslationKey)
    If LenB(TranslationKey) = 0 Then
        Exit Function
    End If

    Set db = CurrentDb
    If Not TableExists(db, TRANSLATION_TABLE_NAME) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".EnsureTranslationPlaceholders", _
            "Translation table not found: " & TRANSLATION_TABLE_NAME & "."
        Exit Function
    End If

    languageCodes = Array("DE-CH", "FR-FR", "EN-US")

    For Each LanguageCode In languageCodes
        If EnsureTranslationPlaceholderRow(db, TranslationKey, CStr(LanguageCode)) Then
            insertedCount = insertedCount + 1
        End If
    Next LanguageCode

    If insertedCount > 0 Then
        modLoggingHandler.LogInfo MODULE_NAME & ".EnsureTranslationPlaceholders", _
            "Inserted " & CStr(insertedCount) & " placeholder translation row(s) for " & TranslationKey & "."
    End If

    EnsureTranslationPlaceholders = insertedCount
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureTranslationPlaceholders", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Public Function ValidateTranslationsReady( _
    Optional ByVal ObjectType As String = "", _
    Optional ByVal ObjectName As String = "" _
) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim missingCount As Long
    Dim whereClause As String
    Dim scopeDescription As String

    Set db = CurrentDb

    If Not TableExists(db, TRANSLATION_TABLE_NAME) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".ValidateTranslationsReady", _
            "Translation table not found: " & TRANSLATION_TABLE_NAME & "."
        Exit Function
    End If

    whereClause = BuildMissingTranslationWhereClause(ObjectType, ObjectName)
    missingCount = CountTranslationsByWhereClause(db, whereClause)
    scopeDescription = BuildValidationScopeDescription(ObjectType, ObjectName)

    modLoggingHandler.LogInfo MODULE_NAME & ".ValidateTranslationsReady", _
        "Missing translations found: " & CStr(missingCount) & " (" & scopeDescription & ")."

    ValidateTranslationsReady = (missingCount = 0)
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ValidateTranslationsReady", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Public Function ApplyTranslationTagsToObject( _
    ByVal ObjectType As String, _
    ByVal ObjectName As String, _
    Optional ByRef AppliedCount As Long = 0, _
    Optional ByRef SkippedCount As Long = 0 _
) As Boolean
    On Error GoTo ErrorHandler

    Dim normalizedObjectType As String
    Dim controlList As Collection
    Dim ControlName As Variant
    Dim TranslationKey As String
    Dim currentCaption As String
    Dim WasLoaded As Boolean
    Dim WasVisible As Boolean
    Dim openedByService As Boolean
    Dim updatedObject As Boolean

    normalizedObjectType = NormalizeObjectType(ObjectType)
    ObjectName = Trim$(ObjectName)

    If LenB(ObjectName) = 0 Then
        Exit Function
    End If

    Set controlList = GetComposerControlList(normalizedObjectType, ObjectName, "lbl")
    If controlList Is Nothing Or controlList.count = 0 Then
        modLoggingHandler.LogInfo MODULE_NAME & ".ApplyTranslationTagsToObject", _
            "No translation-relevant controls found for " & normalizedObjectType & "." & ObjectName & "."
        ApplyTranslationTagsToObject = True
        Exit Function
    End If

    WasLoaded = IsObjectLoaded(normalizedObjectType, ObjectName)
    WasVisible = GetObjectVisible(normalizedObjectType, ObjectName)
    openedByService = OpenObjectForUpdate(normalizedObjectType, ObjectName)

    For Each ControlName In controlList
        TranslationKey = SuggestTranslationKey(normalizedObjectType, ObjectName, CStr(ControlName))

        If LenB(TranslationKey) = 0 Then
            SkippedCount = SkippedCount + 1
            GoTo NextControl
        End If

        currentCaption = GetControlCaptionForLoadedObject(normalizedObjectType, ObjectName, CStr(ControlName))

        If LenB(currentCaption) = 0 Then
            SkippedCount = SkippedCount + 1
            GoTo NextControl
        End If

        If StrComp(Left$(Trim$(currentCaption), 3), "TR:", vbTextCompare) = 0 Then
            SkippedCount = SkippedCount + 1
            GoTo NextControl
        End If

        If SetControlCaptionForLoadedObject(normalizedObjectType, ObjectName, CStr(ControlName), "TR:" & TranslationKey) Then
            AppliedCount = AppliedCount + 1
            updatedObject = True
        Else
            SkippedCount = SkippedCount + 1
        End If

NextControl:
    Next ControlName

    If updatedObject Then
        SaveOpenedObject normalizedObjectType, ObjectName
    End If

    If openedByService Then
        CloseObjectNoSave normalizedObjectType, ObjectName
    End If

    RestoreObjectState normalizedObjectType, ObjectName, WasLoaded, WasVisible

    modLoggingHandler.LogInfo MODULE_NAME & ".ApplyTranslationTagsToObject", _
        "Updated " & normalizedObjectType & "." & ObjectName & _
        " | applied=" & CStr(AppliedCount) & _
        " | skipped=" & CStr(SkippedCount) & "."

    ApplyTranslationTagsToObject = True
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyTranslationTagsToObject", Err
    On Error Resume Next
    If openedByService Then
        CloseObjectNoSave normalizedObjectType, ObjectName
    End If
    RestoreObjectState normalizedObjectType, ObjectName, WasLoaded, WasVisible
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Private Function GetControlStringProperty( _
    ByVal ObjectType As String, _
    ByVal ObjectName As String, _
    ByVal ControlName As String, _
    ByVal PropertyName As String _
) As String
    On Error GoTo ErrorHandler

    Dim normalizedObjectType As String
    Dim wasOpenedByService As Boolean

    normalizedObjectType = NormalizeObjectType(ObjectType)

    If IsComposerInternalObject(ObjectName) Then
        GetControlStringProperty = vbNullString
        Exit Function
    End If

    wasOpenedByService = OpenObjectHiddenDesign(normalizedObjectType, ObjectName)

    If Not ControlExists(normalizedObjectType, ObjectName, ControlName) Then
        GetControlStringProperty = vbNullString
        GoTo CleanExit
    End If

    Select Case normalizedObjectType
        Case OBJECT_TYPE_FORM
            GetControlStringProperty = GetObjectPropertySafely(Forms(ObjectName).Controls(ControlName), PropertyName)
        Case OBJECT_TYPE_REPORT
            GetControlStringProperty = GetObjectPropertySafely(Reports(ObjectName).Controls(ControlName), PropertyName)
    End Select

CleanExit:
    If wasOpenedByService Then
        CloseObjectNoSave normalizedObjectType, ObjectName
    End If
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetControlStringProperty", Err
    Resume CleanExit
End Function

Private Function NormalizeObjectType(ByVal ObjectType As String) As String
    Dim normalizedType As String

    normalizedType = UCase$(Trim$(ObjectType))

    Select Case normalizedType
        Case OBJECT_TYPE_FORM, OBJECT_TYPE_REPORT
            NormalizeObjectType = normalizedType
        Case Else
            Err.Raise vbObjectError + 3201, MODULE_NAME & ".NormalizeObjectType", _
                "Unsupported object type: " & ObjectType
    End Select
End Function

Private Function IsComposerInternalObject(ByVal ObjectName As String) As Boolean
    Select Case UCase$(Trim$(ObjectName))
        Case "FRMFWCOMPOSER", _
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

Private Function ObjectExists(ByVal ObjectType As String, ByVal ObjectName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim accessObject As Access.accessObject
    Dim normalizedObjectType As String

    normalizedObjectType = NormalizeObjectType(ObjectType)

    Select Case normalizedObjectType
        Case OBJECT_TYPE_FORM
            For Each accessObject In CurrentProject.AllForms
                If StrComp(accessObject.Name, ObjectName, vbTextCompare) = 0 Then
                    ObjectExists = True
                    Exit Function
                End If
            Next accessObject

        Case OBJECT_TYPE_REPORT
            For Each accessObject In CurrentProject.AllReports
                If StrComp(accessObject.Name, ObjectName, vbTextCompare) = 0 Then
                    ObjectExists = True
                    Exit Function
                End If
            Next accessObject
    End Select

    Exit Function

ErrorHandler:
    ObjectExists = False
End Function

Private Function OpenObjectHiddenDesign(ByVal ObjectType As String, ByVal ObjectName As String) As Boolean
    On Error GoTo ErrorHandler

    OpenObjectHiddenDesign = False

    If Not ObjectExists(ObjectType, ObjectName) Then
        Err.Raise vbObjectError + 3202, MODULE_NAME & ".OpenObjectHiddenDesign", _
            ObjectType & " '" & ObjectName & "' does not exist."
    End If

    Select Case NormalizeObjectType(ObjectType)
        Case OBJECT_TYPE_FORM
            If Not CurrentProject.AllForms(ObjectName).IsLoaded Then
                DoCmd.OpenForm ObjectName, acDesign, , , , acHidden
                OpenObjectHiddenDesign = True
            End If

        Case OBJECT_TYPE_REPORT
            If Not CurrentProject.AllReports(ObjectName).IsLoaded Then
                DoCmd.OpenReport ObjectName, acViewDesign, , , acHidden
                OpenObjectHiddenDesign = True
            End If
    End Select
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "OpenObjectHiddenDesign", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Private Sub CloseObjectNoSave(ByVal ObjectType As String, ByVal ObjectName As String)
    On Error Resume Next

    Select Case UCase$(Trim$(ObjectType))
        Case OBJECT_TYPE_FORM
            If ObjectExists(OBJECT_TYPE_FORM, ObjectName) Then
                If CurrentProject.AllForms(ObjectName).IsLoaded Then
                    DoCmd.Close acForm, ObjectName, acSaveNo
                End If
            End If

        Case OBJECT_TYPE_REPORT
            If ObjectExists(OBJECT_TYPE_REPORT, ObjectName) Then
                If CurrentProject.AllReports(ObjectName).IsLoaded Then
                    DoCmd.Close acReport, ObjectName, acSaveNo
                End If
            End If
    End Select
End Sub

Private Function IsObjectLoaded(ByVal ObjectType As String, ByVal ObjectName As String) As Boolean
    On Error GoTo SafeExit

    Select Case NormalizeObjectType(ObjectType)
        Case OBJECT_TYPE_FORM
            IsObjectLoaded = CurrentProject.AllForms(ObjectName).IsLoaded
        Case OBJECT_TYPE_REPORT
            IsObjectLoaded = CurrentProject.AllReports(ObjectName).IsLoaded
    End Select

SafeExit:
End Function

Private Function GetObjectVisible(ByVal ObjectType As String, ByVal ObjectName As String) As Boolean
    On Error GoTo SafeExit

    If NormalizeObjectType(ObjectType) = OBJECT_TYPE_FORM Then
        If CurrentProject.AllForms(ObjectName).IsLoaded Then
            GetObjectVisible = Forms(ObjectName).Visible
        End If
    End If

SafeExit:
End Function

Private Function OpenObjectForUpdate(ByVal ObjectType As String, ByVal ObjectName As String) As Boolean
    On Error GoTo ErrorHandler

    OpenObjectForUpdate = False

    Select Case NormalizeObjectType(ObjectType)
        Case OBJECT_TYPE_FORM
            If CurrentProject.AllForms(ObjectName).IsLoaded Then
                DoCmd.Close acForm, ObjectName, acSaveYes
            End If
            DoCmd.OpenForm ObjectName, acDesign, , , , acHidden
            OpenObjectForUpdate = True

        Case OBJECT_TYPE_REPORT
            If CurrentProject.AllReports(ObjectName).IsLoaded Then
                DoCmd.Close acReport, ObjectName, acSaveYes
            End If
            DoCmd.OpenReport ObjectName, acViewDesign, , , acHidden
            OpenObjectForUpdate = True
    End Select
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "OpenObjectForUpdate", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Function

Private Sub SaveOpenedObject(ByVal ObjectType As String, ByVal ObjectName As String)
    On Error GoTo ErrorHandler

    Select Case NormalizeObjectType(ObjectType)
        Case OBJECT_TYPE_FORM
            DoCmd.Save acForm, ObjectName
        Case OBJECT_TYPE_REPORT
            DoCmd.Save acReport, ObjectName
    End Select
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SaveOpenedObject", Err
    Err.Raise Err.Number, Err.Source, Err.description
End Sub

Private Sub RestoreObjectState( _
    ByVal ObjectType As String, _
    ByVal ObjectName As String, _
    ByVal WasLoaded As Boolean, _
    ByVal WasVisible As Boolean)
    On Error GoTo SafeExit

    If Not WasLoaded Then
        Exit Sub
    End If

    Select Case NormalizeObjectType(ObjectType)
        Case OBJECT_TYPE_FORM
            If Not CurrentProject.AllForms(ObjectName).IsLoaded Then
                DoCmd.OpenForm ObjectName, acNormal
            End If
            If Not WasVisible Then
                Forms(ObjectName).Visible = False
            End If

        Case OBJECT_TYPE_REPORT
            If Not CurrentProject.AllReports(ObjectName).IsLoaded Then
                DoCmd.OpenReport ObjectName, acViewPreview
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

Private Function ControlExists(ByVal ObjectType As String, ByVal ObjectName As String, ByVal ControlName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ctl As Control

    Select Case NormalizeObjectType(ObjectType)
        Case OBJECT_TYPE_FORM
            For Each ctl In Forms(ObjectName).Controls
                If StrComp(ctl.Name, ControlName, vbTextCompare) = 0 Then
                    ControlExists = True
                    Exit Function
                End If
            Next ctl

        Case OBJECT_TYPE_REPORT
            For Each ctl In Reports(ObjectName).Controls
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

Private Function GetObjectPropertySafely(ByVal Obj As Object, ByVal PropertyName As String) As String
    On Error GoTo SafeExit

    GetObjectPropertySafely = Nz(Obj.Properties(PropertyName).Value, vbNullString)
    Exit Function

SafeExit:
    GetObjectPropertySafely = vbNullString
End Function

Private Function GetControlCaptionForLoadedObject( _
    ByVal ObjectType As String, _
    ByVal ObjectName As String, _
    ByVal ControlName As String _
) As String
    On Error GoTo SafeExit

    Select Case NormalizeObjectType(ObjectType)
        Case OBJECT_TYPE_FORM
            GetControlCaptionForLoadedObject = GetObjectPropertySafely(Forms(ObjectName).Controls(ControlName), "Caption")
        Case OBJECT_TYPE_REPORT
            GetControlCaptionForLoadedObject = GetObjectPropertySafely(Reports(ObjectName).Controls(ControlName), "Caption")
    End Select
    Exit Function

SafeExit:
    GetControlCaptionForLoadedObject = vbNullString
End Function

Private Function SetControlCaptionForLoadedObject( _
    ByVal ObjectType As String, _
    ByVal ObjectName As String, _
    ByVal ControlName As String, _
    ByVal CaptionValue As String _
) As Boolean
    On Error GoTo SafeExit

    Select Case NormalizeObjectType(ObjectType)
        Case OBJECT_TYPE_FORM
            Forms(ObjectName).Controls(ControlName).Properties("Caption").Value = CaptionValue
        Case OBJECT_TYPE_REPORT
            Reports(ObjectName).Controls(ControlName).Properties("Caption").Value = CaptionValue
    End Select

    SetControlCaptionForLoadedObject = True
    Exit Function

SafeExit:
    SetControlCaptionForLoadedObject = False
End Function

Private Function EnsureTranslationPlaceholderRow( _
    ByVal db As DAO.Database, _
    ByVal TranslationKey As String, _
    ByVal LanguageCode As String _
) As Boolean
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset

    If db Is Nothing Then
        Exit Function
    End If

    If TranslationRowExists(db, TranslationKey, LanguageCode) Then
        Exit Function
    End If

    Set rs = db.OpenRecordset(TRANSLATION_TABLE_NAME, dbOpenDynaset)

    rs.AddNew
    rs.Fields(FIELD_TRANSLATION_KEY).Value = TranslationKey
    rs.Fields(FIELD_LANGUAGE_CODE).Value = LanguageCode
    rs.Fields(FIELD_TRANSLATION_VALUE).Value = PLACEHOLDER_TRANSLATION_VALUE
    SetRecordsetFieldIfExists rs, FIELD_IS_ACTIVE, True
    SetRecordsetFieldIfExists rs, FIELD_MODULE_CODE, GetTranslationModuleCode(TranslationKey)
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
    ByVal TranslationKey As String, _
    ByVal LanguageCode As String _
) As Boolean
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT TOP 1 " & FIELD_TRANSLATION_KEY & _
          " FROM " & TRANSLATION_TABLE_NAME & _
          " WHERE " & FIELD_TRANSLATION_KEY & " = " & SqlText(TranslationKey) & _
          " AND " & FIELD_LANGUAGE_CODE & " = " & SqlText(LanguageCode)

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
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
    ByVal FieldName As String, _
    ByVal FieldValue As Variant _
)
    On Error GoTo SafeExit

    Dim fld As DAO.Field

    If rs Is Nothing Then
        Exit Sub
    End If

    For Each fld In rs.Fields
        If StrComp(fld.Name, FieldName, vbTextCompare) = 0 Then
            fld.Value = FieldValue
            Exit For
        End If
    Next fld

SafeExit:
End Sub

Private Function GetTranslationModuleCode(ByVal TranslationKey As String) As String
    TranslationKey = UCase$(Trim$(TranslationKey))

    If Left$(TranslationKey, 7) = "REPORT." Then
        GetTranslationModuleCode = "REPORT"
    ElseIf Left$(TranslationKey, 5) = "FORM." Then
        GetTranslationModuleCode = "FORM"
    Else
        GetTranslationModuleCode = "FRAMEWORK"
    End If
End Function

Private Function SqlText(ByVal Value As String) As String
    SqlText = "'" & Replace(Nz(Value, vbNullString), "'", "''") & "'"
End Function

Private Function BuildMissingTranslationWhereClause( _
    ByVal ObjectType As String, _
    ByVal ObjectName As String _
) As String
    Dim scopeClause As String

    scopeClause = BuildTranslationScopeWhereClause(ObjectType, ObjectName)
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
    ByVal ObjectType As String, _
    ByVal ObjectName As String _
) As String
    Dim normalizedObjectType As String
    Dim normalizedObjectName As String

    normalizedObjectType = vbNullString
    normalizedObjectName = UCase$(Trim$(ObjectName))

    If LenB(Trim$(ObjectType)) > 0 Then
        normalizedObjectType = NormalizeObjectType(ObjectType)
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
    Dim sql As String

    sql = "SELECT Count(*) AS MissingCount FROM " & TRANSLATION_TABLE_NAME
    If LenB(Trim$(whereClause)) > 0 Then
        sql = sql & " WHERE " & whereClause
    End If

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
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
    ByVal ObjectType As String, _
    ByVal ObjectName As String _
) As String
    Dim normalizedObjectType As String
    Dim normalizedObjectName As String

    normalizedObjectType = UCase$(Trim$(ObjectType))
    normalizedObjectName = Trim$(ObjectName)

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

Private Function TableExists(ByVal db As DAO.Database, ByVal TableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tdf As DAO.tableDef

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


