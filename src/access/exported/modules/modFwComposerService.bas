Attribute VB_Name = "modFwComposerService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFwComposerService
' Purpose   : Provides object and control inspection helpers for a future
'             framework composer form.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modFwComposerService"

Public Const COMPOSER_MODE_TAGS As String = "TAGS"
Public Const COMPOSER_MODE_TRANSLATIONS As String = "TRANSLATIONS"
Public Const OBJECT_TYPE_FORM As String = "FORM"
Public Const OBJECT_TYPE_REPORT As String = "REPORT"

Public Function GetComposerObjectList(ByVal ObjectType As String) As Collection
    On Error GoTo ErrorHandler

    Dim normalizedObjectType As String
    Dim result As Collection
    Dim accessObject As Access.AccessObject

    normalizedObjectType = NormalizeObjectType(ObjectType)
    Set result = New Collection

    Select Case normalizedObjectType
        Case OBJECT_TYPE_FORM
            For Each accessObject In CurrentProject.AllForms
                If Not accessObject.IsLoaded Or accessObject.IsLoaded Then
                    If Left$(accessObject.Name, 1) <> "~" Then
                        CollectionAddSorted result, accessObject.Name
                    End If
                End If
            Next accessObject

        Case OBJECT_TYPE_REPORT
            For Each accessObject In CurrentProject.AllReports
                If Not accessObject.IsLoaded Or accessObject.IsLoaded Then
                    If Left$(accessObject.Name, 1) <> "~" Then
                        CollectionAddSorted result, accessObject.Name
                    End If
                End If
            Next accessObject
    End Select

    Set GetComposerObjectList = result
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GetComposerObjectList", Err
    Err.Raise Err.Number, Err.Source, Err.Description
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

    normalizedObjectType = NormalizeObjectType(ObjectType)
    normalizedPrefix = Trim$(OnlyNamedPrefix)
    Set result = New Collection

    OpenObjectHiddenDesign normalizedObjectType, ObjectName

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
    CloseObjectNoSave normalizedObjectType, ObjectName
    Set GetComposerControlList = result
    Exit Function

ErrorHandler:
    CloseObjectNoSave normalizedObjectType, ObjectName
    modErrorHandler.HandleError MODULE_NAME, "GetComposerControlList", Err
    Err.Raise Err.Number, Err.Source, Err.Description
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
    Err.Raise Err.Number, Err.Source, Err.Description
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
    Err.Raise Err.Number, Err.Source, Err.Description
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
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetControlStringProperty( _
    ByVal ObjectType As String, _
    ByVal ObjectName As String, _
    ByVal ControlName As String, _
    ByVal PropertyName As String _
) As String
    On Error GoTo ErrorHandler

    Dim normalizedObjectType As String

    normalizedObjectType = NormalizeObjectType(ObjectType)
    OpenObjectHiddenDesign normalizedObjectType, ObjectName

    If Not ControlExists(normalizedObjectType, ObjectName, ControlName) Then
        Err.Raise vbObjectError + 3200, MODULE_NAME & ".GetControlStringProperty", _
            "Control '" & ControlName & "' does not exist on " & normalizedObjectType & " '" & ObjectName & "'."
    End If

    Select Case normalizedObjectType
        Case OBJECT_TYPE_FORM
            GetControlStringProperty = GetObjectPropertySafely(Forms(ObjectName).Controls(ControlName), PropertyName)
        Case OBJECT_TYPE_REPORT
            GetControlStringProperty = GetObjectPropertySafely(Reports(ObjectName).Controls(ControlName), PropertyName)
    End Select

CleanExit:
    CloseObjectNoSave normalizedObjectType, ObjectName
    Exit Function

ErrorHandler:
    CloseObjectNoSave normalizedObjectType, ObjectName
    modErrorHandler.HandleError MODULE_NAME, "GetControlStringProperty", Err
    Err.Raise Err.Number, Err.Source, Err.Description
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

    Do While Left$(result, 1) = "_"
        result = Mid$(result, 2)
    Loop

    Do While Right$(result, 1) = "_"
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

    Dim accessObject As Access.AccessObject
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

Private Sub OpenObjectHiddenDesign(ByVal ObjectType As String, ByVal ObjectName As String)
    On Error GoTo ErrorHandler

    If Not ObjectExists(ObjectType, ObjectName) Then
        Err.Raise vbObjectError + 3202, MODULE_NAME & ".OpenObjectHiddenDesign", _
            ObjectType & " '" & ObjectName & "' does not exist."
    End If

    Select Case NormalizeObjectType(ObjectType)
        Case OBJECT_TYPE_FORM
            If Not CurrentProject.AllForms(ObjectName).IsLoaded Then
                DoCmd.OpenForm ObjectName, acDesign, , , , acHidden
            End If

        Case OBJECT_TYPE_REPORT
            If Not CurrentProject.AllReports(ObjectName).IsLoaded Then
                DoCmd.OpenReport ObjectName, acViewDesign, , , acHidden
            End If
    End Select
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "OpenObjectHiddenDesign", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

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

Private Sub CollectionAddSorted(ByVal target As Collection, ByVal ItemValue As String)
    On Error GoTo ErrorHandler

    Dim tempValues() As String
    Dim i As Long
    Dim j As Long
    Dim inserted As Boolean
    Dim currentValue As String

    If target Is Nothing Then
        Exit Sub
    End If

    If target.Count = 0 Then
        target.Add ItemValue
        Exit Sub
    End If

    ReDim tempValues(1 To target.Count + 1)

    For i = 1 To target.Count
        tempValues(i) = CStr(target(i))
    Next i

    tempValues(target.Count + 1) = ItemValue

    For i = LBound(tempValues) To UBound(tempValues) - 1
        For j = i + 1 To UBound(tempValues)
            If StrComp(tempValues(i), tempValues(j), vbTextCompare) > 0 Then
                currentValue = tempValues(i)
                tempValues(i) = tempValues(j)
                tempValues(j) = currentValue
            End If
        Next j
    Next i

    Do While target.Count > 0
        target.Remove 1
    Loop

    For i = LBound(tempValues) To UBound(tempValues)
        If Not inserted Or StrComp(currentValue, tempValues(i), vbTextCompare) <> 0 Then
            target.Add tempValues(i)
            inserted = True
        End If
        currentValue = tempValues(i)
    Next i

    Exit Sub

ErrorHandler:
    target.Add ItemValue
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

