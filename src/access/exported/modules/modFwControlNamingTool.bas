Attribute VB_Name = "modFwControlNamingTool"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFwControlNamingTool
' Purpose   : Developer helper to normalize bound Access controls and labels
'             to the Easis naming convention.
' Author    : Codex
' Version   : 0.1.0
' Notes     : Manual tool. Not part of bootstrap/runtime.
'===============================================================================

Private Const MODULE_NAME As String = "modFwControlNamingTool"
Private Const OBJECT_TYPE_FORM As String = "FORM"
Private Const OBJECT_TYPE_REPORT As String = "REPORT"

Private Const PREFIX_TEXTBOX As String = "txt"
Private Const PREFIX_COMBOBOX As String = "cbo"
Private Const PREFIX_CHECKBOX As String = "chk"
Private Const PREFIX_LABEL As String = "lbl"

Public Sub PreviewNormalizeControls(ByVal objectName As String, Optional ByVal objectType As String = OBJECT_TYPE_FORM)
    NormalizeControlsInternal objectName, objectType, False
End Sub

Public Sub ApplyNormalizeControls(ByVal objectName As String, Optional ByVal objectType As String = OBJECT_TYPE_FORM)
    NormalizeControlsInternal objectName, objectType, True
End Sub

Public Function PascalCaseFromSnake(ByVal fieldName As String) As String
    Dim cleanedValue As String
    Dim parts() As String
    Dim part As Variant
    Dim result As String

    cleanedValue = Trim$(Nz(fieldName, vbNullString))
    cleanedValue = Replace(cleanedValue, "[", vbNullString)
    cleanedValue = Replace(cleanedValue, "]", vbNullString)
    cleanedValue = Replace(cleanedValue, ".", "_")
    cleanedValue = Replace(cleanedValue, "-", "_")
    cleanedValue = Replace(cleanedValue, " ", "_")

    Do While InStr(cleanedValue, "__") > 0
        cleanedValue = Replace(cleanedValue, "__", "_")
    Loop

    If Left$(cleanedValue, 1) = "_" Then cleanedValue = Mid$(cleanedValue, 2)
    If Right$(cleanedValue, 1) = "_" Then cleanedValue = Left$(cleanedValue, Len(cleanedValue) - 1)

    If LenB(cleanedValue) = 0 Then
        Exit Function
    End If

    parts = Split(cleanedValue, "_")
    For Each part In parts
        If LenB(CStr(part)) > 0 Then
            result = result & UCase$(Left$(CStr(part), 1)) & LCase$(Mid$(CStr(part), 2))
        End If
    Next part

    PascalCaseFromSnake = result
End Function

Public Function IsGenericAccessLabelName(ByVal controlName As String) As Boolean
    controlName = Trim$(Nz(controlName, vbNullString))

    If LenB(controlName) = 0 Then
        Exit Function
    End If

    IsGenericAccessLabelName = _
        (LCase$(controlName) Like "bezeichnungsfeld*") Or _
        (LCase$(controlName) Like "label*")
End Function

Public Function GetExpectedControlPrefix(ByVal ctl As Access.Control) As String
    Select Case ctl.ControlType
        Case acTextBox
            GetExpectedControlPrefix = PREFIX_TEXTBOX
        Case acComboBox
            GetExpectedControlPrefix = PREFIX_COMBOBOX
        Case acCheckBox
            GetExpectedControlPrefix = PREFIX_CHECKBOX
    End Select
End Function

Public Function GetAttachedLabel(ByVal ctl As Access.Control) As Object
    On Error GoTo SafeExit

    Set GetAttachedLabel = CallByName(ctl, "AttachedLabel", VbGet)

SafeExit:
End Function

Private Sub NormalizeControlsInternal(ByVal objectName As String, ByVal objectType As String, ByVal applyChanges As Boolean)
    On Error GoTo ErrorHandler

    Dim normalizedObjectType As String
    Dim wasLoaded As Boolean
    Dim wasVisible As Boolean
    Dim openedByTool As Boolean
    Dim target As Object
    Dim plans As Collection
    Dim plan As Object
    Dim i As Long
    Dim changedCount As Long

    normalizedObjectType = NormalizeObjectType(objectType)

    If Not ObjectExists(normalizedObjectType, objectName) Then
        Err.Raise vbObjectError + 6500, MODULE_NAME & ".NormalizeControlsInternal", _
            normalizedObjectType & " '" & objectName & "' does not exist."
    End If

    wasLoaded = IsObjectLoaded(normalizedObjectType, objectName)
    wasVisible = GetObjectVisible(normalizedObjectType, objectName)
    openedByTool = OpenObjectForNaming(normalizedObjectType, objectName)

    Set target = GetLoadedObject(normalizedObjectType, objectName)
    If target Is Nothing Then
        Err.Raise vbObjectError + 6501, MODULE_NAME & ".NormalizeControlsInternal", _
            "Could not resolve loaded " & LCase$(normalizedObjectType) & " '" & objectName & "'."
    End If

    Set plans = BuildRenamePlan(target)

    LogToolInfo IIf(applyChanges, "APPLY", "PREVIEW"), _
        normalizedObjectType & " '" & objectName & "' analyzed. plan_count=" & CStr(plans.Count) & "."

    For i = 1 To plans.Count
        Set plan = plans(i)
        LogRenamePlan plan
        If applyChanges Then
            If ApplyRenamePlan(target, plan) Then
                changedCount = changedCount + 1
            End If
        End If
    Next i

    If applyChanges Then
        SaveOpenedObject normalizedObjectType, objectName
        LogToolInfo "APPLY", normalizedObjectType & " '" & objectName & "' saved. changed_count=" & CStr(changedCount) & "."
    End If

CleanExit:
    On Error Resume Next
    If openedByTool Then
        CloseOpenedObject normalizedObjectType, objectName, applyChanges
    ElseIf wasLoaded And wasVisible = False And normalizedObjectType = OBJECT_TYPE_FORM Then
        If CurrentProject.AllForms(objectName).IsLoaded Then
            Forms(objectName).Visible = False
        End If
    End If
    Set target = Nothing
    Set plans = Nothing
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "NormalizeControlsInternal", Err
    Resume CleanExit
End Sub

Private Function BuildRenamePlan(ByVal target As Object) As Collection
    On Error GoTo ErrorHandler

    Dim plans As Collection
    Dim ctl As Control
    Dim plan As Object
    Dim expectedPrefix As String
    Dim expectedControlName As String
    Dim sourceName As String
    Dim attachedLabel As Object
    Dim expectedLabelName As String
    Dim labelCaption As String

    Set plans = New Collection

    For Each ctl In target.Controls
        Set plan = CreateRenamePlan()

        expectedPrefix = GetExpectedControlPrefix(ctl)
        If LenB(expectedPrefix) = 0 Then
            GoTo NextControl
        End If

        sourceName = Trim$(Nz(GetControlSourceValue(ctl), vbNullString))
        If LenB(sourceName) = 0 Then
            GoTo NextControl
        End If

        If Left$(sourceName, 1) = "=" Then
            plan("ControlOldName") = ctl.Name
            plan("SkipReason") = "calculated_control"
            plans.Add plan
            GoTo NextControl
        End If

        expectedControlName = expectedPrefix & PascalCaseFromSnake(sourceName)
        If LenB(expectedControlName) = 0 Then
            plan("ControlOldName") = ctl.Name
            plan("SkipReason") = "empty_expected_name"
            plans.Add plan
            GoTo NextControl
        End If

        plan("ControlOldName") = ctl.Name
        plan("ControlNewName") = expectedControlName
        plan("HasControlRename") = (StrComp(ctl.Name, expectedControlName, vbTextCompare) <> 0)

        If plan("HasControlRename") Then
            If NameExistsOnObject(target, expectedControlName, ctl.Name) Then
                plan("SkipReason") = "control_name_conflict"
            End If
        End If

        Set attachedLabel = GetAttachedLabel(ctl)
        If Not attachedLabel Is Nothing Then
            plan("LabelOldName") = Nz(attachedLabel.Name, vbNullString)
            expectedLabelName = PREFIX_LABEL & PascalCaseFromSnake(sourceName)
            plan("LabelNewName") = expectedLabelName

            If LenB(expectedLabelName) > 0 Then
                If IsGenericAccessLabelName(plan("LabelOldName")) Or StrComp(plan("LabelOldName"), expectedLabelName, vbTextCompare) = 0 Then
                    If StrComp(plan("LabelOldName"), expectedLabelName, vbTextCompare) <> 0 Then
                        If Not NameExistsOnObject(target, expectedLabelName, plan("LabelOldName")) Then
                            plan("HasLabelRename") = True
                        Else
                            AppendSkipReason plan, "label_name_conflict"
                        End If
                    End If
                End If
            End If

            labelCaption = Nz(GetCaptionValue(attachedLabel), vbNullString)
            If IsGenericLabelCaption(labelCaption, plan("LabelOldName")) Then
                plan("LabelNewCaption") = ReadableCaptionFromFieldName(sourceName)
                If LenB(plan("LabelNewCaption")) > 0 And StrComp(labelCaption, plan("LabelNewCaption"), vbBinaryCompare) <> 0 Then
                    plan("HasLabelCaptionUpdate") = True
                End If
            End If
        End If

        plans.Add plan

NextControl:
        Set attachedLabel = Nothing
    Next ctl

    Set BuildRenamePlan = plans
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "BuildRenamePlan", Err
End Function

Private Function ApplyRenamePlan(ByVal target As Object, ByVal plan As Object) As Boolean
    On Error GoTo ErrorHandler

    Dim ctl As Control
    Dim attachedLabel As Object
    Dim anyChange As Boolean

    If LenB(plan("SkipReason")) > 0 Then
        Exit Function
    End If

    Set ctl = target.Controls(plan("ControlOldName"))

    If plan("HasControlRename") Then
        ctl.Name = plan("ControlNewName")
        anyChange = True
        Set ctl = target.Controls(plan("ControlNewName"))
    End If

    If plan("HasLabelRename") Or plan("HasLabelCaptionUpdate") Then
        Set attachedLabel = GetAttachedLabel(ctl)
        If Not attachedLabel Is Nothing Then
            If plan("HasLabelRename") Then
                attachedLabel.Name = plan("LabelNewName")
                anyChange = True
                Set attachedLabel = target.Controls(plan("LabelNewName"))
            End If

            If plan("HasLabelCaptionUpdate") Then
                attachedLabel.Caption = plan("LabelNewCaption")
                anyChange = True
            End If
        End If
    End If

    ApplyRenamePlan = anyChange
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyRenamePlan", Err
End Function

Private Function CreateRenamePlan() As Object
    Dim plan As Object

    Set plan = CreateObject("Scripting.Dictionary")
    plan.CompareMode = vbTextCompare
    plan("ControlOldName") = vbNullString
    plan("ControlNewName") = vbNullString
    plan("LabelOldName") = vbNullString
    plan("LabelNewName") = vbNullString
    plan("LabelNewCaption") = vbNullString
    plan("HasControlRename") = False
    plan("HasLabelRename") = False
    plan("HasLabelCaptionUpdate") = False
    plan("SkipReason") = vbNullString
    Set CreateRenamePlan = plan
End Function

Private Sub AppendSkipReason(ByVal plan As Object, ByVal reasonText As String)
    If LenB(plan("SkipReason")) = 0 Then
        plan("SkipReason") = reasonText
    Else
        plan("SkipReason") = plan("SkipReason") & "," & reasonText
    End If
End Sub

Private Sub LogRenamePlan(ByVal plan As Object)
    Dim messageText As String

    messageText = "control='" & plan("ControlOldName") & "'"

    If plan("HasControlRename") Then
        messageText = messageText & " -> '" & plan("ControlNewName") & "'"
    End If

    If LenB(plan("LabelOldName")) > 0 Then
        messageText = messageText & "; label='" & plan("LabelOldName") & "'"
        If plan("HasLabelRename") Then
            messageText = messageText & " -> '" & plan("LabelNewName") & "'"
        End If
    End If

    If plan("HasLabelCaptionUpdate") Then
        messageText = messageText & "; caption='" & plan("LabelNewCaption") & "'"
    End If

    If LenB(plan("SkipReason")) > 0 Then
        LogToolWarning "SKIP", messageText & "; reason=" & plan("SkipReason") & "."
    ElseIf plan("HasControlRename") Or plan("HasLabelRename") Or plan("HasLabelCaptionUpdate") Then
        LogToolInfo "PLAN", messageText & "."
    End If
End Sub

Private Function ReadableCaptionFromFieldName(ByVal fieldName As String) As String
    Dim cleanedValue As String
    Dim parts() As String
    Dim part As Variant
    Dim captionText As String

    cleanedValue = Trim$(Nz(fieldName, vbNullString))
    cleanedValue = Replace(cleanedValue, "[", vbNullString)
    cleanedValue = Replace(cleanedValue, "]", vbNullString)
    cleanedValue = Replace(cleanedValue, ".", "_")
    cleanedValue = Replace(cleanedValue, "-", "_")
    cleanedValue = Replace(cleanedValue, " ", "_")

    Do While InStr(cleanedValue, "__") > 0
        cleanedValue = Replace(cleanedValue, "__", "_")
    Loop

    parts = Split(cleanedValue, "_")
    For Each part In parts
        If LenB(CStr(part)) > 0 Then
            If LenB(captionText) > 0 Then
                captionText = captionText & " "
            End If
            captionText = captionText & UCase$(Left$(CStr(part), 1)) & LCase$(Mid$(CStr(part), 2))
        End If
    Next part

    ReadableCaptionFromFieldName = captionText
End Function

Private Function IsGenericLabelCaption(ByVal captionText As String, ByVal labelName As String) As Boolean
    captionText = Trim$(Nz(captionText, vbNullString))

    If LenB(captionText) = 0 Then
        IsGenericLabelCaption = True
        Exit Function
    End If

    IsGenericLabelCaption = _
        StrComp(captionText, labelName, vbTextCompare) = 0 Or _
        LCase$(captionText) Like "bezeichnungsfeld*" Or _
        LCase$(captionText) Like "label*"
End Function

Private Function GetControlSourceValue(ByVal ctl As Control) As String
    On Error GoTo SafeExit

    GetControlSourceValue = Nz(CallByName(ctl, "ControlSource", VbGet), vbNullString)

SafeExit:
End Function

Private Function GetCaptionValue(ByVal ctl As Object) As String
    On Error GoTo SafeExit

    GetCaptionValue = Nz(CallByName(ctl, "Caption", VbGet), vbNullString)

SafeExit:
End Function

Private Function NameExistsOnObject(ByVal target As Object, ByVal candidateName As String, Optional ByVal ignoreName As String = "") As Boolean
    On Error GoTo SafeExit

    Dim ctl As Control

    For Each ctl In target.Controls
        If StrComp(ctl.Name, candidateName, vbTextCompare) = 0 Then
            If LenB(ignoreName) = 0 Or StrComp(ctl.Name, ignoreName, vbTextCompare) <> 0 Then
                NameExistsOnObject = True
                Exit Function
            End If
        End If
    Next ctl

SafeExit:
End Function

Private Function NormalizeObjectType(ByVal objectType As String) As String
    objectType = UCase$(Trim$(Nz(objectType, OBJECT_TYPE_FORM)))

    Select Case objectType
        Case OBJECT_TYPE_FORM, OBJECT_TYPE_REPORT
            NormalizeObjectType = objectType
        Case Else
            Err.Raise vbObjectError + 6502, MODULE_NAME & ".NormalizeObjectType", _
                "Unsupported objectType '" & objectType & "'. Expected FORM or REPORT."
    End Select
End Function

Private Function ObjectExists(ByVal objectType As String, ByVal objectName As String) As Boolean
    On Error GoTo SafeExit

    Dim accessObject As Access.AccessObject

    Select Case NormalizeObjectType(objectType)
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

SafeExit:
End Function

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

Private Function OpenObjectForNaming(ByVal objectType As String, ByVal objectName As String) As Boolean
    On Error GoTo ErrorHandler

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            If CurrentProject.AllForms(objectName).IsLoaded Then
                Err.Raise vbObjectError + 6503, MODULE_NAME & ".OpenObjectForNaming", _
                    "Close form '" & objectName & "' before running the control naming tool."
            End If
            DoCmd.OpenForm objectName, acDesign, , , , acHidden
            OpenObjectForNaming = True

        Case OBJECT_TYPE_REPORT
            If CurrentProject.AllReports(objectName).IsLoaded Then
                Err.Raise vbObjectError + 6504, MODULE_NAME & ".OpenObjectForNaming", _
                    "Close report '" & objectName & "' before running the control naming tool."
            End If
            DoCmd.OpenReport objectName, acViewDesign, , , acHidden
            OpenObjectForNaming = True
    End Select
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "OpenObjectForNaming", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetLoadedObject(ByVal objectType As String, ByVal objectName As String) As Object
    On Error GoTo SafeExit

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            If CurrentProject.AllForms(objectName).IsLoaded Then
                Set GetLoadedObject = Forms(objectName)
            End If
        Case OBJECT_TYPE_REPORT
            If CurrentProject.AllReports(objectName).IsLoaded Then
                Set GetLoadedObject = Reports(objectName)
            End If
    End Select

SafeExit:
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
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub CloseOpenedObject(ByVal objectType As String, ByVal objectName As String, ByVal saveChanges As Boolean)
    On Error Resume Next

    Select Case NormalizeObjectType(objectType)
        Case OBJECT_TYPE_FORM
            If CurrentProject.AllForms(objectName).IsLoaded Then
                DoCmd.Close acForm, objectName, IIf(saveChanges, acSaveYes, acSaveNo)
            End If
        Case OBJECT_TYPE_REPORT
            If CurrentProject.AllReports(objectName).IsLoaded Then
                DoCmd.Close acReport, objectName, IIf(saveChanges, acSaveYes, acSaveNo)
            End If
    End Select
End Sub

Private Sub LogToolInfo(ByVal category As String, ByVal messageText As String)
    Debug.Print Format$(Now(), "yyyy-mm-dd hh:nn:ss") & " | " & category & " | " & messageText
    On Error Resume Next
    modLoggingHandler.LogInfo MODULE_NAME & "." & category, messageText
End Sub

Private Sub LogToolWarning(ByVal category As String, ByVal messageText As String)
    Debug.Print Format$(Now(), "yyyy-mm-dd hh:nn:ss") & " | WARN | " & category & " | " & messageText
    On Error Resume Next
    modLoggingHandler.LogWarning MODULE_NAME & "." & category, messageText
End Sub
