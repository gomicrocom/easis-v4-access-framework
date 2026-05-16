Attribute VB_Name = "modFwValidationRuntime"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFwValidationRuntime
' Purpose   : Runtime validation for framework rules stored in control tags.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modFwValidationRuntime"
Private Const TAG_REQUIRED As String = "REQUIRED"
Private Const TAG_MINLEN As String = "MINLEN"
Private Const TAG_MAXLEN As String = "MAXLEN"
Private Const TAG_NUMERIC As String = "NUMERIC"
Private Const TAG_INTEGER As String = "INTEGER"
Private Const TAG_DATE As String = "DATE"

Public Function ValidateForm(ByVal frm As Form) As Boolean
    On Error GoTo ErrorHandler

    Dim ctl As Control

    ValidateForm = True

    If frm Is Nothing Then
        Exit Function
    End If

    For Each ctl In frm.Controls
        If ControlSupportsValidation(ctl) Then
            If LenB(Trim$(NzString(ctl.Tag))) > 0 Then
                If Not ValidateControl(ctl) Then
                    TryFocusControl ctl
                    ValidateForm = False
                    Exit Function
                End If
            End If
        End If
    Next ctl

    Exit Function

ErrorHandler:
    ValidateForm = False
    modErrorHandler.HandleError MODULE_NAME, "ValidateForm", Err
End Function

Public Function ValidateControl(ByVal ctl As Control) As Boolean
    On Error GoTo ErrorHandler

    Dim tagTokens As Object
    Dim displayName As String
    Dim rawValue As Variant
    Dim textValue As String
    Dim minLen As Long
    Dim maxLen As Long
    Dim messageText As String

    ValidateControl = True

    If ctl Is Nothing Then
        Exit Function
    End If

    If Not ControlSupportsValidation(ctl) Then
        Exit Function
    End If

    Set tagTokens = ParseTagTokens(NzString(ctl.Tag))
    If tagTokens Is Nothing Then
        Exit Function
    End If

    If tagTokens.Count = 0 Then
        Exit Function
    End If

    displayName = GetControlDisplayName(ctl)
    rawValue = GetControlValue(ctl)
    textValue = Trim$(NzString(rawValue))

    If HasTag(tagTokens, TAG_REQUIRED) Then
        If IsMissingValue(ctl, rawValue) Then
            messageText = "Feld '" & displayName & "' ist erforderlich."
            ShowValidationFailure ctl, messageText
            ValidateControl = False
            Exit Function
        End If
    End If

    If LenB(textValue) = 0 Then
        Exit Function
    End If

    If HasTag(tagTokens, TAG_MINLEN) Then
        If TryParseLong(GetTagValue(tagTokens, TAG_MINLEN), minLen, ctl.Name, TAG_MINLEN) Then
            If Len(textValue) < minLen Then
                messageText = "Feld '" & displayName & "' muss mindestens " & CStr(minLen) & " Zeichen enthalten."
                ShowValidationFailure ctl, messageText
                ValidateControl = False
                Exit Function
            End If
        End If
    End If

    If HasTag(tagTokens, TAG_MAXLEN) Then
        If TryParseLong(GetTagValue(tagTokens, TAG_MAXLEN), maxLen, ctl.Name, TAG_MAXLEN) Then
            If Len(textValue) > maxLen Then
                messageText = "Feld '" & displayName & "' darf maximal " & CStr(maxLen) & " Zeichen enthalten."
                ShowValidationFailure ctl, messageText
                ValidateControl = False
                Exit Function
            End If
        End If
    End If

    If HasTag(tagTokens, TAG_NUMERIC) Then
        If Not IsNumeric(textValue) Then
            messageText = "Feld '" & displayName & "' muss numerisch sein."
            ShowValidationFailure ctl, messageText
            ValidateControl = False
            Exit Function
        End If
    End If

    If HasTag(tagTokens, TAG_INTEGER) Then
        If Not IsNumeric(textValue) Or CLng(CDbl(textValue)) <> CDbl(textValue) Then
            messageText = "Feld '" & displayName & "' muss eine ganze Zahl enthalten."
            ShowValidationFailure ctl, messageText
            ValidateControl = False
            Exit Function
        End If
    End If

    If HasTag(tagTokens, TAG_DATE) Then
        If Not IsDate(textValue) Then
            messageText = "Feld '" & displayName & "' muss ein gueltiges Datum enthalten."
            ShowValidationFailure ctl, messageText
            ValidateControl = False
            Exit Function
        End If
    End If

    Exit Function

ErrorHandler:
    ValidateControl = False
    modErrorHandler.HandleError MODULE_NAME, "ValidateControl", Err
End Function

Private Function ParseTagTokens(ByVal TagValue As String) As Object
    On Error GoTo ErrorHandler

    Dim tokens As Object
    Dim parts() As String
    Dim partValue As Variant
    Dim tokenText As String
    Dim tokenName As String
    Dim tokenValue As String
    Dim separatorPosition As Long

    Set tokens = CreateObject("Scripting.Dictionary")
    tokens.CompareMode = vbTextCompare

    TagValue = Trim$(TagValue)
    If LenB(TagValue) = 0 Then
        Set ParseTagTokens = tokens
        Exit Function
    End If

    parts = Split(TagValue, ";")
    For Each partValue In parts
        tokenText = Trim$(CStr(partValue))
        If LenB(tokenText) = 0 Then
            GoTo NextToken
        End If

        separatorPosition = InStr(1, tokenText, ":", vbTextCompare)
        If separatorPosition > 0 Then
            tokenName = UCase$(Trim$(Left$(tokenText, separatorPosition - 1)))
            tokenValue = Trim$(Mid$(tokenText, separatorPosition + 1))
        Else
            tokenName = UCase$(tokenText)
            tokenValue = vbNullString
        End If

        Select Case tokenName
            Case TAG_REQUIRED, TAG_MINLEN, TAG_MAXLEN, TAG_NUMERIC, TAG_INTEGER, TAG_DATE
                tokens(tokenName) = tokenValue
            Case Else
                modLoggingHandler.LogInfo MODULE_NAME & ".ParseTagTokens", _
                    "Ignoring unsupported validation token '" & tokenName & "'."
        End Select

NextToken:
    Next partValue

    Set ParseTagTokens = tokens
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ParseTagTokens", Err
    Set ParseTagTokens = Nothing
End Function

Private Function HasTag(ByVal tagTokens As Object, ByVal TagName As String) As Boolean
    On Error GoTo SafeExit

    If tagTokens Is Nothing Then
        Exit Function
    End If

    HasTag = tagTokens.Exists(UCase$(Trim$(TagName)))

SafeExit:
End Function

Private Function GetTagValue(ByVal tagTokens As Object, ByVal TagName As String) As String
    On Error GoTo SafeExit

    If tagTokens Is Nothing Then
        Exit Function
    End If

    If tagTokens.Exists(UCase$(Trim$(TagName))) Then
        GetTagValue = NzString(tagTokens(UCase$(Trim$(TagName))))
    End If

SafeExit:
End Function

Private Function GetControlDisplayName(ByVal ctl As Control) As String
    On Error GoTo ErrorHandler

    Dim captionValue As String

    If ctl Is Nothing Then
        Exit Function
    End If

    If ctl.ControlType = acCheckBox Then
        captionValue = GetCaptionPropertySafely(ctl)
        If LenB(captionValue) > 0 Then
            GetControlDisplayName = ResolveDisplayCaption(captionValue)
            Exit Function
        End If
    End If

    captionValue = GetAttachedLabelCaption(ctl)
    If LenB(captionValue) > 0 Then
        GetControlDisplayName = ResolveDisplayCaption(captionValue)
        Exit Function
    End If

    captionValue = GetCaptionPropertySafely(ctl)
    If LenB(captionValue) > 0 Then
        GetControlDisplayName = ResolveDisplayCaption(captionValue)
        Exit Function
    End If

    GetControlDisplayName = ctl.Name
    Exit Function

ErrorHandler:
    GetControlDisplayName = ctl.Name
    modErrorHandler.HandleError MODULE_NAME, "GetControlDisplayName", Err
End Function

Private Function ResolveDisplayCaption(ByVal CaptionValue As String) As String
    CaptionValue = Trim$(NzString(CaptionValue))

    If LenB(CaptionValue) = 0 Then
        Exit Function
    End If

    If StrComp(Left$(CaptionValue, 3), "TR:", vbTextCompare) = 0 Then
        ResolveDisplayCaption = modFwTranslationRuntime.ResolveTranslation(CaptionValue, modFwTranslationRuntime.GetCurrentLanguageCode())
    Else
        ResolveDisplayCaption = CaptionValue
    End If
End Function

Private Function GetAttachedLabelCaption(ByVal ctl As Control) As String
    On Error GoTo SafeExit

    If ctl Is Nothing Then
        Exit Function
    End If

    If ctl.Controls.Count > 0 Then
        GetAttachedLabelCaption = GetCaptionPropertySafely(ctl.Controls(0))
    End If
    Exit Function

SafeExit:
    GetAttachedLabelCaption = vbNullString
End Function

Private Function GetCaptionPropertySafely(ByVal ctl As Control) As String
    On Error GoTo SafeExit

    If ctl Is Nothing Then
        Exit Function
    End If

    GetCaptionPropertySafely = NzString(ctl.Properties("Caption").Value)
    Exit Function

SafeExit:
    GetCaptionPropertySafely = vbNullString
End Function

Private Function ControlSupportsValidation(ByVal ctl As Control) As Boolean
    On Error GoTo SafeExit

    If ctl Is Nothing Then
        Exit Function
    End If

    Select Case ctl.ControlType
        Case acTextBox, acComboBox, acCheckBox
            ControlSupportsValidation = True
    End Select

SafeExit:
End Function

Private Function GetControlValue(ByVal ctl As Control) As Variant
    On Error GoTo SafeExit

    If ctl Is Nothing Then
        GetControlValue = Null
        Exit Function
    End If

    GetControlValue = ctl.Value
    Exit Function

SafeExit:
    GetControlValue = Null
End Function

Private Function IsMissingValue(ByVal ctl As Control, ByVal rawValue As Variant) As Boolean
    If ctl Is Nothing Then
        Exit Function
    End If

    Select Case ctl.ControlType
        Case acCheckBox
            IsMissingValue = IsNull(rawValue)
        Case Else
            IsMissingValue = (LenB(Trim$(NzString(rawValue))) = 0)
    End Select
End Function

Private Function TryParseLong( _
    ByVal TextValue As String, _
    ByRef ParsedValue As Long, _
    ByVal ControlName As String, _
    ByVal TagName As String) As Boolean
    On Error GoTo ErrorHandler

    TextValue = Trim$(NzString(TextValue))
    If LenB(TextValue) = 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".TryParseLong", _
            "Invalid tag format for " & ControlName & ": " & TagName & " requires a numeric value."
        Exit Function
    End If

    If Not IsNumeric(TextValue) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".TryParseLong", _
            "Invalid tag format for " & ControlName & ": " & TagName & "='" & TextValue & "'."
        Exit Function
    End If

    ParsedValue = CLng(TextValue)
    TryParseLong = True
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "TryParseLong", Err
End Function

Private Sub ShowValidationFailure(ByVal ctl As Control, ByVal messageText As String)
    modLoggingHandler.LogWarning MODULE_NAME & ".ShowValidationFailure", _
        "Validation failed for control '" & ctl.Name & "': " & messageText
    MsgBox messageText, vbExclamation, "Validierung"
End Sub

Private Sub TryFocusControl(ByVal ctl As Control)
    On Error GoTo SafeExit

    If ctl Is Nothing Then
        Exit Sub
    End If

    ctl.SetFocus

SafeExit:
End Sub

Private Function NzString(ByVal Value As Variant, Optional ByVal DefaultValue As String = "") As String
    If IsNull(Value) Or IsEmpty(Value) Then
        NzString = DefaultValue
    Else
        NzString = CStr(Value)
    End If
End Function

' Example integration:
' Private Sub Form_BeforeUpdate(Cancel As Integer)
'     If Not modFwValidationRuntime.ValidateForm(Me) Then
'         Cancel = True
'     End If
' End Sub
