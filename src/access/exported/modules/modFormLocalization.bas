Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFormLocalization
' Purpose   : Applies translation keys from Tag metadata to Access forms and controls.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modFormLocalization"
Private Const TAG_PREFIX_TRANSLATION As String = "TR:"

Public Sub LocalizeForm(ByVal FormInstance As Access.Form)
    On Error GoTo ErrorHandler

    Dim translationKey As String
    Dim localizedCount As Long
    Dim ctl As Control

    If FormInstance Is Nothing Then
        Exit Sub
    End If

    translationKey = ExtractTranslationKeyFromTag(FormInstance.Tag)
    If LenB(translationKey) > 0 Then
        SetFormCaption FormInstance, translationKey, NzString(FormInstance.Caption)
        localizedCount = localizedCount + 1
    End If

    For Each ctl In FormInstance.Controls
        translationKey = ExtractTranslationKeyFromTag(ctl.Tag)
        If LenB(translationKey) > 0 Then
            LocalizeControl ctl, translationKey, GetControlFallbackCaption(ctl)
            localizedCount = localizedCount + 1
        End If

        If ctl.ControlType = acTabCtl Then
            localizedCount = localizedCount + LocalizeTabPages(ctl)
        End If
    Next ctl

    If localizedCount > 0 Then
        modLoggingHandler.LogInfo MODULE_NAME & ".LocalizeForm", _
            "Localized " & CStr(localizedCount) & " element(s) on form '" & FormInstance.Name & "'."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LocalizeForm", Err
End Sub

Public Sub SetFormCaption(ByVal FormInstance As Access.Form, ByVal TranslationKey As String, Optional ByVal Fallback As String = "")
    On Error GoTo ErrorHandler

    If FormInstance Is Nothing Then
        Exit Sub
    End If

    If LenB(Trim$(TranslationKey)) = 0 Then
        Exit Sub
    End If

    FormInstance.Caption = modTranslationService.T(TranslationKey, Fallback)
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SetFormCaption", Err
End Sub

Public Sub LocalizeControl(ByVal ControlInstance As Control, ByVal TranslationKey As String, Optional ByVal Fallback As String = "")
    On Error GoTo ErrorHandler

    If ControlInstance Is Nothing Then
        Exit Sub
    End If

    If LenB(Trim$(TranslationKey)) = 0 Then
        Exit Sub
    End If

    If Not SupportsCaptionLocalization(ControlInstance) Then
        Exit Sub
    End If

    ApplyCaptionToControl ControlInstance, modTranslationService.T(TranslationKey, Fallback)
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LocalizeControl", Err
End Sub

Private Function ExtractTranslationKeyFromTag(ByVal TagValue As String) As String
    Dim trimmedTag As String

    trimmedTag = Trim$(TagValue)
    If LenB(trimmedTag) = 0 Then
        Exit Function
    End If

    If UCase$(Left$(trimmedTag, Len(TAG_PREFIX_TRANSLATION))) <> TAG_PREFIX_TRANSLATION Then
        Exit Function
    End If

    ExtractTranslationKeyFromTag = Trim$(Mid$(trimmedTag, Len(TAG_PREFIX_TRANSLATION) + 1))
End Function

Private Function SupportsCaptionLocalization(ByVal ControlInstance As Control) As Boolean
    On Error GoTo ErrorHandler

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    Select Case ControlInstance.ControlType
        Case acLabel, acCommandButton, acCheckBox, acOptionButton, acToggleButton
            SupportsCaptionLocalization = True
    End Select
    Exit Function

ErrorHandler:
    SupportsCaptionLocalization = False
    modErrorHandler.HandleError MODULE_NAME, "SupportsCaptionLocalization", Err
End Function

Private Sub ApplyCaptionToControl(ByVal ControlInstance As Control, ByVal CaptionValue As String)
    On Error GoTo ErrorHandler

    If ControlInstance Is Nothing Then
        Exit Sub
    End If

    Select Case ControlInstance.ControlType
        Case acLabel, acCommandButton, acCheckBox, acOptionButton, acToggleButton
            ControlInstance.Caption = CaptionValue
    End Select
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyCaptionToControl", Err
End Sub

Private Function GetControlFallbackCaption(ByVal ControlInstance As Control) As String
    On Error GoTo ErrorHandler

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    If SupportsCaptionLocalization(ControlInstance) Then
        GetControlFallbackCaption = NzString(ControlInstance.Caption)
    End If
    Exit Function

ErrorHandler:
    GetControlFallbackCaption = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "GetControlFallbackCaption", Err
End Function

Private Function LocalizeTabPages(ByVal TabControlInstance As Control) As Long
    On Error GoTo ErrorHandler

    Dim page As Access.Page
    Dim translationKey As String
    Dim fallbackCaption As String

    If TabControlInstance Is Nothing Then
        Exit Function
    End If

    If TabControlInstance.ControlType <> acTabCtl Then
        Exit Function
    End If

    For Each page In TabControlInstance.Pages
        translationKey = ExtractTranslationKeyFromTag(NzString(page.Tag))
        If LenB(translationKey) > 0 Then
            fallbackCaption = NzString(page.Caption)
            page.Caption = modTranslationService.T(translationKey, fallbackCaption)
            LocalizeTabPages = LocalizeTabPages + 1
        End If
    Next page
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LocalizeTabPages", Err
End Function

Private Function NzString(ByVal Value As Variant, Optional ByVal DefaultValue As String = "") As String
    If IsNull(Value) Or IsEmpty(Value) Then
        NzString = DefaultValue
    Else
        NzString = CStr(Value)
    End If
End Function
