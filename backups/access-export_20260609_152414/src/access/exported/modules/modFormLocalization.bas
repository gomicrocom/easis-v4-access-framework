Attribute VB_Name = "modFormLocalization"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFormLocalization
' Purpose   : Applies translation keys from Tag metadata to Access forms and controls.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modFormLocalization"

Public Sub LocalizeForm(ByVal formInstance As Access.Form)
    On Error GoTo ErrorHandler

    Dim translationKey As String
    Dim localizedCount As Long
    Dim ctl As Control

    If formInstance Is Nothing Then
        Exit Sub
    End If

    translationKey = modFwTranslationRuntime.GetTranslationKeyFromTag(formInstance.Tag)
    If LenB(translationKey) > 0 Then
        SetFormCaption formInstance, translationKey, NzString(formInstance.Caption)
        localizedCount = localizedCount + 1
    End If

    For Each ctl In formInstance.Controls
        translationKey = modFwTranslationRuntime.GetTranslationKeyFromTag(ctl.Tag)
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
            "Localized " & CStr(localizedCount) & " element(s) on form '" & formInstance.Name & "'."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LocalizeForm", Err
End Sub

Public Sub SetFormCaption(ByVal formInstance As Access.Form, ByVal translationKey As String, Optional ByVal Fallback As String = "")
    On Error GoTo ErrorHandler

    If formInstance Is Nothing Then
        Exit Sub
    End If

    If LenB(Trim$(translationKey)) = 0 Then
        Exit Sub
    End If

    formInstance.Caption = modTranslationService.T(translationKey, Fallback)
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SetFormCaption", Err
End Sub

Public Sub LocalizeControl(ByVal ControlInstance As Control, ByVal translationKey As String, Optional ByVal Fallback As String = "")
    On Error GoTo ErrorHandler

    If ControlInstance Is Nothing Then
        Exit Sub
    End If

    If LenB(Trim$(translationKey)) = 0 Then
        Exit Sub
    End If

    If Not SupportsCaptionLocalization(ControlInstance) Then
        Exit Sub
    End If

    ApplyCaptionToControl ControlInstance, modTranslationService.T(translationKey, Fallback)
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LocalizeControl", Err
End Sub

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

    Dim Page As Access.Page
    Dim translationKey As String
    Dim fallbackCaption As String

    If TabControlInstance Is Nothing Then
        Exit Function
    End If

    If TabControlInstance.ControlType <> acTabCtl Then
        Exit Function
    End If

    For Each Page In TabControlInstance.Pages
        translationKey = modFwTranslationRuntime.GetTranslationKeyFromTag(NzString(Page.Tag))
        If LenB(translationKey) > 0 Then
            fallbackCaption = NzString(Page.Caption)
            Page.Caption = modTranslationService.T(translationKey, fallbackCaption)
            LocalizeTabPages = LocalizeTabPages + 1
        End If
    Next Page
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
