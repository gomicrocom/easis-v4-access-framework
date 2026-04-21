Attribute VB_Name = "modFormRuntime"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFormRuntime
' Purpose   : Standard runtime initialization entry point for Access forms.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modFormRuntime"
Private Const TAG_PREFIX_MODULE As String = "MOD:"
Private Const TAG_TOKEN_READONLY As String = "READONLY"
Private Const TAG_TOKEN_LOCKED As String = "LOCKED"

Public Sub InitializeForm(ByVal FormInstance As Access.Form)
    On Error GoTo ErrorHandler

    Dim FormName As String
    Dim RequiredModule As String

    If FormInstance Is Nothing Then
        Exit Sub
    End If

    FormName = GetFormName(FormInstance)

    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeForm", _
        "Initializing form '" & FormName & "'."

    RequiredModule = ExtractRequiredModuleFromTag(FormInstance.Tag)
    If LenB(RequiredModule) > 0 Then
        If Not modModuleManager.IsModuleActive(RequiredModule) Then
            TryShowMissingModuleMessage RequiredModule, FormName
            modLoggingHandler.LogWarning MODULE_NAME & ".InitializeForm", _
                "Form '" & FormName & "' requires inactive module '" & RequiredModule & "'."
            Exit Sub
        End If
    End If

    modFormLocalization.LocalizeForm FormInstance

    If HasReadOnlyTag(FormInstance.Tag) Then
        ApplyReadOnlyPolicy FormInstance
        modLoggingHandler.LogInfo MODULE_NAME & ".InitializeForm", _
            "Read-only policy applied to form '" & FormName & "'."
    End If

    ApplyControlPolicies FormInstance

    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeForm", _
        "Form '" & FormName & "' initialized successfully."
    Exit Sub

ErrorHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String

    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description

    modErrorHandler.HandleError MODULE_NAME, "InitializeForm", Err

    On Error GoTo 0
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Function GetFormName(ByVal FormInstance As Access.Form) As String
    On Error GoTo ErrorHandler

    If FormInstance Is Nothing Then
        GetFormName = "<unknown>"
    Else
        GetFormName = FormInstance.Name
    End If
    Exit Function

ErrorHandler:
    GetFormName = "<unknown>"
    modErrorHandler.HandleError MODULE_NAME, "GetFormName", Err
End Function

Private Function ExtractRequiredModuleFromTag(ByVal TagValue As String) As String
    On Error GoTo ErrorHandler

    Dim tokens() As String
    Dim token As Variant
    Dim trimmedToken As String

    trimmedToken = Trim$(TagValue)
    If LenB(trimmedToken) = 0 Then
        Exit Function
    End If

    tokens = Split(trimmedToken, ";")
    For Each token In tokens
        trimmedToken = Trim$(CStr(token))
        If LenB(trimmedToken) >= Len(TAG_PREFIX_MODULE) Then
            If UCase$(Left$(trimmedToken, Len(TAG_PREFIX_MODULE))) = TAG_PREFIX_MODULE Then
                ExtractRequiredModuleFromTag = UCase$(Trim$(Mid$(trimmedToken, Len(TAG_PREFIX_MODULE) + 1)))
                Exit Function
            End If
        End If
    Next token
    Exit Function

ErrorHandler:
    ExtractRequiredModuleFromTag = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "ExtractRequiredModuleFromTag", Err
End Function

Private Function HasReadOnlyTag(ByVal TagValue As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tokens() As String
    Dim token As Variant
    Dim trimmedToken As String

    trimmedToken = Trim$(TagValue)
    If LenB(trimmedToken) = 0 Then
        Exit Function
    End If

    tokens = Split(trimmedToken, ";")
    For Each token In tokens
        trimmedToken = UCase$(Trim$(CStr(token)))
        If trimmedToken = TAG_TOKEN_READONLY Then
            HasReadOnlyTag = True
            Exit Function
        End If
    Next token
    Exit Function

ErrorHandler:
    HasReadOnlyTag = False
    modErrorHandler.HandleError MODULE_NAME, "HasReadOnlyTag", Err
End Function

Private Sub ApplyControlPolicies(ByVal FormInstance As Access.Form)
    On Error GoTo ErrorHandler

    Dim ctl As Control
    Dim lockedCount As Long

    If FormInstance Is Nothing Then
        Exit Sub
    End If

    For Each ctl In FormInstance.Controls
        If ControlHasLockedTag(ctl.Tag) Then
            If TryApplyLockedPolicy(ctl) Then
                lockedCount = lockedCount + 1
            End If
        End If
    Next ctl

    If lockedCount > 0 Then
        modLoggingHandler.LogInfo MODULE_NAME & ".ApplyControlPolicies", _
            "Locked policy applied to " & CStr(lockedCount) & " control(s) on form '" & GetFormName(FormInstance) & "'."
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyControlPolicies", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function ControlHasLockedTag(ByVal TagValue As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tokens() As String
    Dim token As Variant
    Dim trimmedToken As String

    trimmedToken = Trim$(TagValue)
    If LenB(trimmedToken) = 0 Then
        Exit Function
    End If

    tokens = Split(trimmedToken, ";")
    For Each token In tokens
        trimmedToken = UCase$(Trim$(CStr(token)))
        If trimmedToken = TAG_TOKEN_LOCKED Then
            ControlHasLockedTag = True
            Exit Function
        End If
    Next token
    Exit Function

ErrorHandler:
    ControlHasLockedTag = False
    modErrorHandler.HandleError MODULE_NAME, "ControlHasLockedTag", Err
End Function

Private Function TryApplyLockedPolicy(ByVal ControlInstance As Control) As Boolean
    On Error GoTo SafeExit

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    ControlInstance.Locked = True
    TryApplyLockedPolicy = True

SafeExit:
End Function

Private Sub ApplyReadOnlyPolicy(ByVal FormInstance As Access.Form)
    On Error GoTo ErrorHandler

    If FormInstance Is Nothing Then
        Exit Sub
    End If

    FormInstance.AllowEdits = False
    FormInstance.AllowAdditions = False
    FormInstance.AllowDeletions = False
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyReadOnlyPolicy", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub TryShowMissingModuleMessage(ByVal RequiredModule As String, ByVal FormName As String)
    On Error Resume Next

    Dim messageText As String
    Dim baseMessage As String

    baseMessage = "The required module is not active"
    messageText = modTranslationService.T("MSG_MODULE_NOT_ACTIVE", baseMessage)
    If LenB(Trim$(messageText)) = 0 Then
        messageText = baseMessage
    End If

    messageText = messageText & ": " & Trim$(RequiredModule)

    If LenB(Trim$(FormName)) > 0 And FormName <> "<unknown>" Then
        messageText = messageText & vbCrLf & "(" & FormName & ")"
    End If

    MsgBox messageText, vbExclamation, APP_NAME
End Sub
