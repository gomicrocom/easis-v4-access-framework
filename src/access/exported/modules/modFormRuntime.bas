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
Private Const TAG_TOKEN_DISABLED As String = "DISABLED"
Private Const TAG_TOKEN_HIDDEN As String = "HIDDEN"
Private Const TAG_TOKEN_SETFOCUS As String = "SETFOCUS"

Public Sub InitializeForm(ByVal FormInstance As Access.Form)
    On Error GoTo ErrorHandler

    Dim FormName As String
    Dim RequiredModule As String
    Dim formTokens As Object

    If FormInstance Is Nothing Then
        Exit Sub
    End If

    FormName = GetFormName(FormInstance)
    Set formTokens = ParseTagTokens(FormInstance.Tag)

    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeForm", _
        "Initializing form '" & FormName & "'."

    If formTokens.Exists("MOD") Then
        RequiredModule = UCase$(Trim$(CStr(formTokens("MOD"))))
    End If

    If LenB(RequiredModule) > 0 Then
        If Not modModuleManager.IsModuleActive(RequiredModule) Then
            TryShowMissingModuleMessage RequiredModule, FormName
            modLoggingHandler.LogWarning MODULE_NAME & ".InitializeForm", _
                "Form '" & FormName & "' requires inactive module '" & RequiredModule & "'."
            Exit Sub
        End If
    End If

    modFormLocalization.LocalizeForm FormInstance

    If formTokens.Exists(TAG_TOKEN_READONLY) Then
        ApplyReadOnlyPolicy FormInstance
        modLoggingHandler.LogInfo MODULE_NAME & ".InitializeForm", _
            "Read-only policy applied to form '" & FormName & "'."
    End If

    ApplyInitialFocusPolicy FormInstance
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

Private Function ParseTagTokens(ByVal TagValue As String) As Object
    On Error GoTo ErrorHandler

    Dim parsedTokens As Object
    Dim tagParts() As String
    Dim token As Variant
    Dim trimmedToken As String
    Dim separatorPosition As Long
    Dim tokenKey As String
    Dim tokenValue As String

    Set parsedTokens = CreateObject("Scripting.Dictionary")
    parsedTokens.CompareMode = vbTextCompare

    trimmedToken = Trim$(TagValue)
    If LenB(trimmedToken) = 0 Then
        Set ParseTagTokens = parsedTokens
        Exit Function
    End If

    tagParts = Split(trimmedToken, ";")
    For Each token In tagParts
        trimmedToken = Trim$(CStr(token))
        If LenB(trimmedToken) = 0 Then
            GoTo NextToken
        End If

        separatorPosition = InStr(1, trimmedToken, ":", vbTextCompare)
        If separatorPosition > 0 Then
            tokenKey = UCase$(Trim$(Left$(trimmedToken, separatorPosition - 1)))
            tokenValue = Trim$(Mid$(trimmedToken, separatorPosition + 1))

            If LenB(tokenKey) > 0 Then
                parsedTokens(tokenKey) = tokenValue
            End If
        Else
            tokenKey = UCase$(trimmedToken)
            If LenB(tokenKey) > 0 Then
                parsedTokens(tokenKey) = True
            End If
        End If

NextToken:
    Next token

    Set ParseTagTokens = parsedTokens
    Exit Function

ErrorHandler:
    Set ParseTagTokens = CreateObject("Scripting.Dictionary")
    ParseTagTokens.CompareMode = vbTextCompare
    modErrorHandler.HandleError MODULE_NAME, "ParseTagTokens", Err
End Function

Private Sub ApplyInitialFocusPolicy(ByVal FormInstance As Access.Form)
    On Error GoTo ErrorHandler

    Dim ctl As Control
    Dim controlTokens As Object

    If FormInstance Is Nothing Then
        Exit Sub
    End If

    For Each ctl In FormInstance.Controls
        Set controlTokens = ParseTagTokens(ctl.Tag)
        If controlTokens.Exists(TAG_TOKEN_SETFOCUS) Then
            If TrySetInitialFocus(ctl) Then
                modLoggingHandler.LogInfo MODULE_NAME & ".ApplyInitialFocusPolicy", _
                    "Initial focus set to control '" & GetControlName(ctl) & "' on form '" & GetFormName(FormInstance) & "'."
            Else
                modLoggingHandler.LogWarning MODULE_NAME & ".ApplyInitialFocusPolicy", _
                    "Initial focus could not be set to control '" & GetControlName(ctl) & "' on form '" & GetFormName(FormInstance) & "'."
            End If
            Exit Sub
        End If
    Next ctl
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyInitialFocusPolicy", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ApplyControlPolicies(ByVal FormInstance As Access.Form)
    On Error GoTo ErrorHandler

    Dim ctl As Control
    Dim controlTokens As Object
    Dim lockedCount As Long
    Dim disabledCount As Long
    Dim hiddenCount As Long

    If FormInstance Is Nothing Then
        Exit Sub
    End If

    For Each ctl In FormInstance.Controls
        Set controlTokens = ParseTagTokens(ctl.Tag)

        If controlTokens.Exists(TAG_TOKEN_LOCKED) Then
            If TryApplyLockedPolicy(ctl) Then
                lockedCount = lockedCount + 1
            End If
        End If

        If controlTokens.Exists(TAG_TOKEN_DISABLED) Then
            If TryApplyDisabledPolicy(ctl) Then
                disabledCount = disabledCount + 1
            End If
        End If

        If controlTokens.Exists(TAG_TOKEN_HIDDEN) Then
            If TryApplyHiddenPolicy(FormInstance, ctl) Then
                hiddenCount = hiddenCount + 1
            End If
        End If
    Next ctl

    If lockedCount > 0 Or disabledCount > 0 Or hiddenCount > 0 Then
        modLoggingHandler.LogInfo MODULE_NAME & ".ApplyControlPolicies", _
            "Control policies applied on form '" & GetFormName(FormInstance) & "': " & _
            "LOCKED=" & CStr(lockedCount) & ", DISABLED=" & CStr(disabledCount) & _
            ", HIDDEN=" & CStr(hiddenCount) & "."
    End If
    Exit Sub

ErrorHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String

    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description

    modErrorHandler.HandleError MODULE_NAME, "ApplyControlPolicies", Err

    On Error GoTo 0
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Function TryApplyLockedPolicy(ByVal ControlInstance As Control) As Boolean
    On Error GoTo SafeExit

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    ControlInstance.Locked = True
    TryApplyLockedPolicy = True

SafeExit:
End Function

Private Function TryApplyDisabledPolicy(ByVal ControlInstance As Control) As Boolean
    On Error GoTo SafeExit

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    ControlInstance.Enabled = False
    TryApplyDisabledPolicy = True

SafeExit:
End Function

Private Function TrySetInitialFocus(ByVal ControlInstance As Control) As Boolean
    On Error GoTo SafeExit

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    ControlInstance.SetFocus
    TrySetInitialFocus = True

SafeExit:
End Function

Private Function TryApplyHiddenPolicy(ByVal FormInstance As Access.Form, ByVal ControlInstance As Control) As Boolean
    On Error GoTo SafeExit

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    ControlInstance.Visible = False
    TryApplyHiddenPolicy = True

SafeExit:
End Function

Private Function GetControlName(ByVal ControlInstance As Control) As String
    On Error GoTo SafeExit

    If ControlInstance Is Nothing Then
        GetControlName = "<unknown>"
        Exit Function
    End If

    GetControlName = ControlInstance.Name
    Exit Function

SafeExit:
    GetControlName = "<unknown>"
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
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String

    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description

    modErrorHandler.HandleError MODULE_NAME, "ApplyReadOnlyPolicy", Err

    On Error GoTo 0
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
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
