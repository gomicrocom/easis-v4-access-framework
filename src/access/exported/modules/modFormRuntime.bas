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
Private Const TAG_TOKEN_READONLY As String = "READONLY"
Private Const TAG_TOKEN_LOCKED As String = "LOCKED"
Private Const TAG_TOKEN_DISABLED As String = "DISABLED"
Private Const TAG_TOKEN_ROLE As String = "ROLE"
Private Const TAG_TOKEN_HIDDEN As String = "HIDDEN"
Private Const TAG_TOKEN_SETFOCUS As String = "SETFOCUS"
Private Const TAG_TOKEN_REQUIRED As String = "REQUIRED"

Public Sub InitializeForm(ByVal FormInstance As Access.Form)
    On Error GoTo ErrorHandler

    Dim FormName As String
    Dim RequiredModule As String
    Dim AllowedRoles As String
    Dim formTokens As Object
    Dim CurrentUserRoles As Collection

    If FormInstance Is Nothing Then
        Exit Sub
    End If

    FormName = GetFormName(FormInstance)
    Set formTokens = ParseTagTokens(FormInstance.Tag)
    Set CurrentUserRoles = modSessionContext.GetCurrentUserRoles()

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

    If formTokens.Exists(TAG_TOKEN_ROLE) Then
        AllowedRoles = CStr(formTokens(TAG_TOKEN_ROLE))
        If Not IsRoleAllowed(AllowedRoles, CurrentUserRoles) Then
            TryShowRoleDeniedMessage FormName
            modLoggingHandler.LogWarning MODULE_NAME & ".InitializeForm", _
                "Form '" & FormName & "' denied by ROLE policy '" & AllowedRoles & "'."
            DoCmd.Close acForm, FormName, acSaveNo
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

Public Function ValidateRequiredFields(ByVal FormInstance As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    Dim ctl As Control
    Dim controlTokens As Object
    Dim missingControls As Collection
    Dim missingFieldNames As Collection
    Dim firstMissingControl As Control

    ValidateRequiredFields = True

    If FormInstance Is Nothing Then
        Exit Function
    End If

    Set missingControls = New Collection
    Set missingFieldNames = New Collection

    For Each ctl In FormInstance.Controls
        Set controlTokens = ParseTagTokens(ctl.Tag)

        If IsControlRequired(controlTokens) Then
            If IsControlValueMissing(ctl) Then
                missingControls.Add ctl
                missingFieldNames.Add GetDisplayNameForRequiredControl(FormInstance, ctl)
            End If
        End If
    Next ctl

    If missingControls.Count = 0 Then
        Exit Function
    End If

    ValidateRequiredFields = False
    TryShowRequiredFieldsMessage missingFieldNames, GetFormName(FormInstance)

    Set firstMissingControl = missingControls.Item(1)
    Call TryFocusControl(firstMissingControl)

    modLoggingHandler.LogWarning MODULE_NAME & ".ValidateRequiredFields", _
        "Required-field validation failed on form '" & GetFormName(FormInstance) & "' for " & _
        CStr(missingControls.Count) & " control(s)."
    Exit Function

ErrorHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String

    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description

    modErrorHandler.HandleError MODULE_NAME, "ValidateRequiredFields", Err

    On Error GoTo 0
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Function

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
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String

    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description

    modErrorHandler.HandleError MODULE_NAME, "ApplyInitialFocusPolicy", Err
    
    On Error GoTo 0
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Sub

Private Function IsControlRequired(ByVal controlTokens As Object) As Boolean
    On Error GoTo ErrorHandler

    If controlTokens Is Nothing Then
        Exit Function
    End If

    IsControlRequired = controlTokens.Exists(TAG_TOKEN_REQUIRED)
    Exit Function

ErrorHandler:
    IsControlRequired = False
    modErrorHandler.HandleError MODULE_NAME, "IsControlRequired", Err
End Function

Private Sub ApplyControlPolicies(ByVal FormInstance As Access.Form)
    On Error GoTo ErrorHandler

    Dim ctl As Control
    Dim controlTokens As Object
    Dim CurrentUserRoles As Collection
    Dim lockedCount As Long
    Dim disabledCount As Long
    Dim roleHiddenCount As Long
    Dim hiddenCount As Long
    Dim requiredCount As Long

    If FormInstance Is Nothing Then
        Exit Sub
    End If

    Set CurrentUserRoles = modSessionContext.GetCurrentUserRoles()

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

        If controlTokens.Exists(TAG_TOKEN_ROLE) Then
            If TryApplyRolePolicy(ctl, controlTokens, CurrentUserRoles) Then
                roleHiddenCount = roleHiddenCount + 1
            End If
        End If

        If controlTokens.Exists(TAG_TOKEN_HIDDEN) Then
            If TryApplyHiddenPolicy(ctl) Then
                hiddenCount = hiddenCount + 1
            End If
        End If

        If controlTokens.Exists(TAG_TOKEN_REQUIRED) Then
            If TryApplyRequiredPolicy(FormInstance, ctl) Then
                requiredCount = requiredCount + 1
            End If
        End If
    Next ctl

    If lockedCount > 0 Or disabledCount > 0 Or roleHiddenCount > 0 Or hiddenCount > 0 Or requiredCount > 0 Then
        modLoggingHandler.LogInfo MODULE_NAME & ".ApplyControlPolicies", _
            "Control policies applied on form '" & GetFormName(FormInstance) & "': " & _
            "LOCKED=" & CStr(lockedCount) & ", DISABLED=" & CStr(disabledCount) & _
            ", ROLE_HIDDEN=" & CStr(roleHiddenCount) & ", HIDDEN=" & CStr(hiddenCount) & _
            ", REQUIRED=" & CStr(requiredCount) & "."
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

Private Function TryApplyRolePolicy(ByVal ControlInstance As Control, ByVal controlTokens As Object, ByVal CurrentUserRoles As Collection) As Boolean
    On Error GoTo SafeExit

    Dim AllowedRoles As String

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    If controlTokens Is Nothing Then
        Exit Function
    End If

    If Not controlTokens.Exists(TAG_TOKEN_ROLE) Then
        Exit Function
    End If

    AllowedRoles = CStr(controlTokens(TAG_TOKEN_ROLE))
    If IsRoleAllowed(AllowedRoles, CurrentUserRoles) Then
        Exit Function
    End If

    ControlInstance.Visible = False
    TryApplyRolePolicy = True

SafeExit:
End Function

Private Function IsRoleAllowed(ByVal AllowedRoles As String, ByVal CurrentUserRoles As Collection) As Boolean
    On Error GoTo ErrorHandler

    Dim roleParts() As String
    Dim roleItem As Variant
    Dim normalizedAllowedRole As String
    Dim currentRole As Variant

    If CurrentUserRoles Is Nothing Then
        Exit Function
    End If

    roleParts = Split(AllowedRoles, ",")
    For Each roleItem In roleParts
        normalizedAllowedRole = UCase$(Trim$(CStr(roleItem)))
        If LenB(normalizedAllowedRole) > 0 Then
            For Each currentRole In CurrentUserRoles
                If normalizedAllowedRole = UCase$(Trim$(CStr(currentRole))) Then
                    IsRoleAllowed = True
                    Exit Function
                End If
            Next currentRole
        End If
    Next roleItem
    Exit Function

ErrorHandler:
    IsRoleAllowed = False
    modErrorHandler.HandleError MODULE_NAME, "IsRoleAllowed", Err
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

Private Function TryApplyHiddenPolicy(ByVal ControlInstance As Control) As Boolean
    On Error GoTo SafeExit

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    If ControlInstance.Visible = False Then
        Exit Function
    End If

    ControlInstance.Visible = False
    TryApplyHiddenPolicy = True

SafeExit:
End Function

Private Function TryApplyRequiredPolicy(ByVal FormInstance As Access.Form, ByVal ControlInstance As Control) As Boolean
    On Error GoTo SafeExit

    Dim associatedLabel As Control
    Dim labelCaption As String

    If FormInstance Is Nothing Then
        Exit Function
    End If

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    Set associatedLabel = GetAssociatedLabel(FormInstance, ControlInstance)
    If associatedLabel Is Nothing Then
        Exit Function
    End If

    labelCaption = CStr(associatedLabel.Caption)
    If Right$(Trim$(labelCaption), 1) = "*" Then
        Exit Function
    End If

    associatedLabel.Caption = RTrim$(labelCaption) & " *"
    TryApplyRequiredPolicy = True

SafeExit:
End Function

Private Function IsControlValueMissing(ByVal ControlInstance As Control) As Boolean
    On Error GoTo SafeExit

    Dim controlValue As Variant

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    controlValue = ControlInstance.Value

    If IsNull(controlValue) Or IsEmpty(controlValue) Then
        IsControlValueMissing = True
        Exit Function
    End If

    If VarType(controlValue) = vbString Then
        IsControlValueMissing = (LenB(Trim$(CStr(controlValue))) = 0)
    End If

SafeExit:
End Function

Private Function GetAssociatedLabel(ByVal FormInstance As Access.Form, ByVal ControlInstance As Control) As Control
    On Error GoTo SafeExit

    Dim ctl As Control
    Dim childControl As Control
    Dim labelControlName As String

    If FormInstance Is Nothing Then Exit Function
    If ControlInstance Is Nothing Then Exit Function

    On Error Resume Next
    For Each childControl In ControlInstance.Controls
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If

        If childControl.ControlType = acLabel Then
            Set GetAssociatedLabel = childControl
            On Error GoTo SafeExit
            Exit Function
        End If
    Next childControl
    On Error GoTo SafeExit

    For Each ctl In FormInstance.Controls
        If ctl.ControlType = acLabel Then
            labelControlName = vbNullString

            On Error Resume Next
            labelControlName = CStr(ctl.ControlName)
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo SafeExit
                GoTo NextControl
            End If
            On Error GoTo SafeExit

            If StrComp(Trim$(labelControlName), Trim$(ControlInstance.Name), vbTextCompare) = 0 Then
                Set GetAssociatedLabel = ctl
                Exit Function
            End If
        End If

NextControl:
    Next ctl

SafeExit:
    If GetAssociatedLabel Is Nothing Then
        Set GetAssociatedLabel = Nothing
    End If
End Function

Private Function TryFocusControl(ByVal ControlInstance As Control) As Boolean
    On Error GoTo SafeExit

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    ControlInstance.SetFocus
    TryFocusControl = True

SafeExit:
End Function

Private Function GetDisplayNameForRequiredControl(ByVal FormInstance As Access.Form, ByVal ControlInstance As Control) As String
    On Error GoTo SafeExit

    Dim associatedLabel As Control
    Dim displayName As String

    If ControlInstance Is Nothing Then
        GetDisplayNameForRequiredControl = "<unknown>"
        Exit Function
    End If

    Set associatedLabel = GetAssociatedLabel(FormInstance, ControlInstance)
    If Not associatedLabel Is Nothing Then
        displayName = Trim$(CStr(associatedLabel.Caption))
        If Right$(displayName, 1) = "*" Then
            displayName = RTrim$(Left$(displayName, Len(displayName) - 1))
        End If

        If LenB(displayName) > 0 Then
            GetDisplayNameForRequiredControl = displayName
            Exit Function
        End If
    End If

    GetDisplayNameForRequiredControl = GetControlName(ControlInstance)
    Exit Function

SafeExit:
    GetDisplayNameForRequiredControl = GetControlName(ControlInstance)
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

Private Sub TryShowRoleDeniedMessage(ByVal FormName As String)
    On Error Resume Next

    Dim messageText As String
    Dim baseMessage As String

    baseMessage = "You do not have permission to open this form"
    messageText = modTranslationService.T("MSG_ROLE_NOT_ALLOWED", baseMessage)
    If LenB(Trim$(messageText)) = 0 Then
        messageText = baseMessage
    End If

    If LenB(Trim$(FormName)) > 0 And FormName <> "<unknown>" Then
        messageText = messageText & vbCrLf & "(" & FormName & ")"
    End If

    MsgBox messageText, vbExclamation, APP_NAME
End Sub

Private Sub TryShowRequiredFieldsMessage(ByVal MissingFieldNames As Collection, ByVal FormName As String)
    On Error Resume Next

    Dim messageText As String
    Dim baseMessage As String
    Dim fieldName As Variant
    Dim fieldList As String
    Dim fieldCount As Long

    baseMessage = "Please fill in all required fields."
    messageText = modTranslationService.T("MSG_REQUIRED_FIELDS_MISSING", baseMessage)
    If LenB(Trim$(messageText)) = 0 Then
        messageText = baseMessage
    End If

    If Not MissingFieldNames Is Nothing Then
        For Each fieldName In MissingFieldNames
            fieldCount = fieldCount + 1

            If fieldCount > 5 Then
                Exit For
            End If

            fieldList = fieldList & vbCrLf & "- " & CStr(fieldName)
        Next fieldName
    End If

    If LenB(fieldList) > 0 Then
        messageText = messageText & vbCrLf & fieldList
    End If

    If LenB(Trim$(FormName)) > 0 And FormName <> "<unknown>" Then
        messageText = messageText & vbCrLf & "(" & FormName & ")"
    End If

    MsgBox messageText, vbExclamation, APP_NAME
End Sub


