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
Private Const TAG_TOKEN_NUMERIC As String = "NUMERIC"
Private Const TAG_TOKEN_INTEGER As String = "INTEGER"
Private Const TAG_TOKEN_MIN As String = "MIN"
Private Const TAG_TOKEN_MAX As String = "MAX"
Private Const TAG_TOKEN_DATE As String = "DATE"
Private Const TAG_TOKEN_MINLEN As String = "MINLEN"
Private Const TAG_TOKEN_MAXLEN As String = "MAXLEN"
Private Const VALIDATION_HIGHLIGHT_BACKCOLOR As Long = 13434879
Private Const VALIDATION_DEFAULT_BACKCOLOR As Long = -2147483633
Private Const VALIDATION_COLORSTORE_PREFIX As String = "__VALORIG__:"

Private mValidationOriginalColors As Object

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
    
    ClearStoredValidationColorsForForm FormInstance
    
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
    Dim MissingFieldNames As Collection
    Dim firstMissingControl As Control

    ValidateRequiredFields = True

    If FormInstance Is Nothing Then
        Exit Function
    End If

    ClearValidationHighlights FormInstance
    
    Set missingControls = New Collection
    Set MissingFieldNames = New Collection

    For Each ctl In FormInstance.Controls
        If Not ShouldValidateControl(ctl) Then
            GoTo NextControl
        End If

        Set controlTokens = ParseTagTokens(ctl.Tag)

        If IsControlRequired(controlTokens) Then
            If IsControlValueMissing(ctl) Then
                missingControls.Add ctl
                MissingFieldNames.Add GetDisplayNameForRequiredControl(FormInstance, ctl)
                HighlightInvalidControl ctl
            End If
        End If
        
NextControl:
    Next ctl

    If missingControls.Count = 0 Then
        Exit Function
    End If

    ValidateRequiredFields = False
    TryShowRequiredFieldsMessage MissingFieldNames, GetFormName(FormInstance)

    Set firstMissingControl = missingControls.item(1)
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

Public Function ValidateFormPolicies(ByVal FormInstance As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    Dim ctl As Control
    Dim controlTokens As Object
    Dim missingRequiredControls As Collection
    Dim missingRequiredFieldNames As Collection
    Dim invalidFormatControls As Collection
    Dim invalidFormatFieldNames As Collection
    Dim firstInvalidControl As Control
    Dim isValueMissing As Boolean
    Dim errorMessage As String
        
    ValidateFormPolicies = True

    If FormInstance Is Nothing Then
        Exit Function
    End If
    
    ClearValidationHighlights FormInstance

    Set missingRequiredControls = New Collection
    Set missingRequiredFieldNames = New Collection
    Set invalidFormatControls = New Collection
    Set invalidFormatFieldNames = New Collection

    For Each ctl In FormInstance.Controls
        If Not ShouldValidateControl(ctl) Then
            GoTo NextControl
        End If

        Set controlTokens = ParseTagTokens(ctl.Tag)

        isValueMissing = IsControlValueMissing(ctl)

        If IsControlRequired(controlTokens) And isValueMissing Then
            missingRequiredControls.Add ctl
            missingRequiredFieldNames.Add GetDisplayNameForRequiredControl(FormInstance, ctl)
            HighlightInvalidControl ctl
        ElseIf IsControlValueInvalidForPolicies(ctl, controlTokens) Then
            invalidFormatControls.Add ctl
            HighlightInvalidControl ctl
            errorMessage = BuildControlValidationMessage(ctl, controlTokens)
            
            If LenB(errorMessage) > 0 Then
                invalidFormatFieldNames.Add _
                    GetDisplayNameForRequiredControl(FormInstance, ctl) & ": " & errorMessage
            End If
        End If
NextControl:
    Next ctl

    If missingRequiredControls.Count = 0 And invalidFormatControls.Count = 0 Then
        Exit Function
    End If

    ValidateFormPolicies = False
    TryShowValidationSummaryMessage missingRequiredFieldNames, invalidFormatFieldNames, GetFormName(FormInstance)

    If missingRequiredControls.Count > 0 Then
        Set firstInvalidControl = missingRequiredControls.item(1)
    Else
        Set firstInvalidControl = invalidFormatControls.item(1)
    End If

    Call TryFocusControl(firstInvalidControl)

    modLoggingHandler.LogWarning MODULE_NAME & ".ValidateFormPolicies", _
        "Form policy validation failed on form '" & GetFormName(FormInstance) & "' for " & _
        CStr(missingRequiredControls.Count + invalidFormatControls.Count) & " control(s)."
    Exit Function

ErrorHandler:
    Dim savedErrNumber As Long
    Dim savedErrSource As String
    Dim savedErrDescription As String

    savedErrNumber = Err.Number
    savedErrSource = Err.Source
    savedErrDescription = Err.Description

    modErrorHandler.HandleError MODULE_NAME, "ValidateFormPolicies", Err

    On Error GoTo 0
    Err.Raise savedErrNumber, savedErrSource, savedErrDescription
End Function

Private Function ShouldValidateControl(ByVal ControlInstance As Control) As Boolean
    On Error GoTo SafeExit

    If ControlInstance Is Nothing Then Exit Function

    On Error Resume Next
    Dim isVisible As Variant
    Dim isEnabled As Variant

    isVisible = ControlInstance.Visible
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If

    isEnabled = ControlInstance.Enabled
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo SafeExit

    If isVisible = False Then Exit Function
    If isEnabled = False Then Exit Function

    ShouldValidateControl = True

SafeExit:
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
    Dim TokenKey As String
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
            TokenKey = UCase$(Trim$(Left$(trimmedToken, separatorPosition - 1)))
            tokenValue = Trim$(Mid$(trimmedToken, separatorPosition + 1))

            If LenB(TokenKey) > 0 Then
                parsedTokens(TokenKey) = tokenValue
            End If
        Else
            TokenKey = UCase$(trimmedToken)
            If LenB(TokenKey) > 0 Then
                parsedTokens(TokenKey) = True
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

Private Function IsControlNumericTag(ByVal controlTokens As Object) As Boolean
    On Error GoTo ErrorHandler

    If controlTokens Is Nothing Then
        Exit Function
    End If

    IsControlNumericTag = controlTokens.Exists(TAG_TOKEN_NUMERIC)
    Exit Function

ErrorHandler:
    IsControlNumericTag = False
    modErrorHandler.HandleError MODULE_NAME, "IsControlNumericTag", Err
End Function

Private Function IsControlDateTag(ByVal controlTokens As Object) As Boolean
    On Error GoTo ErrorHandler

    If controlTokens Is Nothing Then
        Exit Function
    End If

    IsControlDateTag = controlTokens.Exists(TAG_TOKEN_DATE)
    Exit Function

ErrorHandler:
    IsControlDateTag = False
    modErrorHandler.HandleError MODULE_NAME, "IsControlDateTag", Err
End Function

Private Function IsControlIntegerTag(ByVal controlTokens As Object) As Boolean
    On Error GoTo ErrorHandler

    If controlTokens Is Nothing Then
        Exit Function
    End If

    IsControlIntegerTag = controlTokens.Exists(TAG_TOKEN_INTEGER)
    Exit Function

ErrorHandler:
    IsControlIntegerTag = False
    modErrorHandler.HandleError MODULE_NAME, "IsControlIntegerTag", Err
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

Private Function IsControlValueInvalidForPolicies(ByVal ControlInstance As Control, ByVal controlTokens As Object) As Boolean
    On Error GoTo SafeExit

    Dim minValue As Double
    Dim maxValue As Double
    Dim minLenValue As Long
    Dim maxLenValue As Long
    
    If ControlInstance Is Nothing Then
        Exit Function
    End If

    If controlTokens Is Nothing Then
        Exit Function
    End If

    If IsControlRequired(controlTokens) Then
        If IsControlValueMissing(ControlInstance) Then
            IsControlValueInvalidForPolicies = True
            Exit Function
        End If
    End If

    If IsControlNumericTag(controlTokens) Then
        If Not IsControlValueNumericValid(ControlInstance) Then
            IsControlValueInvalidForPolicies = True
            Exit Function
        End If
    End If

    If IsControlIntegerTag(controlTokens) Then
    
        If Not IsControlValueNumericValid(ControlInstance) Then
            IsControlValueInvalidForPolicies = True
            Exit Function
        End If
    
        If Not IsControlValueIntegerValid(ControlInstance) Then
            IsControlValueInvalidForPolicies = True
            Exit Function
        End If
    
    End If

    If IsControlNumericTag(controlTokens) Then
    
        If GetMinValue(controlTokens, minValue) Then
            If Not IsControlValueMinValid(ControlInstance, minValue) Then
                IsControlValueInvalidForPolicies = True
                Exit Function
            End If
        End If
    
        If GetMaxValue(controlTokens, maxValue) Then
            If Not IsControlValueMaxValid(ControlInstance, maxValue) Then
                IsControlValueInvalidForPolicies = True
                Exit Function
            End If
        End If
    
    End If

    If GetMinLenValue(controlTokens, minLenValue) Then
        If Not IsControlValueMinLenValid(ControlInstance, minLenValue) Then
            IsControlValueInvalidForPolicies = True
            Exit Function
        End If
    End If
    
    If GetMaxLenValue(controlTokens, maxLenValue) Then
        If Not IsControlValueMaxLenValid(ControlInstance, maxLenValue) Then
            IsControlValueInvalidForPolicies = True
            Exit Function
        End If
    End If
    
    If IsControlDateTag(controlTokens) Then
        If Not IsControlValueDateValid(ControlInstance) Then
            IsControlValueInvalidForPolicies = True
        End If
    End If

SafeExit:
End Function

Private Function GetMinValue(ByVal controlTokens As Object, ByRef minValue As Double) As Boolean
    On Error GoTo SafeExit

    Dim tokenValue As String

    If controlTokens Is Nothing Then
        Exit Function
    End If

    If Not controlTokens.Exists(TAG_TOKEN_MIN) Then
        Exit Function
    End If

    tokenValue = Trim$(CStr(controlTokens(TAG_TOKEN_MIN)))
    If LenB(tokenValue) = 0 Then
        Exit Function
    End If

    If Not IsNumeric(tokenValue) Then
        Exit Function
    End If

    minValue = CDbl(tokenValue)
    GetMinValue = True

SafeExit:
End Function

Private Function GetMaxValue(ByVal controlTokens As Object, ByRef maxValue As Double) As Boolean
    On Error GoTo SafeExit

    Dim tokenValue As String

    If controlTokens Is Nothing Then
        Exit Function
    End If

    If Not controlTokens.Exists(TAG_TOKEN_MAX) Then
        Exit Function
    End If

    tokenValue = Trim$(CStr(controlTokens(TAG_TOKEN_MAX)))
    If LenB(tokenValue) = 0 Then
        Exit Function
    End If

    If Not IsNumeric(tokenValue) Then
        Exit Function
    End If

    maxValue = CDbl(tokenValue)
    GetMaxValue = True

SafeExit:
End Function

Private Function IsControlValueNumericValid(ByVal ControlInstance As Control) As Boolean
    On Error GoTo SafeExit

    Dim controlValue As Variant

    IsControlValueNumericValid = True

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    controlValue = ControlInstance.Value

    If IsNull(controlValue) Or IsEmpty(controlValue) Then
        Exit Function
    End If

    If VarType(controlValue) = vbString Then
        If LenB(Trim$(CStr(controlValue))) = 0 Then
            Exit Function
        End If
    End If

    IsControlValueNumericValid = IsNumeric(controlValue)

SafeExit:
End Function

Private Function IsControlValueIntegerValid(ByVal ControlInstance As Control) As Boolean
    On Error GoTo SafeExit

    Dim controlValue As Variant
    Dim numericValue As Double

    IsControlValueIntegerValid = True

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    controlValue = ControlInstance.Value

    If IsNull(controlValue) Or IsEmpty(controlValue) Then
        Exit Function
    End If

    If VarType(controlValue) = vbString Then
        If LenB(Trim$(CStr(controlValue))) = 0 Then
            Exit Function
        End If
    End If

    If Not IsNumeric(controlValue) Then
        IsControlValueIntegerValid = False
        Exit Function
    End If

    numericValue = CDbl(controlValue)
    IsControlValueIntegerValid = (numericValue = Fix(numericValue))

SafeExit:
End Function

Private Function IsControlValueMinValid(ByVal ControlInstance As Control, ByVal minValue As Double) As Boolean
    On Error GoTo SafeExit

    Dim controlValue As Variant

    IsControlValueMinValid = True

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    controlValue = ControlInstance.Value

    If IsNull(controlValue) Or IsEmpty(controlValue) Then
        Exit Function
    End If

    If VarType(controlValue) = vbString Then
        If LenB(Trim$(CStr(controlValue))) = 0 Then
            Exit Function
        End If
    End If

    If Not IsNumeric(controlValue) Then
        IsControlValueMinValid = False
        Exit Function
    End If

    IsControlValueMinValid = (CDbl(controlValue) >= minValue)

SafeExit:
End Function

Private Function IsControlValueMaxValid(ByVal ControlInstance As Control, ByVal maxValue As Double) As Boolean
    On Error GoTo SafeExit

    Dim controlValue As Variant

    IsControlValueMaxValid = True

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    controlValue = ControlInstance.Value

    If IsNull(controlValue) Or IsEmpty(controlValue) Then
        Exit Function
    End If

    If VarType(controlValue) = vbString Then
        If LenB(Trim$(CStr(controlValue))) = 0 Then
            Exit Function
        End If
    End If

    If Not IsNumeric(controlValue) Then
        IsControlValueMaxValid = False
        Exit Function
    End If

    IsControlValueMaxValid = (CDbl(controlValue) <= maxValue)

SafeExit:
End Function

Private Function IsControlValueDateValid(ByVal ControlInstance As Control) As Boolean
    On Error GoTo SafeExit

    Dim controlValue As Variant

    IsControlValueDateValid = True

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    controlValue = ControlInstance.Value

    If IsNull(controlValue) Or IsEmpty(controlValue) Then
        Exit Function
    End If

    If VarType(controlValue) = vbString Then
        If LenB(Trim$(CStr(controlValue))) = 0 Then
            Exit Function
        End If
    End If

    IsControlValueDateValid = IsDate(controlValue)

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
    Dim DisplayName As String

    If ControlInstance Is Nothing Then
        GetDisplayNameForRequiredControl = "<unknown>"
        Exit Function
    End If

    Set associatedLabel = GetAssociatedLabel(FormInstance, ControlInstance)
    If Not associatedLabel Is Nothing Then
        DisplayName = Trim$(CStr(associatedLabel.Caption))
        If Right$(DisplayName, 1) = "*" Then
            DisplayName = RTrim$(Left$(DisplayName, Len(DisplayName) - 1))
        End If

        If LenB(DisplayName) > 0 Then
            GetDisplayNameForRequiredControl = DisplayName
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

Private Sub TryShowInvalidFieldsMessage(ByVal InvalidFieldNames As Collection, ByVal FormName As String)
    On Error Resume Next

    Dim messageText As String
    Dim baseMessage As String
    Dim fieldName As Variant
    Dim fieldList As String
    Dim fieldCount As Long

    baseMessage = "Please correct invalid field values."
    messageText = modTranslationService.T("MSG_INVALID_FIELD_VALUES", baseMessage)
    If LenB(Trim$(messageText)) = 0 Then
        messageText = baseMessage
    End If

    If Not InvalidFieldNames Is Nothing Then
        For Each fieldName In InvalidFieldNames
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

Private Sub TryShowValidationSummaryMessage(ByVal MissingFieldNames As Collection, ByVal InvalidFieldNames As Collection, ByVal FormName As String)
    On Error Resume Next

    Dim messageText As String
    Dim requiredMessage As String
    Dim invalidMessage As String
    Dim requiredFieldList As String
    Dim invalidFieldList As String

    requiredMessage = modTranslationService.T("MSG_REQUIRED_FIELDS_MISSING", "Please fill in all required fields.")
    If LenB(Trim$(requiredMessage)) = 0 Then
        requiredMessage = "Please fill in all required fields."
    End If

    invalidMessage = modTranslationService.T("MSG_INVALID_FIELD_VALUES", "Please correct invalid field values.")
    If LenB(Trim$(invalidMessage)) = 0 Then
        invalidMessage = "Please correct invalid field values."
    End If

    requiredFieldList = BuildValidationFieldList(MissingFieldNames)
    invalidFieldList = BuildValidationFieldList(InvalidFieldNames)

    If LenB(requiredFieldList) > 0 Then
        messageText = requiredMessage & vbCrLf & requiredFieldList
    End If

    If LenB(invalidFieldList) > 0 Then
        If LenB(messageText) > 0 Then
            messageText = messageText & vbCrLf & vbCrLf
        End If
        messageText = messageText & invalidMessage & vbCrLf & invalidFieldList
    End If

    If LenB(Trim$(FormName)) > 0 And FormName <> "<unknown>" Then
        messageText = messageText & vbCrLf & vbCrLf & "(" & FormName & ")"
    End If

    If LenB(Trim$(messageText)) > 0 Then
        MsgBox messageText, vbExclamation, APP_NAME
    End If
End Sub

Private Function BuildValidationFieldList(ByVal FieldNames As Collection) As String
    On Error GoTo SafeExit

    Dim fieldName As Variant
    Dim fieldCount As Long

    If FieldNames Is Nothing Then
        Exit Function
    End If

    For Each fieldName In FieldNames
        fieldCount = fieldCount + 1

        If fieldCount > 5 Then
            Exit For
        End If

        BuildValidationFieldList = BuildValidationFieldList & "- " & CStr(fieldName) & vbCrLf
    Next fieldName

    If LenB(BuildValidationFieldList) > 0 Then
        BuildValidationFieldList = Left$(BuildValidationFieldList, Len(BuildValidationFieldList) - Len(vbCrLf))
    End If

SafeExit:
End Function

Private Function GetMinLenValue(ByVal controlTokens As Object, ByRef minLenValue As Long) As Boolean
    On Error GoTo SafeExit

    Dim tokenValue As String

    If controlTokens Is Nothing Then Exit Function
    If Not controlTokens.Exists(TAG_TOKEN_MINLEN) Then Exit Function

    tokenValue = Trim$(CStr(controlTokens(TAG_TOKEN_MINLEN)))
    If LenB(tokenValue) = 0 Then Exit Function
    If Not IsNumeric(tokenValue) Then Exit Function

    minLenValue = CLng(tokenValue)
    If minLenValue < 0 Then Exit Function

    GetMinLenValue = True

SafeExit:
End Function

Private Function GetMaxLenValue(ByVal controlTokens As Object, ByRef maxLenValue As Long) As Boolean
    On Error GoTo SafeExit

    Dim tokenValue As String

    If controlTokens Is Nothing Then Exit Function
    If Not controlTokens.Exists(TAG_TOKEN_MAXLEN) Then Exit Function

    tokenValue = Trim$(CStr(controlTokens(TAG_TOKEN_MAXLEN)))
    If LenB(tokenValue) = 0 Then Exit Function
    If Not IsNumeric(tokenValue) Then Exit Function

    maxLenValue = CLng(tokenValue)
    If maxLenValue < 0 Then Exit Function

    GetMaxLenValue = True

SafeExit:
End Function

Private Function IsControlValueMinLenValid(ByVal ControlInstance As Control, ByVal minLenValue As Long) As Boolean
    On Error GoTo SafeExit

    Dim controlValue As Variant
    Dim textValue As String

    IsControlValueMinLenValid = True

    If ControlInstance Is Nothing Then Exit Function

    controlValue = ControlInstance.Value

    If IsNull(controlValue) Or IsEmpty(controlValue) Then Exit Function
    If VarType(controlValue) <> vbString Then Exit Function

    textValue = Trim$(CStr(controlValue))
    If LenB(textValue) = 0 Then Exit Function

    IsControlValueMinLenValid = (Len(textValue) >= minLenValue)

SafeExit:
End Function

Private Function IsControlValueMaxLenValid(ByVal ControlInstance As Control, ByVal maxLenValue As Long) As Boolean
    On Error GoTo SafeExit

    Dim controlValue As Variant
    Dim textValue As String

    IsControlValueMaxLenValid = True

    If ControlInstance Is Nothing Then Exit Function

    controlValue = ControlInstance.Value

    If IsNull(controlValue) Or IsEmpty(controlValue) Then Exit Function
    If VarType(controlValue) <> vbString Then Exit Function

    textValue = Trim$(CStr(controlValue))
    If LenB(textValue) = 0 Then Exit Function

    IsControlValueMaxLenValid = (Len(textValue) <= maxLenValue)

SafeExit:
End Function

Private Function BuildControlValidationMessage( _
    ByVal ControlInstance As Control, _
    ByVal controlTokens As Object) As String

    On Error GoTo SafeExit

    Dim minValue As Double
    Dim maxValue As Double
    Dim minLenValue As Long
    Dim maxLenValue As Long

    ' REQUIRED
    If IsControlRequired(controlTokens) Then
        If IsControlValueMissing(ControlInstance) Then
            BuildControlValidationMessage = modTranslationService.T("ERR_REQUIRED", "is required")
            Exit Function
        End If
    End If

    ' NUMERIC
    If IsControlNumericTag(controlTokens) Then
        If Not IsControlValueNumericValid(ControlInstance) Then
            BuildControlValidationMessage = modTranslationService.T("ERR_NUMERIC", "must be a number")
            Exit Function
        End If
    End If

    ' INTEGER
    If IsControlIntegerTag(controlTokens) Then
        If Not IsControlValueIntegerValid(ControlInstance) Then
            BuildControlValidationMessage = modTranslationService.T("ERR_INTEGER", "must be an integer")
            Exit Function
        End If
    End If

    ' MIN / MAX
    If IsControlNumericTag(controlTokens) Then
        If GetMinValue(controlTokens, minValue) Then
            If Not IsControlValueMinValid(ControlInstance, minValue) Then
                BuildControlValidationMessage = modTranslationService.TEx("ERR_MIN", "must be >= {0}", minValue)
                Exit Function
            End If
        End If

        If GetMaxValue(controlTokens, maxValue) Then
            If Not IsControlValueMaxValid(ControlInstance, maxValue) Then
                BuildControlValidationMessage = modTranslationService.TEx("ERR_MAX", "must be <= {0}", maxValue)
                Exit Function
            End If
        End If
    End If

    ' MINLEN / MAXLEN
    If GetMinLenValue(controlTokens, minLenValue) Then
        If Not IsControlValueMinLenValid(ControlInstance, minLenValue) Then
            BuildControlValidationMessage = modTranslationService.TEx("ERR_MINLEN", "minimum length is {0}", minLenValue)
            Exit Function
        End If
    End If

    If GetMaxLenValue(controlTokens, maxLenValue) Then
        If Not IsControlValueMaxLenValid(ControlInstance, maxLenValue) Then
            BuildControlValidationMessage = modTranslationService.TEx("ERR_MAXLEN", "maximum length is {0}", maxLenValue)
            Exit Function
        End If
    End If

    ' DATE
    If IsControlDateTag(controlTokens) Then
        If Not IsControlValueDateValid(ControlInstance) Then
            BuildControlValidationMessage = modTranslationService.T("ERR_DATE", "must be a valid date")
        End If
    End If

SafeExit:
End Function

Private Sub HighlightInvalidControl(ByVal ControlInstance As Control)
    On Error GoTo SafeExit

    Dim colorKey As String
    Dim originalColor As Variant

    If ControlInstance Is Nothing Then
        Exit Sub
    End If

    If Not CanHighlightControl(ControlInstance) Then
        Exit Sub
    End If

    EnsureValidationColorStore

    colorKey = GetValidationColorKey(ControlInstance)
    If LenB(colorKey) = 0 Then
        Exit Sub
    End If

    If Not mValidationOriginalColors.Exists(colorKey) Then
        originalColor = ControlInstance.BackColor
        mValidationOriginalColors.Add colorKey, originalColor
    End If

    ControlInstance.BackColor = VALIDATION_HIGHLIGHT_BACKCOLOR

SafeExit:
End Sub

Private Sub ClearValidationHighlights(ByVal FormInstance As Access.Form)
    On Error GoTo SafeExit

    Dim ctl As Control
    Dim colorKey As String

    If FormInstance Is Nothing Then
        Exit Sub
    End If

    EnsureValidationColorStore

    For Each ctl In FormInstance.Controls
        If CanHighlightControl(ctl) Then
            colorKey = GetValidationColorKey(ctl)

            If LenB(colorKey) > 0 Then
                If mValidationOriginalColors.Exists(colorKey) Then
                    On Error Resume Next
                    ctl.BackColor = mValidationOriginalColors(colorKey)
                    Err.Clear
                    On Error GoTo SafeExit

                    mValidationOriginalColors.Remove colorKey
                End If
            End If
        End If
    Next ctl

SafeExit:
End Sub
Private Function CanHighlightControl(ByVal ControlInstance As Control) As Boolean
    On Error GoTo SafeExit

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    Select Case ControlInstance.ControlType
        Case acTextBox, acComboBox, acListBox, acCheckBox, acOptionGroup
            CanHighlightControl = True
    End Select

SafeExit:
End Function
Private Sub EnsureValidationColorStore()
    If mValidationOriginalColors Is Nothing Then
        Set mValidationOriginalColors = CreateObject("Scripting.Dictionary")
        mValidationOriginalColors.CompareMode = vbTextCompare
    End If
End Sub

Private Function GetValidationColorKey(ByVal ControlInstance As Control) As String
    On Error GoTo SafeExit

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    GetValidationColorKey = GetParentFormName(ControlInstance) & "|" & ControlInstance.Name
    Exit Function

SafeExit:
    GetValidationColorKey = vbNullString
End Function

Private Function GetParentFormName(ByVal ControlInstance As Control) As String
    On Error GoTo SafeExit

    If ControlInstance Is Nothing Then
        Exit Function
    End If

    GetParentFormName = ControlInstance.Parent.Name
    Exit Function

SafeExit:
    GetParentFormName = "<unknown>"
End Function
Private Sub ClearStoredValidationColorsForForm(ByVal FormInstance As Access.Form)
    On Error GoTo SafeExit

    Dim dictKey As Variant
    Dim keysToRemove As Collection
    Dim keyItem As Variant
    Dim formPrefix As String

    If FormInstance Is Nothing Then
        Exit Sub
    End If

    EnsureValidationColorStore
    Set keysToRemove = New Collection

    formPrefix = FormInstance.Name & "|"

    For Each dictKey In mValidationOriginalColors.Keys
        If Left$(CStr(dictKey), Len(formPrefix)) = formPrefix Then
            keysToRemove.Add CStr(dictKey)
        End If
    Next dictKey

    For Each keyItem In keysToRemove
        mValidationOriginalColors.Remove CStr(keyItem)
    Next keyItem

SafeExit:
End Sub
