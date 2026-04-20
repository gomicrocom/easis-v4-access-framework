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

Public Sub InitializeForm(ByVal FormInstance As Access.Form)
    On Error GoTo ErrorHandler

    Dim formName As String
    Dim requiredModule As String

    If FormInstance Is Nothing Then
        Exit Sub
    End If

    formName = GetFormName(FormInstance)

    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeForm", _
        "Initializing form '" & formName & "'."

    requiredModule = ExtractRequiredModuleFromTag(FormInstance.Tag)
    If LenB(requiredModule) > 0 Then
        modModuleManager.RequireModule requiredModule, True
    End If

    modFormLocalization.LocalizeForm FormInstance

    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeForm", _
        "Form '" & formName & "' initialized successfully."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "InitializeForm", Err
    Err.Raise Err.Number, Err.Source, Err.Description
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
