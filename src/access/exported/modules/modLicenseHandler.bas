Attribute VB_Name = "modLicenseHandler"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modLicenseHandler
' Purpose   : Loads and evaluates feature licensing for optional modules.
' Author    : Codex
' Project   : Easis Version 4
'===============================================================================

Private Const MODULE_NAME As String = "modLicenseHandler"

Public Function InitializeLicensing(Optional ByVal IniPath As String = vbNullString) As Boolean
    On Error GoTo ErrorHandler

    Dim featureList As String

    FeatureLicenses.RemoveAll

    featureList = modConfigIni.GetIniString(CONFIG_SECTION_LICENSE, "EnabledFeatures", vbNullString, IniPath)
    LoadFeatureList featureList

    If modConfigIni.GetIniBoolean(CONFIG_SECTION_LICENSE, "EnableCore", True, IniPath) Then
        GrantFeature "CORE"
    End If

    InitializeLicensing = True
    Exit Function

ErrorHandler:
    InitializeLicensing = False
    modErrorHandler.HandleError MODULE_NAME, "InitializeLicensing"
End Function

Public Function IsFeatureLicensed(ByVal FeatureCode As String) As Boolean
    On Error GoTo ErrorHandler

    Dim normalizedCode As String

    normalizedCode = NormalizeFeatureCode(FeatureCode)
    If LenB(normalizedCode) = 0 Then
        Exit Function
    End If

    IsFeatureLicensed = FeatureLicenses.Exists(normalizedCode)
    Exit Function

ErrorHandler:
    IsFeatureLicensed = False
    modErrorHandler.HandleError MODULE_NAME, "IsFeatureLicensed"
End Function

Public Sub EnsureFeatureLicensed(ByVal FeatureCode As String, Optional ByVal SourceContext As String = vbNullString)
    On Error GoTo ErrorHandler

    Dim normalizedCode As String
    Dim messageText As String

    normalizedCode = NormalizeFeatureCode(FeatureCode)
    If IsFeatureLicensed(normalizedCode) Then
        Exit Sub
    End If

    messageText = "Feature license missing: " & normalizedCode
    If LenB(SourceContext) > 0 Then
        messageText = messageText & " | Context: " & SourceContext
    End If

    Err.Raise vbObjectError + 2300, MODULE_NAME & ".EnsureFeatureLicensed", messageText

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureFeatureLicensed", True
End Sub

Public Sub GrantFeature(ByVal FeatureCode As String)
    Dim normalizedCode As String

    normalizedCode = NormalizeFeatureCode(FeatureCode)
    If LenB(normalizedCode) = 0 Then
        Exit Sub
    End If

    If Not FeatureLicenses.Exists(normalizedCode) Then
        FeatureLicenses.Add normalizedCode, True
    End If
End Sub

Private Sub LoadFeatureList(ByVal FeatureList As String)
    Dim tokens() As String
    Dim index As Long
    Dim token As String

    If LenB(Trim$(FeatureList)) = 0 Then
        Exit Sub
    End If

    tokens = Split(Replace(FeatureList, ";", ","), ",")
    For index = LBound(tokens) To UBound(tokens)
        token = Trim$(tokens(index))
        If LenB(token) > 0 Then
            GrantFeature token
        End If
    Next index
End Sub
