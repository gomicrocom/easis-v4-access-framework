Option Compare Database
Option Explicit

'===============================================================================
' Module    : modLicenseHandler
' Purpose   : Feature-based licensing helpers with placeholder logic.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modLicenseHandler"

Public Const FEATURE_CORE As String = "CORE"
Public Const FEATURE_CAMT054 As String = "CAMT054"
Public Const FEATURE_PROPERTY_MGMT As String = "PROPERTY_MGMT"
Public Const FEATURE_WINE_MGMT As String = "WINE_MGMT"
Public FeatureLicenses As Object


Public Function InitializeLicensing(Optional ByVal IniPath As String = vbNullString) As Boolean
    On Error GoTo ErrorHandler

    Dim FeatureList As String

    EnsureLicenseStore

    FeatureLicenses.RemoveAll

    ' Core standardmäßig aktivieren
    GrantFeature FEATURE_CORE

    FeatureList = modConfigIni.GetIniString(CONFIG_SECTION_LICENSE, "EnabledFeatures", vbNullString, IniPath)
    LoadFeatureList FeatureList

    If modConfigIni.GetIniBoolean(CONFIG_SECTION_LICENSE, FEATURE_CAMT054, False, IniPath) Then
        GrantFeature FEATURE_CAMT054
    End If

    If modConfigIni.GetIniBoolean(CONFIG_SECTION_LICENSE, FEATURE_PROPERTY_MGMT, False, IniPath) Then
        GrantFeature FEATURE_PROPERTY_MGMT
    End If

    If modConfigIni.GetIniBoolean(CONFIG_SECTION_LICENSE, FEATURE_WINE_MGMT, False, IniPath) Then
        GrantFeature FEATURE_WINE_MGMT
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeLicensing", "Licensing initialized."
    InitializeLicensing = True
    Exit Function

ErrorHandler:
    InitializeLicensing = False
    modErrorHandler.HandleError MODULE_NAME, "InitializeLicensing", Err
End Function

Public Function IsFeatureEnabled(ByVal FeatureName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim normalizedFeature As String

    EnsureLicenseStore

    normalizedFeature = NormalizeFeatureCode(FeatureName)
    If LenB(normalizedFeature) = 0 Then
        IsFeatureEnabled = False
        Exit Function
    End If

    If Not IsKnownFeature(normalizedFeature) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".IsFeatureEnabled", "Unknown feature requested: " & normalizedFeature
        IsFeatureEnabled = False
        Exit Function
    End If

    IsFeatureEnabled = FeatureLicenses.Exists(normalizedFeature)
    Exit Function

ErrorHandler:
    IsFeatureEnabled = False
    modErrorHandler.HandleError MODULE_NAME, "IsFeatureEnabled", Err
End Function

Public Sub GrantFeature(ByVal FeatureCode As String)
    On Error GoTo ErrorHandler

    Dim normalizedCode As String

    EnsureLicenseStore

    normalizedCode = NormalizeFeatureCode(FeatureCode)
    If LenB(normalizedCode) = 0 Then
        Exit Sub
    End If

    If Not IsKnownFeature(normalizedCode) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".GrantFeature", "Unknown feature ignored: " & normalizedCode
        Exit Sub
    End If

    If Not FeatureLicenses.Exists(normalizedCode) Then
        FeatureLicenses.Add normalizedCode, True
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "GrantFeature", Err
End Sub

Private Sub LoadFeatureList(ByVal FeatureList As String)
    On Error GoTo ErrorHandler

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

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadFeatureList", Err
End Sub

Private Function IsKnownFeature(ByVal FeatureCode As String) As Boolean
    Select Case FeatureCode
        Case FEATURE_CORE, FEATURE_CAMT054, FEATURE_PROPERTY_MGMT, FEATURE_WINE_MGMT
            IsKnownFeature = True
        Case Else
            IsKnownFeature = False
    End Select
End Function

Private Sub EnsureLicenseStore()
    If FeatureLicenses Is Nothing Then
        Set FeatureLicenses = CreateObject("Scripting.Dictionary")
    End If
End Sub
Private Function NormalizeFeatureCode(ByVal FeatureCode As String) As String
    NormalizeFeatureCode = UCase$(Trim$(FeatureCode))
End Function