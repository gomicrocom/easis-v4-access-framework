Attribute VB_Name = "modLicenseHandler"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modLicenseHandler
' Purpose   : Feature-based licensing helpers with placeholder logic.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modLicenseHandler"
Private Const FEATURE_CORE As String = "CORE"
Private Const FEATURE_CAMT054 As String = "CAMT054"
Private Const FEATURE_PROPERTY_MGMT As String = "PROPERTY_MGMT"
Private Const FEATURE_WINE_MGMT As String = "WINE_MGMT"

Public Function InitializeLicensing(Optional ByVal IniPath As String = vbNullString) As Boolean
    On Error GoTo ErrorHandler

    Dim featureList As String

    FeatureLicenses.RemoveAll

    featureList = modConfigIni.GetIniString(CONFIG_SECTION_LICENSE, "EnabledFeatures", vbNullString, IniPath)
    LoadFeatureList featureList

    If modConfigIni.GetIniBoolean(CONFIG_SECTION_LICENSE, "EnableCore", True, IniPath) Then
        GrantFeature FEATURE_CORE
    End If

    If modConfigIni.GetIniBoolean(CONFIG_SECTION_LICENSE, FEATURE_CAMT054, False, IniPath) Then
        GrantFeature FEATURE_CAMT054
    End If

    If modConfigIni.GetIniBoolean(CONFIG_SECTION_LICENSE, FEATURE_PROPERTY_MGMT, False, IniPath) Then
        GrantFeature FEATURE_PROPERTY_MGMT
    End If

    If modConfigIni.GetIniBoolean(CONFIG_SECTION_LICENSE, FEATURE_WINE_MGMT, False, IniPath) Then
        GrantFeature FEATURE_WINE_MGMT
    End If

    InitializeLicensing = True
    Exit Function

ErrorHandler:
    InitializeLicensing = False
    modErrorHandler.HandleError MODULE_NAME, "InitializeLicensing", Err
End Function

Public Function IsFeatureEnabled(ByVal FeatureName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim normalizedFeature As String

    normalizedFeature = NormalizeFeatureCode(FeatureName)
    If LenB(normalizedFeature) = 0 Then
        Exit Function
    End If

    IsFeatureEnabled = FeatureLicenses.Exists(normalizedFeature)
    Exit Function

ErrorHandler:
    IsFeatureEnabled = False
    modErrorHandler.HandleError MODULE_NAME, "IsFeatureEnabled", Err
End Function

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

Public Sub LoadDefaultFeatures()
    GrantFeature FEATURE_CAMT054
    GrantFeature FEATURE_PROPERTY_MGMT
    GrantFeature FEATURE_WINE_MGMT
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
