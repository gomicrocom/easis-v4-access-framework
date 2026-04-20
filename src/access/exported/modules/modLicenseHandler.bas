Option Compare Database
Option Explicit

'===============================================================================
' Module    : modLicenseHandler
' Purpose   : Runtime feature licensing service for framework modules.
' Author    : Codex
' Version   : 0.2.0
'===============================================================================

Private Const MODULE_NAME As String = "modLicenseHandler"
Private Const ERR_FEATURE_REQUIRED As Long = vbObjectError + 2600

Public Const FEATURE_CORE As String = "CORE"
Public Const FEATURE_CAMT054 As String = "CAMT054"
Public Const FEATURE_PROPERTY_MGMT As String = "PROPERTY_MGMT"
Public Const FEATURE_WINE_MGMT As String = "WINE_MGMT"

Public Sub InitializeLicenses()
    On Error GoTo ErrorHandler

    LoadRuntimeLicenses ConfigFilePath
    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeLicenses", _
                              "Licensing initialized with " & CStr(GetLicenseStore().Count) & " active feature(s)."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "InitializeLicenses", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function IsFeatureLicensed(ByVal FeatureName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim normalizedFeature As String

    normalizedFeature = NormalizeFeatureCode(FeatureName)
    If LenB(normalizedFeature) = 0 Then
        IsFeatureLicensed = False
        Exit Function
    End If

    IsFeatureLicensed = GetLicenseStore().Exists(normalizedFeature)
    Exit Function

ErrorHandler:
    IsFeatureLicensed = False
    modErrorHandler.HandleError MODULE_NAME, "IsFeatureLicensed", Err
End Function

Public Function GetLicensedFeatures() As Collection
    On Error GoTo ErrorHandler

    Dim licensedFeatures As Collection
    Dim featureKey As Variant

    Set licensedFeatures = New Collection

    For Each featureKey In GetLicenseStore().Keys
        licensedFeatures.Add CStr(featureKey)
    Next featureKey

    Set GetLicensedFeatures = licensedFeatures
    Exit Function

ErrorHandler:
    Set GetLicensedFeatures = New Collection
    modErrorHandler.HandleError MODULE_NAME, "GetLicensedFeatures", Err
End Function

Public Sub RequireFeature(ByVal FeatureName As String, Optional ByVal RaiseError As Boolean = True)
    On Error GoTo ErrorHandler

    Dim normalizedFeature As String
    Dim messageText As String

    normalizedFeature = NormalizeFeatureCode(FeatureName)
    If LenB(normalizedFeature) = 0 Then
        messageText = "Feature name is required."
        modLoggingHandler.LogWarning MODULE_NAME & ".RequireFeature", messageText

        If RaiseError Then
            Err.Raise ERR_FEATURE_REQUIRED, MODULE_NAME & ".RequireFeature", messageText
        End If
        Exit Sub
    End If

    If IsFeatureLicensed(normalizedFeature) Then
        Exit Sub
    End If

    messageText = "Feature is not licensed: " & normalizedFeature
    modLoggingHandler.LogWarning MODULE_NAME & ".RequireFeature", messageText

    If RaiseError Then
        Err.Raise ERR_FEATURE_REQUIRED, MODULE_NAME & ".RequireFeature", messageText
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RequireFeature", Err

    If RaiseError Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Public Function InitializeLicensing(Optional ByVal IniPath As String = vbNullString) As Boolean
    On Error GoTo ErrorHandler

    LoadRuntimeLicenses IniPath
    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeLicensing", _
                              "Licensing initialized with " & CStr(GetLicenseStore().Count) & " active feature(s)."
    InitializeLicensing = True
    Exit Function

ErrorHandler:
    InitializeLicensing = False
    modErrorHandler.HandleError MODULE_NAME, "InitializeLicensing", Err
End Function

Public Function IsFeatureEnabled(ByVal FeatureName As String) As Boolean
    IsFeatureEnabled = IsFeatureLicensed(FeatureName)
End Function

Private Sub LoadRuntimeLicenses(Optional ByVal IniPath As String = vbNullString)
    On Error GoTo ErrorHandler

    Dim featureList As String

    ClearLicenseStore

    AddLicensedFeature FEATURE_CORE

    featureList = modConfigIni.GetIniString(CONFIG_SECTION_LICENSE, "EnabledFeatures", vbNullString, IniPath)
    LoadFeatureList featureList

    If modConfigIni.GetIniBoolean(CONFIG_SECTION_LICENSE, FEATURE_CAMT054, False, IniPath) Then
        AddLicensedFeature FEATURE_CAMT054
    End If

    If modConfigIni.GetIniBoolean(CONFIG_SECTION_LICENSE, FEATURE_PROPERTY_MGMT, False, IniPath) Then
        AddLicensedFeature FEATURE_PROPERTY_MGMT
    End If

    If modConfigIni.GetIniBoolean(CONFIG_SECTION_LICENSE, FEATURE_WINE_MGMT, False, IniPath) Then
        AddLicensedFeature FEATURE_WINE_MGMT
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadRuntimeLicenses", Err
    Err.Raise Err.Number, Err.Source, Err.Description
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
        token = NormalizeFeatureCode(tokens(index))
        If LenB(token) > 0 Then
            AddLicensedFeature token
        End If
    Next index

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadFeatureList", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub AddLicensedFeature(ByVal FeatureName As String)
    On Error GoTo ErrorHandler

    Dim normalizedFeature As String

    normalizedFeature = NormalizeFeatureCode(FeatureName)
    If LenB(normalizedFeature) = 0 Then
        Exit Sub
    End If

    If Not GetLicenseStore().Exists(normalizedFeature) Then
        GetLicenseStore().Add normalizedFeature, True
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "AddLicensedFeature", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ClearLicenseStore()
    On Error GoTo ErrorHandler

    GetLicenseStore().RemoveAll
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ClearLicenseStore", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function GetLicenseStore() As Object
    Set GetLicenseStore = modGlobals.FeatureLicenses
End Function
