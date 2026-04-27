Attribute VB_Name = "modModuleManager"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modModuleManager
' Purpose   : Runtime activation service for optional framework modules.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modModuleManager"
Private Const ERR_MODULE_REQUIRED As Long = vbObjectError + 2700

Public Const MODULE_CORE As String = "CORE"
Public Const MODULE_CAMT054 As String = "CAMT054"
Public Const MODULE_PROPERTY_MGMT As String = "PROPERTY_MGMT"
Public Const MODULE_WINE_MGMT As String = "WINE_MGMT"

Private mActiveModules As Object

Public Sub InitializeModules()
    On Error GoTo ErrorHandler

    Dim licensedFeatures As Collection
    Dim featureName As Variant

    EnsureModuleStore

    If modLicenseHandler.IsFeatureLicensed(MODULE_CORE) Then
        ActivateModule MODULE_CORE
    End If

    Set licensedFeatures = modLicenseHandler.GetLicensedFeatures()
    For Each featureName In licensedFeatures
        ActivateModule CStr(featureName)
    Next featureName

    modLoggingHandler.LogInfo MODULE_NAME & ".InitializeModules", _
        "Modules initialized: " & BuildActiveModulesSummary()
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "InitializeModules", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function IsModuleActive(ByVal moduleName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim normalizedModule As String

    normalizedModule = NormalizeModuleName(moduleName)
    If LenB(normalizedModule) = 0 Then
        IsModuleActive = False
        Exit Function
    End If

    EnsureModuleStore
    IsModuleActive = mActiveModules.Exists(normalizedModule)
    Exit Function

ErrorHandler:
    IsModuleActive = False
    modErrorHandler.HandleError MODULE_NAME, "IsModuleActive", Err
End Function

Public Function GetActiveModules() As Collection
    On Error GoTo ErrorHandler

    Dim activeModules As Collection
    Dim moduleName As Variant

    EnsureModuleStore
    Set activeModules = New Collection

    For Each moduleName In mActiveModules.Keys
        activeModules.Add CStr(moduleName)
    Next moduleName

    Set GetActiveModules = activeModules
    Exit Function

ErrorHandler:
    Set GetActiveModules = New Collection
    modErrorHandler.HandleError MODULE_NAME, "GetActiveModules", Err
End Function

Public Sub RequireModule(ByVal moduleName As String, Optional ByVal RaiseError As Boolean = True)
    On Error GoTo ErrorHandler

    Dim normalizedModule As String
    Dim messageText As String

    normalizedModule = NormalizeModuleName(moduleName)
    If LenB(normalizedModule) = 0 Then
        messageText = "Module name is required."
        modLoggingHandler.LogWarning MODULE_NAME & ".RequireModule", messageText

        If RaiseError Then
            Err.Raise ERR_MODULE_REQUIRED, MODULE_NAME & ".RequireModule", messageText
        End If
        Exit Sub
    End If

    If IsModuleActive(normalizedModule) Then
        Exit Sub
    End If

    messageText = "Module is not active: " & normalizedModule
    modLoggingHandler.LogWarning MODULE_NAME & ".RequireModule", messageText

    If RaiseError Then
        Err.Raise ERR_MODULE_REQUIRED, MODULE_NAME & ".RequireModule", messageText
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RequireModule", Err

    If RaiseError Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Public Sub ResetModuleState()
    On Error GoTo ErrorHandler

    EnsureModuleStore
    mActiveModules.RemoveAll

    modLoggingHandler.LogInfo MODULE_NAME & ".ResetModuleState", "Module state cleared."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ResetModuleState", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ActivateModule(ByVal moduleName As String)
    On Error GoTo ErrorHandler

    Dim normalizedModule As String

    normalizedModule = NormalizeModuleName(moduleName)
    If LenB(normalizedModule) = 0 Then
        Exit Sub
    End If

    EnsureModuleStore

    If Not modLicenseHandler.IsFeatureLicensed(normalizedModule) Then
        Exit Sub
    End If

    If Not mActiveModules.Exists(normalizedModule) Then
        mActiveModules.Add normalizedModule, True
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ActivateModule", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub EnsureModuleStore()
    On Error GoTo ErrorHandler

    If mActiveModules Is Nothing Then
        Set mActiveModules = CreateObject("Scripting.Dictionary")
        mActiveModules.CompareMode = vbTextCompare
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureModuleStore", Err
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function NormalizeModuleName(ByVal moduleName As String) As String
    NormalizeModuleName = UCase$(Trim$(moduleName))
End Function

Private Function BuildActiveModulesSummary() As String
    On Error GoTo ErrorHandler

    Dim moduleName As Variant
    Dim summary As String

    EnsureModuleStore

    If mActiveModules.Count = 0 Then
        BuildActiveModulesSummary = "(none)"
        Exit Function
    End If

    For Each moduleName In mActiveModules.Keys
        If LenB(summary) > 0 Then
            summary = summary & ", "
        End If
        summary = summary & CStr(moduleName)
    Next moduleName

    BuildActiveModulesSummary = summary
    Exit Function

ErrorHandler:
    BuildActiveModulesSummary = "(unknown)"
    modErrorHandler.HandleError MODULE_NAME, "BuildActiveModulesSummary", Err
End Function
