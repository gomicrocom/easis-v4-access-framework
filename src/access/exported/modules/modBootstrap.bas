Attribute VB_Name = "modBootstrap"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modBootstrap
' Purpose   : Framework startup orchestration for Easis Version 4.
' Author    : Codex
' Project   : Easis Version 4
'===============================================================================

Private Const MODULE_NAME As String = "modBootstrap"

Public Function BootstrapApplication(Optional ByVal IniPath As String = vbNullString) As Boolean
    On Error GoTo ErrorHandler

    ResetApplicationState

    If Not modConfigIni.InitializeConfiguration(IniPath) Then
        Err.Raise vbObjectError + 2200, MODULE_NAME & ".BootstrapApplication", "Configuration initialization failed."
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".BootstrapApplication", "Configuration initialized."

    If Not modLicenseHandler.InitializeLicensing(ConfigFilePath) Then
        Err.Raise vbObjectError + 2201, MODULE_NAME & ".BootstrapApplication", "Licensing initialization failed."
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".BootstrapApplication", "Licensing initialized."

    IsBootstrapped = True
    BootstrapApplication = True
    Exit Function

ErrorHandler:
    BootstrapApplication = False
    IsBootstrapped = False
    modErrorHandler.HandleError MODULE_NAME, "BootstrapApplication"
End Function

Public Function EnsureBootstrapped(Optional ByVal IniPath As String = vbNullString) As Boolean
    If Not IsBootstrapped Then
        EnsureBootstrapped = BootstrapApplication(IniPath)
    Else
        EnsureBootstrapped = True
    End If
End Function
