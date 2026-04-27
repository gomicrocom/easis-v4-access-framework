Attribute VB_Name = "modBootstrap"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modBootstrap
' Purpose   : Application startup sequence for the framework.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modBootstrap"

Public Function BootstrapApplication(Optional ByVal IniPath As String = vbNullString) As Boolean
    On Error GoTo ErrorHandler

    ResetApplicationState
    modTenantContext.ClearTenantContext
    modSessionContext.ClearSessionContext

    modLoggingHandler.LogInfo MODULE_NAME & ".BootstrapApplication", "Bootstrap started."

    If Not modConfigIni.InitializeConfiguration(IniPath) Then
        Err.Raise vbObjectError + 2200, MODULE_NAME & ".BootstrapApplication", "Configuration initialization failed."
    End If
    modLoggingHandler.LogInfo MODULE_NAME & ".BootstrapApplication", "Configuration initialized."

    modLicenseHandler.InitializeLicenses
    modLoggingHandler.LogInfo MODULE_NAME & ".BootstrapApplication", "Licensing initialized."

    modTenantContext.InitializeTenantContext ConfigFilePath
    modLoggingHandler.LogInfo MODULE_NAME & ".BootstrapApplication", "Tenant context initialized."

    modSessionContext.InitializeSessionContext ConfigFilePath
    modLoggingHandler.LogInfo MODULE_NAME & ".BootstrapApplication", "Session context initialized."

    modTranslationService.InitializeTranslations
    modLoggingHandler.LogInfo MODULE_NAME & ".BootstrapApplication", "Translations initialized."

    modModuleManager.ResetModuleState
    modModuleManager.InitializeModules
    modLoggingHandler.LogInfo MODULE_NAME & ".BootstrapApplication", "Modules initialized."

    IsBootstrapped = True
    BootstrapApplication = True

    modLoggingHandler.LogInfo MODULE_NAME & ".BootstrapApplication", "Bootstrap completed successfully."
    Exit Function

ErrorHandler:
    BootstrapApplication = False
    IsBootstrapped = False
    modErrorHandler.HandleError MODULE_NAME, "BootstrapApplication", Err
End Function

Public Function EnsureBootstrapped(Optional ByVal IniPath As String = vbNullString) As Boolean
    If Not IsBootstrapped Then
        EnsureBootstrapped = BootstrapApplication(IniPath)
    Else
        EnsureBootstrapped = True
    End If
End Function
