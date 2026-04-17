Attribute VB_Name = "modGlobals"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modGlobals
' Purpose   : Shared application constants, state, and framework accessors.
' Author    : Codex
' Project   : Easis Version 4
'===============================================================================

Public Const APP_NAME As String = "Easis Version 4"
Public Const APP_VERSION As String = "0.1.0"

Public Const CONFIG_SECTION_APP As String = "Application"
Public Const CONFIG_SECTION_LICENSE As String = "License"
Public Const CONFIG_SECTION_LOGGING As String = "Logging"

Public Const DEFAULT_LANGUAGE As String = "de-CH"
Public Const DEFAULT_LOG_LEVEL As Long = 2

Public Enum LogLevel
    LogLevelDebug = 1
    LogLevelInfo = 2
    LogLevelWarning = 3
    LogLevelError = 4
End Enum

Private mIsBootstrapped As Boolean
Private mConfigFilePath As String
Private mCurrentLanguage As String
Private mLogLevel As LogLevel
Private mFeatureLicenses As Object

Public Property Get IsBootstrapped() As Boolean
    IsBootstrapped = mIsBootstrapped
End Property

Public Property Let IsBootstrapped(ByVal Value As Boolean)
    mIsBootstrapped = Value
End Property

Public Property Get ConfigFilePath() As String
    ConfigFilePath = mConfigFilePath
End Property

Public Property Let ConfigFilePath(ByVal Value As String)
    mConfigFilePath = Trim$(Value)
End Property

Public Property Get CurrentLanguage() As String
    If LenB(mCurrentLanguage) = 0 Then
        mCurrentLanguage = DEFAULT_LANGUAGE
    End If
    CurrentLanguage = mCurrentLanguage
End Property

Public Property Let CurrentLanguage(ByVal Value As String)
    If LenB(Trim$(Value)) = 0 Then
        mCurrentLanguage = DEFAULT_LANGUAGE
    Else
        mCurrentLanguage = Trim$(Value)
    End If
End Property

Public Property Get CurrentLogLevel() As LogLevel
    If mLogLevel = 0 Then
        mLogLevel = DEFAULT_LOG_LEVEL
    End If
    CurrentLogLevel = mLogLevel
End Property

Public Property Let CurrentLogLevel(ByVal Value As LogLevel)
    If Value < LogLevelDebug Or Value > LogLevelError Then
        mLogLevel = DEFAULT_LOG_LEVEL
    Else
        mLogLevel = Value
    End If
End Property

Public Property Get FeatureLicenses() As Object
    If mFeatureLicenses Is Nothing Then
        Set mFeatureLicenses = CreateObject("Scripting.Dictionary")
        mFeatureLicenses.CompareMode = vbTextCompare
    End If
    Set FeatureLicenses = mFeatureLicenses
End Property

Public Sub ResetApplicationState()
    mIsBootstrapped = False
    mConfigFilePath = vbNullString
    mCurrentLanguage = DEFAULT_LANGUAGE
    mLogLevel = DEFAULT_LOG_LEVEL

    If Not mFeatureLicenses Is Nothing Then
        mFeatureLicenses.RemoveAll
    End If
End Sub

Public Function NormalizeFeatureCode(ByVal FeatureCode As String) As String
    NormalizeFeatureCode = UCase$(Trim$(FeatureCode))
End Function
