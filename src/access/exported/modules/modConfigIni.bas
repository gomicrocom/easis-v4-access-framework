Attribute VB_Name = "modConfigIni"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modConfigIni
' Purpose   : Reads framework configuration from INI files.
' Author    : Codex
' Project   : Easis Version 4
'===============================================================================

#If VBA7 Then
    Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpAppName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
#Else
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpAppName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
#End If

Private Const MODULE_NAME As String = "modConfigIni"
Private Const INI_BUFFER_SIZE As Long = 2048

Public Function InitializeConfiguration(Optional ByVal IniPath As String = vbNullString) As Boolean
    On Error GoTo ErrorHandler

    Dim resolvedPath As String

    resolvedPath = ResolveConfigPath(IniPath)
    If LenB(resolvedPath) = 0 Then
        Err.Raise vbObjectError + 2100, MODULE_NAME & ".InitializeConfiguration", "Configuration file path could not be resolved."
    End If

    ConfigFilePath = resolvedPath
    CurrentLanguage = GetIniString(CONFIG_SECTION_APP, "Language", DEFAULT_LANGUAGE, ConfigFilePath)
    CurrentLogLevel = ParseLogLevel(GetIniString(CONFIG_SECTION_LOGGING, "Level", CStr(DEFAULT_LOG_LEVEL), ConfigFilePath))

    InitializeConfiguration = True
    Exit Function

ErrorHandler:
    InitializeConfiguration = False
    modErrorHandler.HandleError MODULE_NAME, "InitializeConfiguration"
End Function

Public Function ResolveConfigPath(Optional ByVal IniPath As String = vbNullString) As String
    Dim candidate As String

    candidate = Trim$(IniPath)
    If LenB(candidate) > 0 Then
        ResolveConfigPath = candidate
        Exit Function
    End If

    candidate = CurrentProject.Path & "\config\easis.ini"
    If LenB(Dir$(candidate, vbNormal)) > 0 Then
        ResolveConfigPath = candidate
        Exit Function
    End If

    candidate = CurrentProject.Path & "\easis.ini"
    If LenB(Dir$(candidate, vbNormal)) > 0 Then
        ResolveConfigPath = candidate
        Exit Function
    End If
End Function

Public Function GetIniString(ByVal SectionName As String, ByVal KeyName As String, Optional ByVal DefaultValue As String = vbNullString, Optional ByVal IniPath As String = vbNullString) As String
    On Error GoTo ErrorHandler

    Dim buffer As String
    Dim charsRead As Long
    Dim effectivePath As String

    effectivePath = ResolveConfigPath(IniPath)
    If LenB(effectivePath) = 0 Then
        GetIniString = DefaultValue
        Exit Function
    End If

    buffer = String$(INI_BUFFER_SIZE, vbNullChar)
    charsRead = GetPrivateProfileString(SectionName, ByVal KeyName, DefaultValue, buffer, Len(buffer), effectivePath)

    If charsRead > 0 Then
        GetIniString = Left$(buffer, charsRead)
    Else
        GetIniString = DefaultValue
    End If
    Exit Function

ErrorHandler:
    GetIniString = DefaultValue
    modErrorHandler.HandleError MODULE_NAME, "GetIniString"
End Function

Public Function GetIniBoolean(ByVal SectionName As String, ByVal KeyName As String, Optional ByVal DefaultValue As Boolean = False, Optional ByVal IniPath As String = vbNullString) As Boolean
    Dim rawValue As String

    rawValue = NormalizeToken(GetIniString(SectionName, KeyName, BoolToIni(DefaultValue), IniPath))

    Select Case rawValue
        Case "1", "TRUE", "YES", "Y", "ON"
            GetIniBoolean = True
        Case "0", "FALSE", "NO", "N", "OFF"
            GetIniBoolean = False
        Case Else
            GetIniBoolean = DefaultValue
    End Select
End Function

Public Function GetIniLong(ByVal SectionName As String, ByVal KeyName As String, Optional ByVal DefaultValue As Long = 0, Optional ByVal IniPath As String = vbNullString) As Long
    Dim rawValue As String

    rawValue = Trim$(GetIniString(SectionName, KeyName, CStr(DefaultValue), IniPath))
    If IsNumeric(rawValue) Then
        GetIniLong = CLng(rawValue)
    Else
        GetIniLong = DefaultValue
    End If
End Function

Private Function ParseLogLevel(ByVal Value As String) As LogLevel
    Dim normalized As String

    normalized = NormalizeToken(Value)

    Select Case normalized
        Case "1", "DEBUG"
            ParseLogLevel = LogLevelDebug
        Case "3", "WARN", "WARNING"
            ParseLogLevel = LogLevelWarning
        Case "4", "ERR", "ERROR"
            ParseLogLevel = LogLevelError
        Case Else
            ParseLogLevel = LogLevelInfo
    End Select
End Function

Private Function NormalizeToken(ByVal Value As String) As String
    NormalizeToken = UCase$(Trim$(Value))
End Function

Private Function BoolToIni(ByVal Value As Boolean) As String
    If Value Then
        BoolToIni = "1"
    Else
        BoolToIni = "0"
    End If
End Function
