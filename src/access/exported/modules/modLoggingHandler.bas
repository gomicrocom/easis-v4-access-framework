Attribute VB_Name = "modLoggingHandler"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modLoggingHandler
' Purpose   : Central logging entry points for the Access framework.
' Author    : Codex
' Project   : Easis Version 4
'===============================================================================

Private Const MODULE_NAME As String = "modLoggingHandler"

Public Sub LogDebug(ByVal SourceProcedure As String, ByVal Message As String)
    WriteLog LogLevelDebug, SourceProcedure, Message
End Sub

Public Sub LogInfo(ByVal SourceProcedure As String, ByVal Message As String)
    WriteLog LogLevelInfo, SourceProcedure, Message
End Sub

Public Sub LogWarning(ByVal SourceProcedure As String, ByVal Message As String)
    WriteLog LogLevelWarning, SourceProcedure, Message
End Sub

Public Sub LogError(ByVal SourceProcedure As String, ByVal Message As String, Optional ByVal ErrorNumber As Long = 0)
    Dim fullMessage As String

    fullMessage = Message
    If ErrorNumber <> 0 Then
        fullMessage = fullMessage & " (Err " & CStr(ErrorNumber) & ")"
    End If

    WriteLog LogLevelError, SourceProcedure, fullMessage
End Sub

Public Sub WriteLog(ByVal EntryLevel As LogLevel, ByVal SourceProcedure As String, ByVal Message As String)
    On Error GoTo SafeExit

    If EntryLevel < CurrentLogLevel Then
        Exit Sub
    End If

    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); " | "; _
                LevelName(EntryLevel); " | "; _
                SourceProcedure; " | "; _
                Message

SafeExit:
End Sub

Private Function LevelName(ByVal EntryLevel As LogLevel) As String
    Select Case EntryLevel
        Case LogLevelDebug
            LevelName = "DEBUG"
        Case LogLevelWarning
            LevelName = "WARN"
        Case LogLevelError
            LevelName = "ERROR"
        Case Else
            LevelName = "INFO"
    End Select
End Function
