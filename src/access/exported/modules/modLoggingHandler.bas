Option Compare Database
Option Explicit

'===============================================================================
' Module    : modLoggingHandler
' Purpose   : Central logging helpers for framework diagnostics.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modLoggingHandler"
Private Const LOG_LEVEL_INFO As String = "INFO"
Private Const LOG_LEVEL_WARN As String = "WARN"
Private Const LOG_LEVEL_ERROR As String = "ERROR"

Public Sub LogInfo(ByVal SourceProcedure As String, ByVal Message As String)
    WriteLog LOG_LEVEL_INFO, SourceProcedure, Message
End Sub

Public Sub LogWarning(ByVal SourceProcedure As String, ByVal Message As String)
    WriteLog LOG_LEVEL_WARN, SourceProcedure, Message
End Sub

Public Sub LogError(ByVal SourceProcedure As String, ByVal Message As String, Optional ByVal ErrorNumber As Long = 0)
    Dim fullMessage As String

    fullMessage = Message
    If ErrorNumber <> 0 Then
        fullMessage = fullMessage & " (Err " & CStr(ErrorNumber) & ")"
    End If

    WriteLog LOG_LEVEL_ERROR, SourceProcedure, fullMessage
End Sub

Public Sub WriteLog(ByVal EntryLevel As String, ByVal SourceProcedure As String, ByVal Message As String)
    On Error GoTo SafeExit

    If ShouldSkipLog(EntryLevel) Then
        Exit Sub
    End If

    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); " | "; _
                UCase$(Trim$(EntryLevel)); " | "; _
                SourceProcedure; " | "; _
                Message

    ' Placeholder: optional table-based logging can be added here later.

SafeExit:
End Sub

Private Function ShouldSkipLog(ByVal EntryLevel As String) As Boolean
    Dim normalizedEntry As String
    Dim normalizedCurrent As String

    normalizedEntry = UCase$(Trim$(EntryLevel))
    normalizedCurrent = UCase$(Trim$(CurrentLogLevel))

    Select Case normalizedCurrent
        Case LOG_LEVEL_ERROR
            ShouldSkipLog = (normalizedEntry <> LOG_LEVEL_ERROR)
        Case LOG_LEVEL_WARN
            ShouldSkipLog = (normalizedEntry = LOG_LEVEL_INFO)
        Case Else
            ShouldSkipLog = False
    End Select
End Function