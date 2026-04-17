Attribute VB_Name = "modErrorHandler"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modErrorHandler
' Purpose   : Centralized error handling for framework modules.
' Author    : Codex
' Project   : Easis Version 4
'===============================================================================

Private Const MODULE_NAME As String = "modErrorHandler"

Public Sub HandleError(ByVal SourceModule As String, ByVal SourceProcedure As String, Optional ByVal ReRaise As Boolean = False, Optional ByVal UserMessage As String = vbNullString)
    Dim logMessage As String

    logMessage = BuildErrorMessage(SourceModule, SourceProcedure, UserMessage)
    modLoggingHandler.LogError SourceModule & "." & SourceProcedure, logMessage, Err.Number

    If ReRaise Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Public Function BuildErrorMessage(ByVal SourceModule As String, ByVal SourceProcedure As String, Optional ByVal UserMessage As String = vbNullString) As String
    Dim prefix As String

    prefix = SourceModule & "." & SourceProcedure & " failed"
    If LenB(UserMessage) > 0 Then
        prefix = prefix & ": " & UserMessage
    End If

    If LenB(Err.Description) > 0 Then
        BuildErrorMessage = prefix & " | " & Err.Description
    Else
        BuildErrorMessage = prefix
    End If
End Function

Public Function ExecuteSafely(ByVal ProcedureName As String) As Boolean
    On Error GoTo ErrorHandler

    Application.Run ProcedureName
    ExecuteSafely = True
    Exit Function

ErrorHandler:
    ExecuteSafely = False
    HandleError MODULE_NAME, "ExecuteSafely"
End Function
