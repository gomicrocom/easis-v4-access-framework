Option Compare Database
Option Explicit

'===============================================================================
' Module    : modErrorHandler
' Purpose   : Centralized error handling for framework modules.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modErrorHandler"

Public Function HandleError(ByVal ModuleName As String, ByVal ProcedureName As String, ByVal SourceError As ErrObject) As String
    Dim errorMessage As String

    errorMessage = ModuleName & "." & ProcedureName & " failed"

    If SourceError Is Nothing Then
        HandleError = errorMessage
        modLoggingHandler.LogError ModuleName & "." & ProcedureName, errorMessage
        Exit Function
    End If

    If LenB(SourceError.Description) > 0 Then
        errorMessage = errorMessage & " | " & SourceError.Description
    End If

    HandleError = errorMessage
    modLoggingHandler.LogError ModuleName & "." & ProcedureName, errorMessage, SourceError.Number
End Function

Public Function ExecuteSafely(ByVal ProcedureName As String) As Boolean
    On Error GoTo ErrorHandler

    Application.Run ProcedureName
    ExecuteSafely = True
    Exit Function

ErrorHandler:
    ExecuteSafely = False
    HandleError MODULE_NAME, "ExecuteSafely", Err
End Function