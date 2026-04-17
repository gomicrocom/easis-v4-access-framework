Attribute VB_Name = "modSecurityHandler"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modSecurityHandler
' Purpose   : Session-based role checks for framework security decisions.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modSecurityHandler"
Private Const ROLE_CODE_SUPERVISOR As String = "SUPERVISOR"

Public Function GetCurrentUserRoleCode() As String
    On Error GoTo ErrorHandler

    Dim resolvedRole As String
    Dim resolvedUserId As String

    If Not IsSessionInitialized() Then
        GetCurrentUserRoleCode = vbNullString
        Exit Function
    End If

    resolvedRole = NormalizeRoleCode(CurrentRoleCode)

    If LenB(resolvedRole) = 0 Then
        resolvedUserId = Trim$(currentUserId)

        If LenB(resolvedUserId) > 0 Then
            resolvedRole = NormalizeRoleCode(modUserRepository.GetUserRoleCode(resolvedUserId, ROLE_CODE_USER))
        End If
    End If

    GetCurrentUserRoleCode = resolvedRole
    Exit Function

ErrorHandler:
    GetCurrentUserRoleCode = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "GetCurrentUserRoleCode", Err
End Function

Public Function HasRole(ByVal RoleCode As String) As Boolean
    On Error GoTo ErrorHandler

    Dim currentRole As String
    Dim targetRole As String

    HasRole = False

    If Not IsSessionInitialized() Then
        Exit Function
    End If

    currentRole = GetCurrentUserRoleCode()
    targetRole = NormalizeRoleCode(RoleCode)

    If LenB(currentRole) = 0 Or LenB(targetRole) = 0 Then
        Exit Function
    End If

    HasRole = (currentRole = targetRole)
    Exit Function

ErrorHandler:
    HasRole = False
    modErrorHandler.HandleError MODULE_NAME, "HasRole", Err
End Function

Public Function RequireRole(ByVal RoleCode As String, Optional ByVal ActionName As String = "") As Boolean
    On Error GoTo ErrorHandler

    RequireRole = False

    If HasRole(RoleCode) Then
        RequireRole = True
        Exit Function
    End If

    modLoggingHandler.LogWarning MODULE_NAME & ".RequireRole", BuildDeniedMessage(RoleCode, ActionName)
    Exit Function

ErrorHandler:
    RequireRole = False
    modErrorHandler.HandleError MODULE_NAME, "RequireRole", Err
End Function

Public Function IsAdmin() As Boolean
    IsAdmin = HasRole(ROLE_CODE_ADMIN)
End Function

Public Function IsSupervisor() As Boolean
    IsSupervisor = HasRole(ROLE_CODE_SUPERVISOR)
End Function

Private Function NormalizeRoleCode(ByVal RoleCode As String) As String
    NormalizeRoleCode = UCase$(Trim$(RoleCode))
End Function

Private Function BuildDeniedMessage(ByVal RoleCode As String, ByVal ActionName As String) As String
    Dim messageText As String
    Dim requiredRole As String

    requiredRole = NormalizeRoleCode(RoleCode)

    messageText = "Access denied. Required role='" & requiredRole & "'"

    If LenB(Trim$(ActionName)) > 0 Then
        messageText = messageText & ", Action='" & Trim$(ActionName) & "'"
    End If

    If IsSessionInitialized() Then
        messageText = messageText & _
            ", UserId='" & currentUserId & "'" & _
            ", CurrentRole='" & GetCurrentUserRoleCode() & "'"
    Else
        messageText = messageText & ", Session=<uninitialized>"
    End If

    BuildDeniedMessage = messageText
End Function

