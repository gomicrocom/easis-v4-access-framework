Attribute VB_Name = "modAuditHelper"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modAuditHelper
' Purpose   : Centralized audit-field handling for bound Access forms.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modAuditHelper"
Private Const FIELD_CREATED_AT As String = "created_at"
Private Const FIELD_CREATED_BY As String = "created_by"
Private Const FIELD_UPDATED_AT As String = "updated_at"
Private Const FIELD_UPDATED_BY As String = "updated_by"

Public Sub ApplyAuditFields(ByVal targetForm As Access.Form)
    On Error GoTo ErrorHandler

    Dim currentTimestamp As Date
    Dim currentUserName As String

    If targetForm Is Nothing Then
        Exit Sub
    End If

    If Not HasEditableRecordset(targetForm) Then
        Exit Sub
    End If

    currentTimestamp = Now()
    currentUserName = ResolveAuditUserName()

    If targetForm.NewRecord Then
        If HasRecordsetField(targetForm, FIELD_CREATED_AT) Then
            If IsValueEmpty(GetRecordsetFieldValue(targetForm, FIELD_CREATED_AT)) Then
                SetRecordsetFieldValue targetForm, FIELD_CREATED_AT, currentTimestamp
            End If
        End If

        If HasRecordsetField(targetForm, FIELD_CREATED_BY) Then
            If IsValueEmpty(GetRecordsetFieldValue(targetForm, FIELD_CREATED_BY)) Then
                SetRecordsetFieldValue targetForm, FIELD_CREATED_BY, currentUserName
            End If
        End If
    End If

    If HasRecordsetField(targetForm, FIELD_UPDATED_AT) Then
        SetRecordsetFieldValue targetForm, FIELD_UPDATED_AT, currentTimestamp
    End If

    If HasRecordsetField(targetForm, FIELD_UPDATED_BY) Then
        SetRecordsetFieldValue targetForm, FIELD_UPDATED_BY, currentUserName
    End If
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyAuditFields", Err
End Sub

Public Function ResolveAuditUserName() As String
    On Error GoTo SafeExit

    If modSessionContext.IsSessionInitialized Then
        ResolveAuditUserName = Trim$(modSessionContext.CurrentUserName)
        If LenB(ResolveAuditUserName) > 0 Then
            Exit Function
        End If

        ResolveAuditUserName = Trim$(modSessionContext.currentUserId)
        If LenB(ResolveAuditUserName) > 0 Then
            Exit Function
        End If
    End If

SafeExit:
    On Error Resume Next
    ResolveAuditUserName = Trim$(ResolveAuditUserName)
    If LenB(ResolveAuditUserName) = 0 Then
        ResolveAuditUserName = Trim$(Environ$("Username"))
    End If
    If LenB(ResolveAuditUserName) = 0 Then
        ResolveAuditUserName = "SYSTEM"
    End If
End Function

Public Function HasRecordsetField(ByVal targetForm As Access.Form, ByVal fieldName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentField As DAO.Field

    If Not HasEditableRecordset(targetForm) Then
        Exit Function
    End If

    For Each currentField In targetForm.Recordset.Fields
        If StrComp(currentField.Name, fieldName, vbTextCompare) = 0 Then
            HasRecordsetField = True
            Exit Function
        End If
    Next currentField

SafeExit:
    Set currentField = Nothing
End Function

Public Function GetRecordsetFieldValue(ByVal targetForm As Access.Form, ByVal fieldName As String) As Variant
    On Error GoTo SafeExit

    If HasRecordsetField(targetForm, fieldName) Then
        GetRecordsetFieldValue = targetForm.Recordset.Fields(fieldName).Value
    End If

SafeExit:
End Function

Public Sub SetRecordsetFieldValue(ByVal targetForm As Access.Form, ByVal fieldName As String, ByVal fieldValue As Variant)
    On Error GoTo SafeExit

    If HasRecordsetField(targetForm, fieldName) Then
        targetForm.Recordset.Fields(fieldName).Value = fieldValue
    End If

SafeExit:
End Sub

Private Function HasEditableRecordset(ByVal targetForm As Access.Form) As Boolean
    On Error GoTo SafeExit

    If targetForm Is Nothing Then
        Exit Function
    End If
    
    If targetForm.Recordset Is Nothing Then
        Exit Function
    End If

    HasEditableRecordset = True

SafeExit:
End Function

Private Function IsValueEmpty(ByVal fieldValue As Variant) As Boolean
    If IsNull(fieldValue) Or IsEmpty(fieldValue) Then
        IsValueEmpty = True
    ElseIf VarType(fieldValue) = vbString Then
        IsValueEmpty = (LenB(Trim$(CStr(fieldValue))) = 0)
    End If
End Function
