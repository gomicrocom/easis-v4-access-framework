Attribute VB_Name = "modOrderEditSessionService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modOrderEditSessionService
' Purpose   : Controlled temp-session entry points for frmOrderDetailNext.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modOrderEditSessionService"
Private Const FORM_ORDER_DETAIL_NEXT As String = "frmOrderDetailNext"

Public Const ORDER_EDIT_OPEN_ARGS_PREFIX As String = "ORDER_EDIT"

Public Function BuildOrderEditWorkspaceArgs(ByVal tmpOrderId As Long) As String
    If tmpOrderId <= 0 Then
        Exit Function
    End If

    BuildOrderEditWorkspaceArgs = ORDER_EDIT_OPEN_ARGS_PREFIX & ";" & CStr(tmpOrderId)
End Function

Public Function ParseOrderEditWorkspaceId(ByVal openArgs As String) As Long
    Dim normalizedOpenArgs As String
    Dim payloadParts() As String

    normalizedOpenArgs = Trim$(modDaoHelper.NzString(openArgs))
    If LenB(normalizedOpenArgs) = 0 Then
        Exit Function
    End If

    payloadParts = Split(normalizedOpenArgs, ";")
    If UBound(payloadParts) < 1 Then
        Exit Function
    End If

    If StrComp(Trim$(payloadParts(0)), ORDER_EDIT_OPEN_ARGS_PREFIX, vbTextCompare) <> 0 Then
        Exit Function
    End If

    ParseOrderEditWorkspaceId = modDaoHelper.NzLong(payloadParts(1), 0)
End Function

Public Function OpenOrderDetailNextStandaloneForAddress(ByVal addressId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim tmpOrderId As Long
    Dim frontendDb As DAO.Database

    Set frontendDb = modDb.GetFrontendDatabase()
    modLoggingHandler.LogInfo MODULE_NAME & ".OpenOrderDetailNextStandaloneForAddress", _
        "address_id=" & CStr(addressId) & "; tenant_backend_path=" & modDb.GetCurrentTenantBackendPath() & "; frontend_db_name=" & SafeDatabaseName(frontendDb) & "; open_requested=True."

    tmpOrderId = modOrderRepository.CreateTemporarySalesOrderForAddress(addressId)
    modLoggingHandler.LogInfo MODULE_NAME & ".OpenOrderDetailNextStandaloneForAddress", _
        "address_id=" & CStr(addressId) & "; tmp_order_id=" & CStr(tmpOrderId) & "."

    OpenOrderDetailNextStandaloneForAddress = OpenOrderDetailNextStandaloneByTempId(tmpOrderId)
    modLoggingHandler.LogInfo MODULE_NAME & ".OpenOrderDetailNextStandaloneForAddress", _
        "address_id=" & CStr(addressId) & "; tmp_order_id=" & CStr(tmpOrderId) & "; return_value=" & CStr(OpenOrderDetailNextStandaloneForAddress) & "."
    Exit Function

ErrorHandler:
    OpenOrderDetailNextStandaloneForAddress = False
    modErrorHandler.HandleError MODULE_NAME, "OpenOrderDetailNextStandaloneForAddress", Err
End Function

Public Function OpenOrderDetailNextStandaloneForExistingOrder(ByVal OrderId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim tmpOrderId As Long
    Dim frontendDb As DAO.Database

    Set frontendDb = modDb.GetFrontendDatabase()
    modLoggingHandler.LogInfo MODULE_NAME & ".OpenOrderDetailNextStandaloneForExistingOrder", _
        "order_id=" & CStr(OrderId) & "; tenant_backend_path=" & modDb.GetCurrentTenantBackendPath() & "; frontend_db_name=" & SafeDatabaseName(frontendDb) & "; open_requested=True."

    tmpOrderId = modOrderRepository.CreateTemporarySalesOrderForExistingOrder(OrderId)
    modLoggingHandler.LogInfo MODULE_NAME & ".OpenOrderDetailNextStandaloneForExistingOrder", _
        "order_id=" & CStr(OrderId) & "; tmp_order_id=" & CStr(tmpOrderId) & "."

    OpenOrderDetailNextStandaloneForExistingOrder = OpenOrderDetailNextStandaloneByTempId(tmpOrderId)
    modLoggingHandler.LogInfo MODULE_NAME & ".OpenOrderDetailNextStandaloneForExistingOrder", _
        "order_id=" & CStr(OrderId) & "; tmp_order_id=" & CStr(tmpOrderId) & "; return_value=" & CStr(OpenOrderDetailNextStandaloneForExistingOrder) & "."
    Exit Function

ErrorHandler:
    OpenOrderDetailNextStandaloneForExistingOrder = False
    modErrorHandler.HandleError MODULE_NAME, "OpenOrderDetailNextStandaloneForExistingOrder", Err
End Function

Public Function OpenOrderDetailNextWorkspaceForAddress(ByVal shellForm As Access.Form, ByVal addressId As Long) As Boolean
    Dim tmpOrderId As Long

    tmpOrderId = modOrderRepository.CreateTemporarySalesOrderForAddress(addressId)
    OpenOrderDetailNextWorkspaceForAddress = OpenOrderDetailNextWorkspaceByTempId(shellForm, tmpOrderId)
End Function

Public Function OpenOrderDetailNextWorkspaceForExistingOrder(ByVal shellForm As Access.Form, ByVal OrderId As Long) As Boolean
    Dim tmpOrderId As Long

    tmpOrderId = modOrderRepository.CreateTemporarySalesOrderForExistingOrder(OrderId)
    OpenOrderDetailNextWorkspaceForExistingOrder = OpenOrderDetailNextWorkspaceByTempId(shellForm, tmpOrderId)
End Function

Private Function OpenOrderDetailNextStandaloneByTempId(ByVal tmpOrderId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim loadedForm As Access.Form

    If tmpOrderId <= 0 Then
        modLoggingHandler.LogWarning MODULE_NAME & ".OpenOrderDetailNextStandaloneByTempId", _
            "Open aborted because tmp_order_id is not positive."
        Exit Function
    End If

    If Not VerifyTemporaryOrderExists(tmpOrderId) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".OpenOrderDetailNextStandaloneByTempId", _
            "Open aborted because tmp_order_id=" & CStr(tmpOrderId) & " could not be verified in frontend database."
        Exit Function
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".OpenOrderDetailNextStandaloneByTempId", _
        "Form opening started. tmp_order_id=" & CStr(tmpOrderId) & "."
    DoCmd.OpenForm FORM_ORDER_DETAIL_NEXT, acNormal, , , acFormEdit, acWindowNormal, BuildOrderEditWorkspaceArgs(tmpOrderId)

    If Not CurrentProject.AllForms(FORM_ORDER_DETAIL_NEXT).IsLoaded Then
        modLoggingHandler.LogWarning MODULE_NAME & ".OpenOrderDetailNextStandaloneByTempId", _
            "Form did not remain loaded after OpenForm. tmp_order_id=" & CStr(tmpOrderId) & "."
        Exit Function
    End If

    Set loadedForm = Forms(FORM_ORDER_DETAIL_NEXT)
    If GetLoadedFormTmpOrderId(loadedForm) <> tmpOrderId Then
        modLoggingHandler.LogWarning MODULE_NAME & ".OpenOrderDetailNextStandaloneByTempId", _
            "Loaded form tmp_order_id mismatch. expected=" & CStr(tmpOrderId) & "; actual=" & CStr(GetLoadedFormTmpOrderId(loadedForm)) & "."
        Exit Function
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".OpenOrderDetailNextStandaloneByTempId", _
        "Form loaded successfully. tmp_order_id=" & CStr(tmpOrderId) & "; loaded_tmp_order_id=" & CStr(GetLoadedFormTmpOrderId(loadedForm)) & "."
    OpenOrderDetailNextStandaloneByTempId = True
    Exit Function

ErrorHandler:
    OpenOrderDetailNextStandaloneByTempId = False
    modErrorHandler.HandleError MODULE_NAME, "OpenOrderDetailNextStandaloneByTempId", Err
End Function

Private Function OpenOrderDetailNextWorkspaceByTempId(ByVal shellForm As Access.Form, ByVal tmpOrderId As Long) As Boolean
    On Error GoTo ErrorHandler

    If shellForm Is Nothing Then
        Exit Function
    End If

    If tmpOrderId <= 0 Then
        Exit Function
    End If

    OpenOrderDetailNextWorkspaceByTempId = modAppWorkspaceService.OpenWorkspaceForm( _
        shellForm, _
        FORM_ORDER_DETAIL_NEXT, _
        vbNullString, _
        acFormPropertySettings, _
        BuildOrderEditWorkspaceArgs(tmpOrderId), _
        True)
    Exit Function

ErrorHandler:
    OpenOrderDetailNextWorkspaceByTempId = False
    modErrorHandler.HandleError MODULE_NAME, "OpenOrderDetailNextWorkspaceByTempId", Err
End Function

Private Function VerifyTemporaryOrderExists(ByVal tmpOrderId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If tmpOrderId <= 0 Then
        Exit Function
    End If

    Set db = modDb.GetFrontendDatabase()
    modLoggingHandler.LogInfo MODULE_NAME & ".VerifyTemporaryOrderExists", _
        "tmp_order_id=" & CStr(tmpOrderId) & "; frontend_db_name=" & SafeDatabaseName(db) & "."

    Set rs = db.OpenRecordset( _
        "SELECT [" & "tmp_order_id" & "], [" & "customer_address_id" & "], [" & "source_order_id" & "] FROM [tmp_order] " & _
        "WHERE [" & "tmp_order_id" & "]=" & CStr(tmpOrderId) & ";", _
        dbOpenSnapshot)

    If Not (rs.BOF And rs.EOF) Then
        VerifyTemporaryOrderExists = True
        modLoggingHandler.LogInfo MODULE_NAME & ".VerifyTemporaryOrderExists", _
            "Verification successful. tmp_order_id=" & CStr(tmpOrderId) & "; customer_address_id=" & CStr(modDaoHelper.NzLong(rs.Fields("customer_address_id").Value, 0)) & "; source_order_id=" & CStr(modDaoHelper.NzLong(rs.Fields("source_order_id").Value, 0)) & "."
    Else
        modLoggingHandler.LogWarning MODULE_NAME & ".VerifyTemporaryOrderExists", _
            "Verification failed. tmp_order_id=" & CStr(tmpOrderId) & " was not found in frontend database."
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    VerifyTemporaryOrderExists = False
    modErrorHandler.HandleError MODULE_NAME, "VerifyTemporaryOrderExists", Err
    Resume CleanExit
End Function

Private Function GetLoadedFormTmpOrderId(ByVal loadedForm As Access.Form) As Long
    On Error GoTo ErrorHandler

    If loadedForm Is Nothing Then
        Exit Function
    End If

    If modDaoHelper.RecordsetHasField(loadedForm.Recordset, "tmp_order_id") Then
        GetLoadedFormTmpOrderId = modDaoHelper.NzLong(loadedForm.Recordset.Fields("tmp_order_id").Value, 0)
    End If
    Exit Function

ErrorHandler:
    GetLoadedFormTmpOrderId = 0
End Function

Private Function SafeDatabaseName(ByVal db As DAO.Database) As String
    On Error GoTo SafeExit

    If db Is Nothing Then
        SafeDatabaseName = "<nothing>"
    Else
        SafeDatabaseName = Trim$(db.Name)
        If LenB(SafeDatabaseName) = 0 Then
            SafeDatabaseName = "<empty>"
        End If
    End If

SafeExit:
    If LenB(SafeDatabaseName) = 0 Then
        SafeDatabaseName = "<unavailable>"
    End If
End Function
