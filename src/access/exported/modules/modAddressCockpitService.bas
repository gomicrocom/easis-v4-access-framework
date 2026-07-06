Attribute VB_Name = "modAddressCockpitService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modAddressCockpitService
' Purpose   : Provides summary and action helpers for frmAddressCockpit.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modAddressCockpitService"

Private Const QUERY_ADDRESS_COCKPIT_SUMMARY As String = "qry_address_cockpit_summary"
Private Const FIELD_ADDRESS_ID As String = "address_id"
Private Const TMP_ORDER_OPEN_ARGS_PREFIX As String = "TMP_ORDER"

Public Function GetAddressCockpitRecordSource() As String
    GetAddressCockpitRecordSource = "SELECT * FROM [" & QUERY_ADDRESS_COCKPIT_SUMMARY & "]"
End Function

Public Function CreateTemporarySalesOrderWorkspaceArgs(ByVal addressId As Long) As String
    On Error GoTo ErrorHandler

    Dim tmpOrderId As Long

    If addressId <= 0 Then
        Exit Function
    End If

    If Not modAddressRepository.AddressExists(addressId) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CreateTemporarySalesOrderWorkspaceArgs", _
            "Address not found for address_id=" & CStr(addressId) & "."
        Exit Function
    End If

    tmpOrderId = modOrderRepository.CreateTemporarySalesOrderForAddress(addressId)
    If tmpOrderId <= 0 Then
        Exit Function
    End If

    CreateTemporarySalesOrderWorkspaceArgs = TMP_ORDER_OPEN_ARGS_PREFIX & ";" & CStr(tmpOrderId)
    Exit Function

ErrorHandler:
    CreateTemporarySalesOrderWorkspaceArgs = vbNullString
    modErrorHandler.HandleError MODULE_NAME, "CreateTemporarySalesOrderWorkspaceArgs", Err
End Function
