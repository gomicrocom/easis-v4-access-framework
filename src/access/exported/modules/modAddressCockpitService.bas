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

Public Function GetAddressCockpitRecordSource() As String
    GetAddressCockpitRecordSource = "SELECT * FROM [" & QUERY_ADDRESS_COCKPIT_SUMMARY & "]"
End Function
