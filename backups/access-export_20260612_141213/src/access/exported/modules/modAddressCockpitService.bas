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
Private Const TABLE_ADR_ADDRESS As String = "adr_address"

Private Const FIELD_ADDRESS_ID As String = "address_id"
Private Const FIELD_COMPANY_NAME As String = "company_name"
Private Const FIELD_FIRST_NAME As String = "first_name"
Private Const FIELD_LAST_NAME As String = "last_name"
Private Const FIELD_LANGUAGE_CODE As String = "language_code"

Private Const DEFAULT_CURRENCY_CODE As String = "CHF"
Private Const DEFAULT_LANGUAGE_CODE As String = "DE-CH"
Private Const DEFAULT_VAT_MODE As String = "EXCLUSIVE"

Public Function GetAddressCockpitRecordSource() As String
    GetAddressCockpitRecordSource = "SELECT * FROM [" & QUERY_ADDRESS_COCKPIT_SUMMARY & "]"
End Function

Public Function CreateDraftSalesOrderForAddress(ByVal addressId As Long) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim displayName As String
    Dim languageCode As String

    CreateDraftSalesOrderForAddress = 0

    If addressId <= 0 Then
        Exit Function
    End If

    If Not modOrderRepository.EnsureOrderRepositoryReady() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset( _
        "SELECT * FROM [" & TABLE_ADR_ADDRESS & "] WHERE [" & FIELD_ADDRESS_ID & "]=" & CStr(addressId) & ";", _
        dbOpenSnapshot)

    If rs.BOF And rs.EOF Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CreateDraftSalesOrderForAddress", _
            "Address not found for address_id=" & CStr(addressId) & "."
        GoTo CleanExit
    End If

    displayName = ResolveDisplayName(rs)
    languageCode = ResolveLanguageCode(rs)

    CreateDraftSalesOrderForAddress = modOrderRepository.CreateSalesOrderHeader( _
        CustomerAddressId:=addressId, _
        OrderDate:=Date, _
        CustomerName:=displayName, _
        languageCode:=languageCode, _
        CurrencyCode:=DEFAULT_CURRENCY_CODE, _
        VatMode:=DEFAULT_VAT_MODE)

    modLoggingHandler.LogInfo MODULE_NAME & ".CreateDraftSalesOrderForAddress", _
        "Draft sales order created for address_id=" & CStr(addressId) & "; order_id=" & CStr(CreateDraftSalesOrderForAddress) & "."

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    CreateDraftSalesOrderForAddress = 0
    modErrorHandler.HandleError MODULE_NAME, "CreateDraftSalesOrderForAddress", Err
    Resume CleanExit
End Function

Private Function ResolveDisplayName(ByVal rs As DAO.Recordset) As String
    Dim companyName As String
    Dim personName As String

    companyName = modDaoHelper.NzString(rs.Fields(FIELD_COMPANY_NAME).Value)
    If LenB(Trim$(companyName)) > 0 Then
        ResolveDisplayName = Trim$(companyName)
        Exit Function
    End If

    personName = Trim$( _
        modDaoHelper.NzString(rs.Fields(FIELD_FIRST_NAME).Value) & " " & _
        modDaoHelper.NzString(rs.Fields(FIELD_LAST_NAME).Value))

    ResolveDisplayName = personName
End Function

Private Function ResolveLanguageCode(ByVal rs As DAO.Recordset) As String
    ResolveLanguageCode = Trim$(modDaoHelper.NzString(rs.Fields(FIELD_LANGUAGE_CODE).Value, DEFAULT_LANGUAGE_CODE))
    If LenB(ResolveLanguageCode) = 0 Then
        ResolveLanguageCode = DEFAULT_LANGUAGE_CODE
    End If
End Function
