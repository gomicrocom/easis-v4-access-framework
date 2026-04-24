Option Compare Database
Option Explicit

'===============================================================================
' Module    : modAddressRepository
' Purpose   : DAO repository helpers for address master data.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modAddressRepository"
Private Const TABLE_ADR_ADDRESS As String = "adr_address"

Private Const FIELD_ADDRESS_ID As String = "address_id"
Private Const FIELD_ADDRESS_TYPE_CODE As String = "address_type_code"
Private Const FIELD_COMPANY_NAME As String = "company_name"
Private Const FIELD_FIRST_NAME As String = "first_name"
Private Const FIELD_LAST_NAME As String = "last_name"
Private Const FIELD_STREET As String = "street"
Private Const FIELD_HOUSE_NO As String = "house_no"
Private Const FIELD_ZIP_CODE As String = "zip_code"
Private Const FIELD_CITY As String = "city"
Private Const FIELD_COUNTRY_CODE As String = "country_code"
Private Const FIELD_LANGUAGE_CODE As String = "language_code"
Private Const FIELD_IS_ACTIVE As String = "is_active"
Private Const FIELD_CREATED_AT As String = "created_at"
Private Const FIELD_CREATED_BY As String = "created_by"

Public Function CreateAddress( _
    ByVal AddressTypeCode As String, _
    Optional ByVal companyName As String = "", _
    Optional ByVal firstName As String = "", _
    Optional ByVal lastName As String = "", _
    Optional ByVal Street As String = "", _
    Optional ByVal HouseNo As String = "", _
    Optional ByVal ZipCode As String = "", _
    Optional ByVal City As String = "", _
    Optional ByVal CountryCode As String = "", _
    Optional ByVal LanguageCode As String = "" _
) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If Not CanWriteAddresses() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset(TABLE_ADR_ADDRESS, dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    SetRecordsetValue rs, FIELD_ADDRESS_TYPE_CODE, UCase$(Trim$(AddressTypeCode))
    SetRecordsetValue rs, FIELD_COMPANY_NAME, Trim$(companyName)
    SetRecordsetValue rs, FIELD_FIRST_NAME, Trim$(firstName)
    SetRecordsetValue rs, FIELD_LAST_NAME, Trim$(lastName)
    SetRecordsetValue rs, FIELD_STREET, Trim$(Street)
    SetRecordsetValue rs, FIELD_HOUSE_NO, Trim$(HouseNo)
    SetRecordsetValue rs, FIELD_ZIP_CODE, Trim$(ZipCode)
    SetRecordsetValue rs, FIELD_CITY, Trim$(City)
    SetRecordsetValue rs, FIELD_COUNTRY_CODE, UCase$(Trim$(CountryCode))
    SetRecordsetValue rs, FIELD_LANGUAGE_CODE, Trim$(LanguageCode)
    SetRecordsetValue rs, FIELD_IS_ACTIVE, True
    SetRecordsetValue rs, FIELD_CREATED_AT, Now()
    SetRecordsetValue rs, FIELD_CREATED_BY, ResolveCreatedBy()
    rs.Update

    rs.Bookmark = rs.LastModified
    If modDaoHelper.RecordsetHasField(rs, FIELD_ADDRESS_ID) Then
        CreateAddress = modDaoHelper.NzLong(rs.Fields(FIELD_ADDRESS_ID).Value, 0)
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".CreateAddress", _
        "Address created. AddressId=" & CStr(CreateAddress) & "."

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    CreateAddress = 0
    modErrorHandler.HandleError MODULE_NAME, "CreateAddress", Err
    Resume CleanExit
End Function

Public Function AddressExists(ByVal addressId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If addressId <= 0 Then
        Exit Function
    End If

    If Not CanReadAddresses() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_ADR_ADDRESS & "] WHERE [" & FIELD_ADDRESS_ID & "]=" & CStr(addressId) & ";", dbOpenSnapshot)

    AddressExists = Not (rs.BOF And rs.EOF)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    AddressExists = False
    modErrorHandler.HandleError MODULE_NAME, "AddressExists", Err
    Resume CleanExit
End Function

Public Function GetAddressDisplayName(ByVal addressId As Long, Optional ByVal DefaultValue As String = "") As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim companyName As String
    Dim firstName As String
    Dim lastName As String

    GetAddressDisplayName = DefaultValue

    If addressId <= 0 Then
        Exit Function
    End If

    If Not CanReadAddresses() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_ADR_ADDRESS & "] WHERE [" & FIELD_ADDRESS_ID & "]=" & CStr(addressId) & ";", dbOpenSnapshot)

    If rs.BOF And rs.EOF Then
        Exit Function
    End If

    companyName = ResolveFieldValue(rs, FIELD_COMPANY_NAME, vbNullString)
    If LenB(Trim$(companyName)) > 0 Then
        GetAddressDisplayName = Trim$(companyName)
        GoTo CleanExit
    End If

    firstName = ResolveFieldValue(rs, FIELD_FIRST_NAME, vbNullString)
    lastName = ResolveFieldValue(rs, FIELD_LAST_NAME, vbNullString)
    GetAddressDisplayName = Trim$(firstName & " " & lastName)

    If LenB(GetAddressDisplayName) = 0 Then
        GetAddressDisplayName = DefaultValue
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    GetAddressDisplayName = DefaultValue
    modErrorHandler.HandleError MODULE_NAME, "GetAddressDisplayName", Err
    Resume CleanExit
End Function

Private Function CanReadAddresses() As Boolean
    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CanReadAddresses", _
            "Backend configuration is not ready for address lookup."
        Exit Function
    End If

    If Not TableExists(TABLE_ADR_ADDRESS) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CanReadAddresses", _
            "Table '" & TABLE_ADR_ADDRESS & "' is not available yet."
        Exit Function
    End If

    CanReadAddresses = True
End Function

Private Function CanWriteAddresses() As Boolean
    CanWriteAddresses = CanReadAddresses()
End Function

Private Function TableExists(ByVal TableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tdf As DAO.tableDef

    Set db = modDb.GetCurrentDatabase()
    For Each tdf In db.TableDefs
        If UCase$(Trim$(tdf.Name)) = UCase$(Trim$(TableName)) Then
            TableExists = True
            Exit For
        End If
    Next tdf

CleanExit:
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    TableExists = False
    modErrorHandler.HandleError MODULE_NAME, "TableExists", Err
    Resume CleanExit
End Function

Private Sub SetRecordsetValue(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal FieldValue As Variant)
    If modDaoHelper.RecordsetHasField(rs, fieldName) Then
        rs.Fields(fieldName).Value = FieldValue
    End If
End Sub

Private Function ResolveFieldValue(ByVal rs As DAO.Recordset, ByVal fieldName As String, ByVal DefaultValue As String) As String
    If modDaoHelper.RecordsetHasField(rs, fieldName) Then
        ResolveFieldValue = modDaoHelper.NzString(rs.Fields(fieldName).Value, DefaultValue)
    Else
        ResolveFieldValue = DefaultValue
    End If
End Function

Private Function ResolveCreatedBy() As String
    If IsSessionInitialized() Then
        ResolveCreatedBy = currentUserId
    Else
        ResolveCreatedBy = "SYSTEM"
    End If
End Function