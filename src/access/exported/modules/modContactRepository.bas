Option Compare Database
Option Explicit

'===============================================================================
' Module    : modContactRepository
' Purpose   : DAO repository helpers for address contact data.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modContactRepository"
Private Const TABLE_ADR_CONTACT As String = "adr_contact"

Private Const FIELD_CONTACT_ID As String = "contact_id"
Private Const FIELD_ADDRESS_ID As String = "address_id"
Private Const FIELD_CONTACT_TYPE_CODE As String = "contact_type_code"
Private Const FIELD_CONTACT_VALUE As String = "contact_value"
Private Const FIELD_IS_PRIMARY As String = "is_primary"
Private Const FIELD_REMARKS As String = "remarks"
Private Const FIELD_CREATED_AT As String = "created_at"
Private Const FIELD_CREATED_BY As String = "created_by"

Public Function CreateContact( _
    ByVal addressId As Long, _
    ByVal ContactTypeCode As String, _
    ByVal ContactValue As String, _
    Optional ByVal IsPrimary As Boolean = False, _
    Optional ByVal Remarks As String = "" _
) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If addressId <= 0 Then
        Exit Function
    End If

    If Not CanWriteContacts() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset(TABLE_ADR_CONTACT, dbOpenDynaset, dbAppendOnly)

    rs.AddNew
    SetRecordsetValue rs, FIELD_ADDRESS_ID, addressId
    SetRecordsetValue rs, FIELD_CONTACT_TYPE_CODE, UCase$(Trim$(ContactTypeCode))
    SetRecordsetValue rs, FIELD_CONTACT_VALUE, Trim$(ContactValue)
    SetRecordsetValue rs, FIELD_IS_PRIMARY, IsPrimary
    SetRecordsetValue rs, FIELD_REMARKS, Trim$(Remarks)
    SetRecordsetValue rs, FIELD_CREATED_AT, Now()
    SetRecordsetValue rs, FIELD_CREATED_BY, ResolveCreatedBy()
    rs.Update

    rs.Bookmark = rs.LastModified
    If modDaoHelper.RecordsetHasField(rs, FIELD_CONTACT_ID) Then
        CreateContact = modDaoHelper.NzLong(rs.Fields(FIELD_CONTACT_ID).Value, 0)
    End If

    modLoggingHandler.LogInfo MODULE_NAME & ".CreateContact", _
        "Contact created. ContactId=" & CStr(CreateContact) & ", AddressId=" & CStr(addressId) & "."

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    CreateContact = 0
    modErrorHandler.HandleError MODULE_NAME, "CreateContact", Err
    Resume CleanExit
End Function

Public Function GetPrimaryContactValue(ByVal addressId As Long, ByVal ContactTypeCode As String, Optional ByVal DefaultValue As String = "") As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim targetType As String

    GetPrimaryContactValue = DefaultValue

    If addressId <= 0 Then
        Exit Function
    End If

    If Not CanReadContacts() Then
        Exit Function
    End If

    targetType = UCase$(Trim$(ContactTypeCode))
    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_ADR_CONTACT & "] WHERE [" & FIELD_ADDRESS_ID & "]=" & CStr(addressId) & ";", dbOpenSnapshot)

    If rs.BOF And rs.EOF Then
        Exit Function
    End If

    rs.MoveFirst
    Do Until rs.EOF
        If UCase$(Trim$(ResolveFieldValue(rs, FIELD_CONTACT_TYPE_CODE, vbNullString))) = targetType Then
            If modDaoHelper.RecordsetHasField(rs, FIELD_IS_PRIMARY) Then
                If modDaoHelper.NzBoolean(rs.Fields(FIELD_IS_PRIMARY).Value, False) Then
                    GetPrimaryContactValue = ResolveFieldValue(rs, FIELD_CONTACT_VALUE, DefaultValue)
                    Exit Do
                End If
            End If
        End If
        rs.MoveNext
    Loop

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    GetPrimaryContactValue = DefaultValue
    modErrorHandler.HandleError MODULE_NAME, "GetPrimaryContactValue", Err
    Resume CleanExit
End Function

Public Function ContactExists(ByVal ContactId As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    If ContactId <= 0 Then
        Exit Function
    End If

    If Not CanReadContacts() Then
        Exit Function
    End If

    Set db = modDb.GetCurrentDatabase()
    Set rs = db.OpenRecordset("SELECT * FROM [" & TABLE_ADR_CONTACT & "] WHERE [" & FIELD_CONTACT_ID & "]=" & CStr(ContactId) & ";", dbOpenSnapshot)

    ContactExists = Not (rs.BOF And rs.EOF)

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    ContactExists = False
    modErrorHandler.HandleError MODULE_NAME, "ContactExists", Err
    Resume CleanExit
End Function

Private Function CanReadContacts() As Boolean
    If Not modDb.ValidateBackendConfiguration() Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CanReadContacts", _
            "Backend configuration is not ready for contact lookup."
        Exit Function
    End If

    If Not TableExists(TABLE_ADR_CONTACT) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".CanReadContacts", _
            "Table '" & TABLE_ADR_CONTACT & "' is not available yet."
        Exit Function
    End If

    CanReadContacts = True
End Function

Private Function CanWriteContacts() As Boolean
    CanWriteContacts = CanReadContacts()
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