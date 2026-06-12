Attribute VB_Name = "modFwTranslationTagGeneratorService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modFwTranslationTagGeneratorService
' Purpose   : Builds and maintains staged FORM translation-tag data for the
'             focused translation tag generator workspace.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modFwTranslationTagGeneratorService"
Private Const TABLE_TAG_GENERATOR As String = "tmp_fw_translation_tag_generator"

Public Const STATUS_HAS_KEY As String = "HAS_KEY"
Public Const STATUS_MISSING_KEY As String = "MISSING_KEY"
Public Const STATUS_NON_TRANSLATABLE As String = "NON_TRANSLATABLE"
Public Const STATUS_INVALID_TAG As String = "INVALID_TAG"

Private Const FIELD_SESSION_ID As String = "session_id"
Private Const FIELD_ROW_NO As String = "row_no"
Private Const FIELD_FORM_NAME As String = "form_name"
Private Const FIELD_CONTROL_NAME As String = "control_name"
Private Const FIELD_CONTROL_TYPE As String = "control_type"
Private Const FIELD_SOURCE_TEXT As String = "source_text"
Private Const FIELD_CURRENT_TAG As String = "current_tag"
Private Const FIELD_ORIGINAL_TAG As String = "original_tag"
Private Const FIELD_CURRENT_TRANSLATION_KEY As String = "current_translation_key"
Private Const FIELD_SUGGESTED_TRANSLATION_KEY As String = "suggested_translation_key"
Private Const FIELD_TAG_STATUS As String = "tag_status"
Private Const FIELD_IS_RELEVANT As String = "is_relevant"
Private Const FIELD_IS_VISIBLE As String = "is_visible"
Private Const FIELD_IS_DIRTY As String = "is_dirty"
Private Const FIELD_CREATED_AT As String = "created_at"

Public Sub EnsureTranslationTagGeneratorTable()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = currentDb

    If TableExists(db, TABLE_TAG_GENERATOR) Then
        Exit Sub
    End If

    db.Execute _
        "CREATE TABLE " & TABLE_TAG_GENERATOR & " (" & _
        FIELD_SESSION_ID & " TEXT(100), " & _
        FIELD_ROW_NO & " LONG, " & _
        FIELD_FORM_NAME & " TEXT(255), " & _
        FIELD_CONTROL_NAME & " TEXT(255), " & _
        FIELD_CONTROL_TYPE & " TEXT(100), " & _
        FIELD_SOURCE_TEXT & " LONGTEXT, " & _
        FIELD_CURRENT_TAG & " LONGTEXT, " & _
        FIELD_ORIGINAL_TAG & " LONGTEXT, " & _
        FIELD_CURRENT_TRANSLATION_KEY & " TEXT(255), " & _
        FIELD_SUGGESTED_TRANSLATION_KEY & " TEXT(255), " & _
        FIELD_TAG_STATUS & " TEXT(50), " & _
        FIELD_IS_RELEVANT & " YESNO, " & _
        FIELD_IS_VISIBLE & " YESNO, " & _
        FIELD_IS_DIRTY & " YESNO, " & _
        FIELD_CREATED_AT & " DATETIME" & _
        ");", dbFailOnError

    modLoggingHandler.LogInfo MODULE_NAME & ".EnsureTranslationTagGeneratorTable", _
        "Created tag generator work table '" & TABLE_TAG_GENERATOR & "'."
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "EnsureTranslationTagGeneratorTable", Err
End Sub

Public Function LoadFormControlRows( _
    ByVal sessionId As String, _
    ByVal FormName As String, _
    Optional ByVal includeHidden As Boolean = False) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim metadataItems As Collection
    Dim metadata As Variant
    Dim RowNo As Long
    Dim currentTag As String
    Dim currentTranslationKey As String
    Dim suggestedTranslationKey As String
    Dim sourceText As String
    Dim isRelevant As Boolean
    Dim hasTranslationMarker As Boolean

    sessionId = Trim$(modDaoHelper.NzString(sessionId))
    FormName = Trim$(modDaoHelper.NzString(FormName))

    If LenB(sessionId) = 0 Or LenB(FormName) = 0 Then
        Exit Function
    End If

    EnsureTranslationTagGeneratorTable
    ClearSession sessionId

    Set db = currentDb
    Set metadataItems = modFwComposerService.GetFormControlMetadata(FormName, includeHidden)
    Set rs = db.OpenRecordset(TABLE_TAG_GENERATOR, dbOpenDynaset)

    For Each metadata In metadataItems
        RowNo = RowNo + 1
        currentTag = modDaoHelper.NzString(metadata("current_tag"))
        currentTranslationKey = modDaoHelper.NzString(metadata("current_translation_key"))
        sourceText = ResolveStoredSourceText(modDaoHelper.NzString(metadata("source_text")), modDaoHelper.NzString(metadata("control_name")))
        suggestedTranslationKey = BuildTranslationKey(FormName, modDaoHelper.NzString(metadata("control_name")))
        isRelevant = IsRelevantControl(metadata)
        hasTranslationMarker = CBool(metadata("has_translation_marker"))

        rs.AddNew
        rs.Fields(FIELD_SESSION_ID).Value = sessionId
        rs.Fields(FIELD_ROW_NO).Value = RowNo
        rs.Fields(FIELD_FORM_NAME).Value = FormName
        rs.Fields(FIELD_CONTROL_NAME).Value = modDaoHelper.NzString(metadata("control_name"))
        rs.Fields(FIELD_CONTROL_TYPE).Value = modDaoHelper.NzString(metadata("control_type"))
        rs.Fields(FIELD_SOURCE_TEXT).Value = sourceText
        rs.Fields(FIELD_CURRENT_TAG).Value = currentTag
        rs.Fields(FIELD_ORIGINAL_TAG).Value = currentTag
        rs.Fields(FIELD_CURRENT_TRANSLATION_KEY).Value = NullIfEmpty(currentTranslationKey)
        rs.Fields(FIELD_SUGGESTED_TRANSLATION_KEY).Value = NullIfEmpty(ResolveSuggestedTranslationKey(isRelevant, suggestedTranslationKey))
        rs.Fields(FIELD_TAG_STATUS).Value = ResolveTagStatus(isRelevant, hasTranslationMarker, currentTranslationKey)
        rs.Fields(FIELD_IS_RELEVANT).Value = isRelevant
        rs.Fields(FIELD_IS_VISIBLE).Value = modDaoHelper.NzBoolean(metadata("is_visible"), True)
        rs.Fields(FIELD_IS_DIRTY).Value = False
        rs.Fields(FIELD_CREATED_AT).Value = Now
        rs.Update
    Next metadata

    LoadFormControlRows = RowNo

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set metadataItems = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "LoadFormControlRows", Err
    Resume CleanExit
End Function

Public Function RegenerateSuggestions(ByVal sessionId As String) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim isRelevant As Boolean
    Dim suggestedTranslationKey As String

    sessionId = Trim$(modDaoHelper.NzString(sessionId))
    If LenB(sessionId) = 0 Then
        Exit Function
    End If

    EnsureTranslationTagGeneratorTable
    Set db = currentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM " & TABLE_TAG_GENERATOR & _
        " WHERE " & FIELD_SESSION_ID & " = " & SqlText(sessionId) & _
        " ORDER BY " & FIELD_ROW_NO & ";", dbOpenDynaset)

    Do While Not rs.EOF
        isRelevant = modDaoHelper.NzBoolean(rs.Fields(FIELD_IS_RELEVANT).Value, False)
        suggestedTranslationKey = BuildTranslationKey( _
            modDaoHelper.NzString(rs.Fields(FIELD_FORM_NAME).Value), _
            modDaoHelper.NzString(rs.Fields(FIELD_CONTROL_NAME).Value))

        rs.Edit
        rs.Fields(FIELD_SUGGESTED_TRANSLATION_KEY).Value = NullIfEmpty(ResolveSuggestedTranslationKey(isRelevant, suggestedTranslationKey))
        rs.Update

        RegenerateSuggestions = RegenerateSuggestions + 1
        rs.MoveNext
    Loop

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RegenerateSuggestions", Err
    Resume CleanExit
End Function

Public Function SetMissingKeys(ByVal sessionId As String) As Long
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    sessionId = Trim$(modDaoHelper.NzString(sessionId))
    If LenB(sessionId) = 0 Then
        Exit Function
    End If

    EnsureTranslationTagGeneratorTable
    Set db = currentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM " & TABLE_TAG_GENERATOR & _
        " WHERE " & FIELD_SESSION_ID & " = " & SqlText(sessionId) & _
        " AND " & FIELD_TAG_STATUS & " = " & SqlText(STATUS_MISSING_KEY) & _
        " ORDER BY " & FIELD_ROW_NO & ";", dbOpenDynaset)

    Do While Not rs.EOF
        If ApplySuggestedKeyToRow(rs) Then
            SetMissingKeys = SetMissingKeys + 1
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
    modErrorHandler.HandleError MODULE_NAME, "SetMissingKeys", Err
    Resume CleanExit
End Function

Public Function SetKeyForControl(ByVal sessionId As String, ByVal ControlName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    sessionId = Trim$(modDaoHelper.NzString(sessionId))
    ControlName = Trim$(modDaoHelper.NzString(ControlName))

    If LenB(sessionId) = 0 Or LenB(ControlName) = 0 Then
        Exit Function
    End If

    EnsureTranslationTagGeneratorTable
    Set db = currentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM " & TABLE_TAG_GENERATOR & _
        " WHERE " & FIELD_SESSION_ID & " = " & SqlText(sessionId) & _
        " AND " & FIELD_CONTROL_NAME & " = " & SqlText(ControlName) & ";", dbOpenDynaset)

    If Not (rs.BOF And rs.EOF) Then
        SetKeyForControl = ApplySuggestedKeyToRow(rs)
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SetKeyForControl", Err
    Resume CleanExit
End Function

Public Function RemoveKeyFromControl(ByVal sessionId As String, ByVal ControlName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim updatedTag As String
    Dim isRelevant As Boolean

    sessionId = Trim$(modDaoHelper.NzString(sessionId))
    ControlName = Trim$(modDaoHelper.NzString(ControlName))

    If LenB(sessionId) = 0 Or LenB(ControlName) = 0 Then
        Exit Function
    End If

    EnsureTranslationTagGeneratorTable
    Set db = currentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM " & TABLE_TAG_GENERATOR & _
        " WHERE " & FIELD_SESSION_ID & " = " & SqlText(sessionId) & _
        " AND " & FIELD_CONTROL_NAME & " = " & SqlText(ControlName) & ";", dbOpenDynaset)

    If Not (rs.BOF And rs.EOF) Then
        updatedTag = modFwTranslationRuntime.RemoveTranslationKeyFromTag(modDaoHelper.NzString(rs.Fields(FIELD_CURRENT_TAG).Value))
        isRelevant = modDaoHelper.NzBoolean(rs.Fields(FIELD_IS_RELEVANT).Value, False)

        rs.Edit
        rs.Fields(FIELD_CURRENT_TAG).Value = NullIfEmpty(updatedTag)
        rs.Fields(FIELD_CURRENT_TRANSLATION_KEY).Value = Null
        rs.Fields(FIELD_TAG_STATUS).Value = ResolveTagStatus(isRelevant, False, vbNullString)
        rs.Fields(FIELD_IS_DIRTY).Value = (StrComp(updatedTag, modDaoHelper.NzString(rs.Fields(FIELD_ORIGINAL_TAG).Value), vbBinaryCompare) <> 0)
        rs.Update

        RemoveKeyFromControl = True
    End If

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "RemoveKeyFromControl", Err
    Resume CleanExit
End Function

Public Function SaveTagChanges( _
    ByVal sessionId As String, _
    ByVal FormName As String, _
    Optional ByRef updatedCount As Long = 0) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim controlTagMap As Object

    sessionId = Trim$(modDaoHelper.NzString(sessionId))
    FormName = Trim$(modDaoHelper.NzString(FormName))

    If LenB(sessionId) = 0 Or LenB(FormName) = 0 Then
        Exit Function
    End If

    EnsureTranslationTagGeneratorTable
    Set db = currentDb
    Set controlTagMap = CreateObject("Scripting.Dictionary")
    controlTagMap.CompareMode = vbTextCompare

    Set rs = db.OpenRecordset( _
        "SELECT * FROM " & TABLE_TAG_GENERATOR & _
        " WHERE " & FIELD_SESSION_ID & " = " & SqlText(sessionId) & _
        " AND " & FIELD_IS_DIRTY & " = True" & _
        " ORDER BY " & FIELD_ROW_NO & ";", dbOpenDynaset)

    Do While Not rs.EOF
        controlTagMap(modDaoHelper.NzString(rs.Fields(FIELD_CONTROL_NAME).Value)) = modDaoHelper.NzString(rs.Fields(FIELD_CURRENT_TAG).Value)
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    If controlTagMap.count = 0 Then
        SaveTagChanges = True
        GoTo CleanExit
    End If

    If Not modFwComposerService.SaveControlTagsToObject(modFwComposerService.OBJECT_TYPE_FORM, FormName, controlTagMap, updatedCount) Then
        Exit Function
    End If

    Set rs = db.OpenRecordset( _
        "SELECT * FROM " & TABLE_TAG_GENERATOR & _
        " WHERE " & FIELD_SESSION_ID & " = " & SqlText(sessionId) & _
        " AND " & FIELD_IS_DIRTY & " = True" & _
        " ORDER BY " & FIELD_ROW_NO & ";", dbOpenDynaset)

    Do While Not rs.EOF
        rs.Edit
        rs.Fields(FIELD_ORIGINAL_TAG).Value = modDaoHelper.NzString(rs.Fields(FIELD_CURRENT_TAG).Value)
        rs.Fields(FIELD_IS_DIRTY).Value = False
        rs.Update
        rs.MoveNext
    Loop

    SaveTagChanges = True

CleanExit:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set controlTagMap = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "SaveTagChanges", Err
    Resume CleanExit
End Function

Public Function HasPendingChanges(ByVal sessionId As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rowCount As Long

    sessionId = Trim$(modDaoHelper.NzString(sessionId))
    If LenB(sessionId) = 0 Then
        Exit Function
    End If

    Set db = currentDb
    If Not TableExists(db, TABLE_TAG_GENERATOR) Then
        Exit Function
    End If

    rowCount = modDaoHelper.NzLong(DCount("*", TABLE_TAG_GENERATOR, _
        FIELD_SESSION_ID & " = " & SqlText(sessionId) & _
        " AND " & FIELD_IS_DIRTY & " = True"), 0)

    HasPendingChanges = (rowCount > 0)
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "HasPendingChanges", Err
End Function

Public Sub ClearSession(ByVal sessionId As String)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database

    sessionId = Trim$(modDaoHelper.NzString(sessionId))
    If LenB(sessionId) = 0 Then
        Exit Sub
    End If

    Set db = currentDb
    If Not TableExists(db, TABLE_TAG_GENERATOR) Then
        Exit Sub
    End If

    db.Execute "DELETE FROM " & TABLE_TAG_GENERATOR & _
               " WHERE " & FIELD_SESSION_ID & " = " & SqlText(sessionId) & ";", dbFailOnError
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ClearSession", Err
End Sub

Private Function ApplySuggestedKeyToRow(ByVal rs As DAO.Recordset) As Boolean
    On Error GoTo ErrorHandler

    Dim currentTag As String
    Dim suggestedTranslationKey As String
    Dim updatedTag As String
    Dim isRelevant As Boolean

    If rs Is Nothing Then
        Exit Function
    End If

    isRelevant = modDaoHelper.NzBoolean(rs.Fields(FIELD_IS_RELEVANT).Value, False)
    suggestedTranslationKey = Trim$(modDaoHelper.NzString(rs.Fields(FIELD_SUGGESTED_TRANSLATION_KEY).Value))

    If Not isRelevant Or LenB(suggestedTranslationKey) = 0 Then
        Exit Function
    End If

    currentTag = modDaoHelper.NzString(rs.Fields(FIELD_CURRENT_TAG).Value)
    updatedTag = modFwTranslationRuntime.SetTranslationKeyInTag(currentTag, suggestedTranslationKey)

    rs.Edit
    rs.Fields(FIELD_CURRENT_TAG).Value = NullIfEmpty(updatedTag)
    rs.Fields(FIELD_CURRENT_TRANSLATION_KEY).Value = suggestedTranslationKey
    rs.Fields(FIELD_TAG_STATUS).Value = STATUS_HAS_KEY
    rs.Fields(FIELD_IS_DIRTY).Value = (StrComp(updatedTag, modDaoHelper.NzString(rs.Fields(FIELD_ORIGINAL_TAG).Value), vbBinaryCompare) <> 0)
    rs.Update

    ApplySuggestedKeyToRow = True
    Exit Function

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplySuggestedKeyToRow", Err
End Function

Private Function IsRelevantControl(ByVal metadata As Variant) As Boolean
    Dim controlTypeId As Long
    Dim CaptionValue As String
    Dim attachedLabelCaption As String

    controlTypeId = modDaoHelper.NzLong(metadata("control_type_id"), 0)
    CaptionValue = Trim$(modDaoHelper.NzString(metadata("caption_value")))
    attachedLabelCaption = Trim$(modDaoHelper.NzString(metadata("attached_label_caption")))

    Select Case controlTypeId
        Case acLabel, acCommandButton, acPage, acOptionButton, acCheckBox, acToggleButton
            IsRelevantControl = (LenB(CaptionValue) > 0 Or LenB(attachedLabelCaption) > 0)

        Case acComboBox, acListBox, acTextBox
            IsRelevantControl = (LenB(attachedLabelCaption) > 0 Or LenB(CaptionValue) > 0)

        Case Else
            IsRelevantControl = False
    End Select
End Function

Private Function ResolveTagStatus( _
    ByVal isRelevant As Boolean, _
    ByVal hasTranslationMarker As Boolean, _
    ByVal currentTranslationKey As String) As String

    currentTranslationKey = Trim$(modDaoHelper.NzString(currentTranslationKey))

    If Not isRelevant Then
        ResolveTagStatus = STATUS_NON_TRANSLATABLE
    ElseIf hasTranslationMarker And LenB(currentTranslationKey) = 0 Then
        ResolveTagStatus = STATUS_INVALID_TAG
    ElseIf LenB(currentTranslationKey) > 0 Then
        ResolveTagStatus = STATUS_HAS_KEY
    Else
        ResolveTagStatus = STATUS_MISSING_KEY
    End If
End Function

Private Function BuildTranslationKey(ByVal FormName As String, ByVal ControlName As String) As String
    FormName = Trim$(modDaoHelper.NzString(FormName))
    ControlName = Trim$(modDaoHelper.NzString(ControlName))

    If LenB(FormName) = 0 Or LenB(ControlName) = 0 Then
        Exit Function
    End If

    BuildTranslationKey = "FORM." & UCase$(FormName) & "." & UCase$(ControlName)
End Function

Private Function ResolveSuggestedTranslationKey(ByVal isRelevant As Boolean, ByVal suggestedTranslationKey As String) As String
    If isRelevant Then
        ResolveSuggestedTranslationKey = Trim$(modDaoHelper.NzString(suggestedTranslationKey))
    Else
        ResolveSuggestedTranslationKey = vbNullString
    End If
End Function

Private Function ResolveStoredSourceText(ByVal sourceText As String, ByVal ControlName As String) As String
    sourceText = Trim$(modDaoHelper.NzString(sourceText))

    If LenB(sourceText) = 0 Then
        ResolveStoredSourceText = Trim$(modDaoHelper.NzString(ControlName))
    Else
        ResolveStoredSourceText = sourceText
    End If
End Function

Private Function NullIfEmpty(ByVal valueText As String) As Variant
    valueText = modDaoHelper.NzString(valueText)

    If LenB(valueText) = 0 Then
        NullIfEmpty = Null
    Else
        NullIfEmpty = valueText
    End If
End Function

Private Function TableExists(ByVal db As DAO.Database, ByVal tableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim tdf As DAO.tableDef

    If db Is Nothing Then
        Exit Function
    End If

    For Each tdf In db.TableDefs
        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next tdf

    Exit Function

ErrorHandler:
    TableExists = False
End Function