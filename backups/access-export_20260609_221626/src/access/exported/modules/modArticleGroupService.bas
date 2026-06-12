Attribute VB_Name = "modArticleGroupService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modArticleGroupService
' Purpose   : Table setup, defaults, search, and validation helpers for
'             Article Group master data.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modArticleGroupService"

Public Const TABLE_ARTICLE_GROUP As String = "art_product_group"
Public Const FORM_ARTICLE_GROUP_LIST As String = "frmArticleGroupList"
Public Const FORM_ARTICLE_GROUP_DETAIL As String = "frmArticleGroupDetail"

Private Const FIELD_ARTICLE_GROUP_ID As String = "product_group_id"
Private Const FIELD_ARTICLE_GROUP_CODE As String = "product_group_code"
Private Const FIELD_ARTICLE_GROUP_NAME As String = "product_group_name"
Private Const FIELD_DESCRIPTION_TEXT As String = "description_text"
Private Const FIELD_IS_ACTIVE As String = "is_active"
Private Const FIELD_SORT_ORDER As String = "sort_order"
Private Const FIELD_CREATED_AT As String = "created_at"
Private Const FIELD_CREATED_BY As String = "created_by"
Private Const FIELD_UPDATED_AT As String = "updated_at"
Private Const FIELD_UPDATED_BY As String = "updated_by"
Private Const FIELD_SEARCH_TEXT As String = "article_group_search_text"

Public Function EnsureArticleGroupTable() As Boolean
    On Error GoTo ErrorHandler

    If Not TableExists(TABLE_ARTICLE_GROUP) Then
        ExecuteSql "CREATE TABLE " & TABLE_ARTICLE_GROUP & " (" & _
                   FIELD_ARTICLE_GROUP_ID & " AUTOINCREMENT CONSTRAINT pk_art_product_group PRIMARY KEY, " & _
                   FIELD_ARTICLE_GROUP_CODE & " TEXT(50), " & _
                   FIELD_ARTICLE_GROUP_NAME & " TEXT(150), " & _
                   FIELD_DESCRIPTION_TEXT & " LONGTEXT, " & _
                   FIELD_IS_ACTIVE & " YESNO, " & _
                   FIELD_SORT_ORDER & " LONG, " & _
                   FIELD_CREATED_AT & " DATETIME, " & _
                   FIELD_CREATED_BY & " TEXT(100), " & _
                   FIELD_UPDATED_AT & " DATETIME, " & _
                   FIELD_UPDATED_BY & " TEXT(100));"
    ElseIf IsLinkedTable(TABLE_ARTICLE_GROUP) Then
        EnsureArticleGroupTable = True
        Exit Function
    End If

    EnsureFieldExists TABLE_ARTICLE_GROUP, FIELD_ARTICLE_GROUP_CODE, "TEXT(50)"
    EnsureFieldExists TABLE_ARTICLE_GROUP, FIELD_ARTICLE_GROUP_NAME, "TEXT(150)"
    EnsureFieldExists TABLE_ARTICLE_GROUP, FIELD_DESCRIPTION_TEXT, "LONGTEXT"
    EnsureFieldExists TABLE_ARTICLE_GROUP, FIELD_IS_ACTIVE, "YESNO"
    EnsureFieldExists TABLE_ARTICLE_GROUP, FIELD_SORT_ORDER, "LONG"
    EnsureFieldExists TABLE_ARTICLE_GROUP, FIELD_CREATED_AT, "DATETIME"
    EnsureFieldExists TABLE_ARTICLE_GROUP, FIELD_CREATED_BY, "TEXT(100)"
    EnsureFieldExists TABLE_ARTICLE_GROUP, FIELD_UPDATED_AT, "DATETIME"
    EnsureFieldExists TABLE_ARTICLE_GROUP, FIELD_UPDATED_BY, "TEXT(100)"

    EnsureIndexExists TABLE_ARTICLE_GROUP, "ux_art_product_group_code", _
        "CREATE UNIQUE INDEX ux_art_product_group_code ON " & TABLE_ARTICLE_GROUP & " (" & FIELD_ARTICLE_GROUP_CODE & ");"
    EnsureIndexExists TABLE_ARTICLE_GROUP, "ix_art_product_group_sort_order", _
        "CREATE INDEX ix_art_product_group_sort_order ON " & TABLE_ARTICLE_GROUP & " (" & FIELD_SORT_ORDER & ");"
    EnsureIndexExists TABLE_ARTICLE_GROUP, "ix_art_product_group_is_active", _
        "CREATE INDEX ix_art_product_group_is_active ON " & TABLE_ARTICLE_GROUP & " (" & FIELD_IS_ACTIVE & ");"

    EnsureArticleGroupTable = True
    Exit Function

ErrorHandler:
    EnsureArticleGroupTable = False
    modErrorHandler.HandleError MODULE_NAME, "EnsureArticleGroupTable", Err
End Function

Public Function BuildArticleGroupListRowSource() As String
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "SELECT "
    sqlStatement = sqlStatement & FIELD_ARTICLE_GROUP_ID & ", "
    sqlStatement = sqlStatement & FIELD_ARTICLE_GROUP_CODE & ", "
    sqlStatement = sqlStatement & FIELD_ARTICLE_GROUP_NAME & ", "
    sqlStatement = sqlStatement & FIELD_DESCRIPTION_TEXT & ", "
    sqlStatement = sqlStatement & FIELD_IS_ACTIVE & ", "
    sqlStatement = sqlStatement & FIELD_SORT_ORDER & ", "
    sqlStatement = sqlStatement & "UCase(Nz(" & FIELD_ARTICLE_GROUP_CODE & ",'')) & ' ' & "
    sqlStatement = sqlStatement & "UCase(Nz(" & FIELD_ARTICLE_GROUP_NAME & ",'')) & ' ' & "
    sqlStatement = sqlStatement & "UCase(Left(Nz(" & FIELD_DESCRIPTION_TEXT & ",''),255)) AS " & FIELD_SEARCH_TEXT & " "
    sqlStatement = sqlStatement & "FROM " & TABLE_ARTICLE_GROUP & " "
    sqlStatement = sqlStatement & "ORDER BY Nz(" & FIELD_IS_ACTIVE & ",True) DESC, Nz(" & FIELD_SORT_ORDER & ",0), " & FIELD_ARTICLE_GROUP_NAME & ", " & FIELD_ARTICLE_GROUP_CODE & ";"

    BuildArticleGroupListRowSource = sqlStatement
End Function

Public Function BuildArticleGroupSearchFilter(ByVal searchText As String) As String
    searchText = EscapeLikeValue(UCase$(Trim$(searchText)))
    BuildArticleGroupSearchFilter = "UCase(Nz([" & FIELD_SEARCH_TEXT & "],'')) Like '*" & searchText & "*'"
End Function

Public Sub ApplyDefaultValues(ByVal formInstance As Access.Form)
    On Error GoTo ErrorHandler

    If formInstance Is Nothing Then
        Exit Sub
    End If

    If HasFormControl(formInstance, FIELD_IS_ACTIVE) Then
        If IsNull(formInstance.Controls(FIELD_IS_ACTIVE).Value) Then
            formInstance.Controls(FIELD_IS_ACTIVE).Value = True
        End If
    End If

    If HasFormControl(formInstance, FIELD_SORT_ORDER) Then
        If modDaoHelper.NzLong(formInstance.Controls(FIELD_SORT_ORDER).Value, 0) <= 0 Then
            formInstance.Controls(FIELD_SORT_ORDER).Value = ResolveNextSortOrder()
        End If
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyDefaultValues", Err
End Sub

Public Function ValidateArticleGroupForm(ByVal formInstance As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    Dim articleGroupId As Long
    Dim articleGroupCode As String
    Dim articleGroupName As String

    If formInstance Is Nothing Then
        Exit Function
    End If

    If Not modFwValidationRuntime.ValidateForm(formInstance) Then
        Exit Function
    End If

    articleGroupId = ResolveArticleGroupId(formInstance)
    articleGroupCode = ResolveFieldText(formInstance, FIELD_ARTICLE_GROUP_CODE)
    articleGroupName = ResolveFieldText(formInstance, FIELD_ARTICLE_GROUP_NAME)

    If LenB(articleGroupCode) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_GROUP_CODE_REQUIRED", "Artikelgruppen-Code ist erforderlich."), vbExclamation, MODULE_NAME
        Exit Function
    End If

    If LenB(articleGroupName) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_GROUP_NAME_REQUIRED", "Artikelgruppen-Name ist erforderlich."), vbExclamation, MODULE_NAME
        Exit Function
    End If

    If ArticleGroupCodeExists(articleGroupCode, articleGroupId) Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_GROUP_DUPLICATE_CODE", "Artikelgruppen-Code existiert bereits."), vbExclamation, MODULE_NAME
        Exit Function
    End If

    ValidateArticleGroupForm = True
    Exit Function

ErrorHandler:
    ValidateArticleGroupForm = False
    modErrorHandler.HandleError MODULE_NAME, "ValidateArticleGroupForm", Err
End Function

Public Sub PrepareArticleGroupForSave(ByVal formInstance As Access.Form)
    On Error GoTo ErrorHandler

    Dim productGroupCode As String
    Dim productGroupName As String
    Dim DescriptionText As String

    If formInstance Is Nothing Then
        Exit Sub
    End If

    productGroupCode = UCase$(ResolveFieldText(formInstance, FIELD_ARTICLE_GROUP_CODE))
    productGroupName = ResolveFieldText(formInstance, FIELD_ARTICLE_GROUP_NAME)
    DescriptionText = ResolveFieldText(formInstance, FIELD_DESCRIPTION_TEXT)

    If LenB(productGroupCode) > 0 Then
        SetFieldValueIfPresent formInstance, FIELD_ARTICLE_GROUP_CODE, productGroupCode
    End If

    If LenB(productGroupName) > 0 Then
        SetFieldValueIfPresent formInstance, FIELD_ARTICLE_GROUP_NAME, productGroupName
    End If

    If HasFormControl(formInstance, FIELD_DESCRIPTION_TEXT) Then
        SetFieldValueIfPresent formInstance, FIELD_DESCRIPTION_TEXT, DescriptionText
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "PrepareArticleGroupForSave", Err
End Sub

Public Function ResolveArticleGroupId(ByVal formInstance As Access.Form) As Long
    If formInstance Is Nothing Then
        Exit Function
    End If

    If HasFormControl(formInstance, FIELD_ARTICLE_GROUP_ID) Then
        ResolveArticleGroupId = modDaoHelper.NzLong(formInstance.Controls(FIELD_ARTICLE_GROUP_ID).Value, 0)
    ElseIf HasFormControl(formInstance, "txtArticleGroupId") Then
        ResolveArticleGroupId = modDaoHelper.NzLong(formInstance.Controls("txtArticleGroupId").Value, 0)
    End If
End Function

Public Function ResolveArticleGroupDisplayName(ByVal formInstance As Access.Form) As String
    Dim articleGroupCode As String
    Dim articleGroupName As String

    articleGroupCode = ResolveFieldText(formInstance, FIELD_ARTICLE_GROUP_CODE)
    articleGroupName = ResolveFieldText(formInstance, FIELD_ARTICLE_GROUP_NAME)

    If LenB(articleGroupName) > 0 Then
        If LenB(articleGroupCode) > 0 Then
            ResolveArticleGroupDisplayName = articleGroupCode & " - " & articleGroupName
        Else
            ResolveArticleGroupDisplayName = articleGroupName
        End If
    Else
        ResolveArticleGroupDisplayName = articleGroupCode
    End If
End Function

Public Function ArticleGroupCodeExists( _
    ByVal articleGroupCode As String, _
    Optional ByVal excludeArticleGroupId As Long = 0) As Boolean
    On Error GoTo ErrorHandler

    Dim criteria As String

    articleGroupCode = UCase$(Trim$(articleGroupCode))

    criteria = "UCase(Nz([" & FIELD_ARTICLE_GROUP_CODE & "],'')) = " & SqlText(articleGroupCode) & " "

    If excludeArticleGroupId > 0 Then
        criteria = criteria & "AND [" & FIELD_ARTICLE_GROUP_ID & "] <> " & CStr(excludeArticleGroupId)
    End If

    ArticleGroupCodeExists = (DCount("*", TABLE_ARTICLE_GROUP, criteria) > 0)
    Exit Function

ErrorHandler:
    ArticleGroupCodeExists = False
    modErrorHandler.HandleError MODULE_NAME, "ArticleGroupCodeExists", Err
End Function

Public Function ResolveNextSortOrder() As Long
    On Error GoTo ErrorHandler

    Dim maxSortOrder As Variant

    maxSortOrder = DMax(FIELD_SORT_ORDER, TABLE_ARTICLE_GROUP)
    If IsNull(maxSortOrder) Or IsEmpty(maxSortOrder) Then
        ResolveNextSortOrder = 10
    Else
        ResolveNextSortOrder = CLng(maxSortOrder) + 10
    End If
    Exit Function

ErrorHandler:
    ResolveNextSortOrder = 10
    modErrorHandler.HandleError MODULE_NAME, "ResolveNextSortOrder", Err
End Function

Private Function ResolveFieldText(ByVal formInstance As Access.Form, ByVal fieldName As String) As String
    On Error GoTo SafeExit

    If formInstance Is Nothing Then
        Exit Function
    End If

    If HasFormControl(formInstance, fieldName) Then
        ResolveFieldText = Trim$(modDaoHelper.NzString(formInstance.Controls(fieldName).Value))
        Exit Function
    End If

    If HasFormControl(formInstance, "txt" & ConvertFieldNameToPascal(fieldName)) Then
        ResolveFieldText = Trim$(modDaoHelper.NzString(formInstance.Controls("txt" & ConvertFieldNameToPascal(fieldName)).Value))
    End If

SafeExit:
End Function

Private Function ConvertFieldNameToPascal(ByVal fieldName As String) As String
    Dim parts() As String
    Dim partValue As Variant

    parts = Split(fieldName, "_")
    For Each partValue In parts
        If LenB(CStr(partValue)) > 0 Then
            ConvertFieldNameToPascal = ConvertFieldNameToPascal & UCase$(Left$(CStr(partValue), 1)) & LCase$(Mid$(CStr(partValue), 2))
        End If
    Next partValue
End Function

Private Sub SetFieldValueIfPresent(ByVal formInstance As Access.Form, ByVal fieldName As String, ByVal fieldValue As Variant)
    On Error GoTo SafeExit

    If formInstance Is Nothing Then
        Exit Sub
    End If

    If HasFormControl(formInstance, fieldName) Then
        formInstance.Controls(fieldName).Value = fieldValue
    End If

SafeExit:
End Sub

Private Function EscapeLikeValue(ByVal valueText As String) As String
    valueText = Replace(valueText, "'", "''")
    valueText = Replace(valueText, "[", "[[]")
    valueText = Replace(valueText, "*", "[*]")
    valueText = Replace(valueText, "?", "[?]")
    valueText = Replace(valueText, "#", "[#]")
    EscapeLikeValue = valueText
End Function

Private Sub EnsureFieldExists(ByVal tableName As String, ByVal fieldName As String, ByVal ddlType As String)
    If Not FieldExists(tableName, fieldName) Then
        ExecuteSql "ALTER TABLE " & tableName & " ADD COLUMN " & fieldName & " " & ddlType & ";"
    End If
End Sub

Private Sub EnsureIndexExists(ByVal tableName As String, ByVal indexName As String, ByVal createSql As String)
    If Not IndexExists(tableName, indexName) Then
        ExecuteSql createSql
    End If
End Sub

Private Function TableExists(ByVal tableName As String) As Boolean
    On Error GoTo SafeExit

    Dim db As DAO.Database
    Dim tableDefinition As DAO.tableDef

    Set db = currentDb
    For Each tableDefinition In db.TableDefs
        If StrComp(tableDefinition.Name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next tableDefinition

SafeExit:
    Set tableDefinition = Nothing
    Set db = Nothing
End Function

Private Function FieldExists(ByVal tableName As String, ByVal fieldName As String) As Boolean
    On Error GoTo SafeExit

    Dim db As DAO.Database
    Dim tableDefinition As DAO.tableDef
    Dim fieldDefinition As DAO.Field

    Set db = currentDb
    Set tableDefinition = db.TableDefs(tableName)
    For Each fieldDefinition In tableDefinition.Fields
        If StrComp(fieldDefinition.Name, fieldName, vbTextCompare) = 0 Then
            FieldExists = True
            Exit Function
        End If
    Next fieldDefinition

SafeExit:
    Set fieldDefinition = Nothing
    Set tableDefinition = Nothing
    Set db = Nothing
End Function

Private Function IndexExists(ByVal tableName As String, ByVal indexName As String) As Boolean
    On Error GoTo SafeExit

    Dim db As DAO.Database
    Dim tableDefinition As DAO.tableDef
    Dim indexDefinition As DAO.index

    Set db = currentDb
    Set tableDefinition = db.TableDefs(tableName)
    For Each indexDefinition In tableDefinition.Indexes
        If StrComp(indexDefinition.Name, indexName, vbTextCompare) = 0 Then
            IndexExists = True
            Exit Function
        End If
    Next indexDefinition

SafeExit:
    Set indexDefinition = Nothing
    Set tableDefinition = Nothing
    Set db = Nothing
End Function

Private Function IsLinkedTable(ByVal tableName As String) As Boolean
    On Error GoTo SafeExit

    Dim db As DAO.Database
    Dim tableDefinition As DAO.tableDef

    Set db = currentDb
    Set tableDefinition = db.TableDefs(tableName)
    IsLinkedTable = (LenB(Trim$(modDaoHelper.NzString(tableDefinition.Connect))) > 0)

SafeExit:
    Set tableDefinition = Nothing
    Set db = Nothing
End Function

Private Sub ExecuteSql(ByVal sqlStatement As String)
    currentDb.Execute sqlStatement, dbFailOnError
End Sub

Private Function HasFormControl(ByVal formInstance As Access.Form, ByVal ControlName As String) As Boolean
    On Error GoTo SafeExit

    Dim currentControl As Control

    If formInstance Is Nothing Then
        Exit Function
    End If

    For Each currentControl In formInstance.Controls
        If StrComp(currentControl.Name, ControlName, vbTextCompare) = 0 Then
            HasFormControl = True
            Exit Function
        End If
    Next currentControl

SafeExit:
End Function