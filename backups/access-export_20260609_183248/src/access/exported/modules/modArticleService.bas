Attribute VB_Name = "modArticleService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modArticleService
' Purpose   : Row-source, lookup, defaults, and validation helpers for the
'             simple Article master workflow.
' Author    : Codex
' Version   : 0.3.0
'===============================================================================

Private Const MODULE_NAME As String = "modArticleService"

Public Const TABLE_ARTICLE As String = "art_article"
Public Const FORM_ARTICLE_LIST As String = "frmArticleList"
Public Const FORM_ARTICLE_DETAIL As String = "frmArticleDetail"

Private Const TABLE_PRODUCT_GROUP As String = "art_product_group"
Private Const TABLE_ARTICLE_TYPE As String = "ref_article_type_code"
Private Const TABLE_REF_UNIT As String = "ref_unit"
Private Const TABLE_REF_VAT As String = "ref_vat_code"

Private Const FIELD_ARTICLE_ID As String = "article_id"
Private Const FIELD_ARTICLE_NO As String = "article_no"
Private Const FIELD_ARTICLE_NAME As String = "article_name"
Private Const FIELD_PRODUCT_GROUP_ID As String = "product_group_id"
Private Const FIELD_PRODUCT_GROUP_NAME As String = "product_group_name"
Private Const FIELD_ARTICLE_TYPE_CODE As String = "article_type_code"
Private Const FIELD_ARTICLE_TYPE_NAME As String = "article_type_name"
Private Const FIELD_TRANSLATION_KEY As String = "translation_key"
Private Const FIELD_UNIT_CODE As String = "unit_code"
Private Const FIELD_VAT_CODE As String = "vat_code"
Private Const FIELD_PURCHASE_PRICE As String = "purchase_price"
Private Const FIELD_SALES_PRICE As String = "sales_price"
Private Const FIELD_BARCODE As String = "barcode"
Private Const FIELD_DESCRIPTION_TEXT As String = "description_text"
Private Const FIELD_IS_ACTIVE As String = "is_active"
Private Const FIELD_CREATED_AT As String = "created_at"
Private Const FIELD_CREATED_BY As String = "created_by"
Private Const FIELD_UPDATED_AT As String = "updated_at"
Private Const FIELD_UPDATED_BY As String = "updated_by"
Private Const FIELD_SEARCH_TEXT As String = "article_search_text"
Private Const DEFAULT_ARTICLE_TYPE_CODE As String = "PRODUCT"
Private Const ARTICLE_NO_PREFIX As String = "ART-"
Private Const ARTICLE_NO_FORMAT As String = "000000"

Public Function BuildArticleListRowSource() As String
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "SELECT "
    sqlStatement = sqlStatement & "a." & FIELD_ARTICLE_ID & ", "
    sqlStatement = sqlStatement & "a." & FIELD_ARTICLE_NO & ", "
    sqlStatement = sqlStatement & "a." & FIELD_ARTICLE_NAME & ", "
    sqlStatement = sqlStatement & "a." & FIELD_PRODUCT_GROUP_ID & ", "
    sqlStatement = sqlStatement & "pg." & FIELD_PRODUCT_GROUP_NAME & ", "
    sqlStatement = sqlStatement & "a." & FIELD_ARTICLE_TYPE_CODE & ", "
    sqlStatement = sqlStatement & "a." & FIELD_UNIT_CODE & ", "
    sqlStatement = sqlStatement & "ResolveText(Nz(u.translation_key,''), Nz(a." & FIELD_UNIT_CODE & ",'')) AS unit_display_name, "
    sqlStatement = sqlStatement & "a." & FIELD_VAT_CODE & ", "
    sqlStatement = sqlStatement & "ResolveText(Nz(v.translation_key,''), Nz(a." & FIELD_VAT_CODE & ",'')) AS vat_display_name, "
    sqlStatement = sqlStatement & "v.vat_rate, "
    sqlStatement = sqlStatement & "a." & FIELD_PURCHASE_PRICE & ", "
    sqlStatement = sqlStatement & "a." & FIELD_SALES_PRICE & ", "
    sqlStatement = sqlStatement & "a." & FIELD_BARCODE & ", "
    sqlStatement = sqlStatement & "a." & FIELD_DESCRIPTION_TEXT & ", "
    sqlStatement = sqlStatement & "a." & FIELD_IS_ACTIVE & ", "
    sqlStatement = sqlStatement & "UCase("
    sqlStatement = sqlStatement & "Nz(a." & FIELD_ARTICLE_NO & ",'') & ' ' & "
    sqlStatement = sqlStatement & "Nz(a." & FIELD_ARTICLE_NAME & ",'') & ' ' & "
    sqlStatement = sqlStatement & "Nz(a." & FIELD_BARCODE & ",'') & ' ' & "
    sqlStatement = sqlStatement & "Left(Nz(a." & FIELD_DESCRIPTION_TEXT & ",''),255) & ' ' & "
    sqlStatement = sqlStatement & "Nz(pg." & FIELD_PRODUCT_GROUP_NAME & ",'') & ' ' & "
    sqlStatement = sqlStatement & "Nz(ResolveText(Nz(u.translation_key,''), Nz(a." & FIELD_UNIT_CODE & ",'')),'') & ' ' & "
    sqlStatement = sqlStatement & "Nz(ResolveText(Nz(v.translation_key,''), Nz(a." & FIELD_VAT_CODE & ",'')),'') & ' ' & "
    sqlStatement = sqlStatement & "Nz(a." & FIELD_UNIT_CODE & ",'') & ' ' & "
    sqlStatement = sqlStatement & "Nz(a." & FIELD_VAT_CODE & ",'')"
    sqlStatement = sqlStatement & ") AS " & FIELD_SEARCH_TEXT & " "

    sqlStatement = sqlStatement & "FROM (("
    sqlStatement = sqlStatement & TABLE_ARTICLE & " AS a "
    sqlStatement = sqlStatement & "LEFT JOIN " & TABLE_PRODUCT_GROUP & " AS pg "
    sqlStatement = sqlStatement & "ON a." & FIELD_PRODUCT_GROUP_ID & " = pg." & FIELD_PRODUCT_GROUP_ID & ") "
    sqlStatement = sqlStatement & "LEFT JOIN ref_unit AS u "
    sqlStatement = sqlStatement & "ON a." & FIELD_UNIT_CODE & " = u.unit_code) "
    sqlStatement = sqlStatement & "LEFT JOIN ref_vat_code AS v "
    sqlStatement = sqlStatement & "ON a." & FIELD_VAT_CODE & " = v.vat_code "

    sqlStatement = sqlStatement & "ORDER BY "
    sqlStatement = sqlStatement & "Nz(a." & FIELD_IS_ACTIVE & ",True) DESC, "
    sqlStatement = sqlStatement & "a." & FIELD_ARTICLE_NAME & ", "
    sqlStatement = sqlStatement & "a." & FIELD_ARTICLE_NO & ";"

    BuildArticleListRowSource = sqlStatement
End Function

Public Function BuildArticleSearchFilter(ByVal searchText As String) As String
    searchText = EscapeLikeValue(UCase$(Trim$(searchText)))
    BuildArticleSearchFilter = "UCase(Nz([" & FIELD_SEARCH_TEXT & "],'')) Like '*" & searchText & "*'"
End Function

Public Function BuildProductGroupComboRowSource() As String
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "SELECT "
    sqlStatement = sqlStatement & FIELD_PRODUCT_GROUP_ID & ", "
    sqlStatement = sqlStatement & FIELD_PRODUCT_GROUP_NAME & " "
    sqlStatement = sqlStatement & "FROM " & TABLE_PRODUCT_GROUP & " "
    sqlStatement = sqlStatement & "ORDER BY Nz(is_active, True) DESC, Nz(sort_order, 0), " & FIELD_PRODUCT_GROUP_NAME & ";"

    BuildProductGroupComboRowSource = sqlStatement
End Function

Public Function BuildArticleTypeComboRowSource() As String
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "SELECT "
    sqlStatement = sqlStatement & FIELD_ARTICLE_TYPE_CODE & ", "
    sqlStatement = sqlStatement & "ResolveText(Nz(" & FIELD_TRANSLATION_KEY & ",''), Nz(" & FIELD_ARTICLE_TYPE_NAME & ", Nz(" & FIELD_ARTICLE_TYPE_CODE & ",''))) AS article_type_display_name "
    sqlStatement = sqlStatement & "FROM " & TABLE_ARTICLE_TYPE & " "
    sqlStatement = sqlStatement & "WHERE Nz(is_active, True) = True "
    sqlStatement = sqlStatement & "ORDER BY Nz(sort_order, 0), " & FIELD_ARTICLE_TYPE_CODE & ";"

    BuildArticleTypeComboRowSource = sqlStatement
End Function

Public Function BuildUnitComboRowSource() As String
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "SELECT "
    sqlStatement = sqlStatement & "unit_code, "
    sqlStatement = sqlStatement & "ResolveText(Nz(translation_key,''), Nz(unit_code,'')) AS unit_display_name "
    sqlStatement = sqlStatement & "FROM " & TABLE_REF_UNIT & " "
    sqlStatement = sqlStatement & "WHERE Nz(is_active, True) = True "
    sqlStatement = sqlStatement & "ORDER BY Nz(sort_order, 0), unit_code;"

    BuildUnitComboRowSource = sqlStatement
End Function

Public Function BuildVatComboRowSource() As String
    Dim sqlStatement As String

    sqlStatement = ""
    sqlStatement = sqlStatement & "SELECT "
    sqlStatement = sqlStatement & "vat_code, "
    sqlStatement = sqlStatement & "ResolveText(Nz(translation_key,''), Nz(vat_code,'')) "
    sqlStatement = sqlStatement & "& ' (' & Format(Nz(vat_rate,0), '0.00') & '%)' AS vat_display_name "
    sqlStatement = sqlStatement & "FROM " & TABLE_REF_VAT & " "
    sqlStatement = sqlStatement & "WHERE Nz(is_active, True) = True "
    sqlStatement = sqlStatement & "ORDER BY Nz(sort_order, 0), vat_code;"

    BuildVatComboRowSource = sqlStatement
End Function

Public Sub ConfigureArticleDetailCombos(ByVal formInstance As Access.Form)
    On Error GoTo ErrorHandler

    If formInstance Is Nothing Then
        Exit Sub
    End If

    ConfigureComboIfPresent formInstance, FIELD_PRODUCT_GROUP_ID, BuildProductGroupComboRowSource(), 2, "0cm;5cm"
    ConfigureComboIfPresent formInstance, FIELD_ARTICLE_TYPE_CODE, BuildArticleTypeComboRowSource(), 2, "0cm;5cm"
    ConfigureComboIfPresent formInstance, FIELD_UNIT_CODE, BuildUnitComboRowSource(), 2, "0cm;4.5cm"
    ConfigureComboIfPresent formInstance, FIELD_VAT_CODE, BuildVatComboRowSource(), 2, "0cm;5cm"
    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ConfigureArticleDetailCombos", Err
End Sub

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

    If HasFormControl(formInstance, FIELD_SALES_PRICE) Then
        If IsNull(formInstance.Controls(FIELD_SALES_PRICE).Value) Then
            formInstance.Controls(FIELD_SALES_PRICE).Value = 0
        End If
    End If

    If HasFormControl(formInstance, FIELD_ARTICLE_TYPE_CODE) Then
        If LenB(Trim$(modDaoHelper.NzString(formInstance.Controls(FIELD_ARTICLE_TYPE_CODE).Value))) = 0 Then
            formInstance.Controls(FIELD_ARTICLE_TYPE_CODE).Value = DEFAULT_ARTICLE_TYPE_CODE
        End If
    End If

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "ApplyDefaultValues", Err
End Sub

Public Function ValidateArticleForm(ByVal formInstance As Access.Form) As Boolean
    On Error GoTo ErrorHandler

    Dim articleId As Long
    Dim articleNo As String
    Dim articleName As String
    Dim productGroupId As Long
    Dim unitCode As String
    Dim vatCode As String
    Dim salesPriceValue As Variant
    Dim purchasePriceValue As Variant

    If formInstance Is Nothing Then
        Exit Function
    End If

    If Not modFwValidationRuntime.ValidateForm(formInstance) Then
        Exit Function
    End If

    articleId = ResolveArticleId(formInstance)
    articleNo = UCase$(ResolveFieldText(formInstance, FIELD_ARTICLE_NO))
    articleName = ResolveFieldText(formInstance, FIELD_ARTICLE_NAME)
    productGroupId = ResolveLongFieldValue(formInstance, FIELD_PRODUCT_GROUP_ID)
    unitCode = ResolveFieldText(formInstance, FIELD_UNIT_CODE)
    vatCode = ResolveFieldText(formInstance, FIELD_VAT_CODE)
    salesPriceValue = ResolveRawFieldValue(formInstance, FIELD_SALES_PRICE)
    purchasePriceValue = ResolveRawFieldValue(formInstance, FIELD_PURCHASE_PRICE)

    If LenB(articleName) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_NAME_REQUIRED", "Artikelname ist erforderlich."), vbExclamation, MODULE_NAME
        Exit Function
    End If

    If productGroupId <= 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_GROUP_REQUIRED", "Artikelgruppe ist erforderlich."), vbExclamation, MODULE_NAME
        Exit Function
    End If

    If LenB(unitCode) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_UNIT_REQUIRED", "Einheit ist erforderlich."), vbExclamation, MODULE_NAME
        Exit Function
    End If

    If LenB(vatCode) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_VAT_REQUIRED", "MWST-Code ist erforderlich."), vbExclamation, MODULE_NAME
        Exit Function
    End If

    If IsNull(salesPriceValue) Or IsEmpty(salesPriceValue) Or LenB(Trim$(modDaoHelper.NzString(salesPriceValue))) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_SALES_PRICE_REQUIRED", "Verkaufspreis ist erforderlich."), vbExclamation, MODULE_NAME
        Exit Function
    End If

    If Not IsNumeric(salesPriceValue) Or CDbl(salesPriceValue) < 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_SALES_PRICE_REQUIRED", "Verkaufspreis ist erforderlich."), vbExclamation, MODULE_NAME
        Exit Function
    End If

    If Not IsNull(purchasePriceValue) And Not IsEmpty(purchasePriceValue) And LenB(Trim$(modDaoHelper.NzString(purchasePriceValue))) > 0 Then
        If Not IsNumeric(purchasePriceValue) Or CDbl(purchasePriceValue) < 0 Then
            MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_SAVE_ERROR", "Fehler beim Speichern des Artikels."), vbExclamation, MODULE_NAME
            Exit Function
        End If
    End If

    If LenB(articleNo) = 0 Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_SAVE_ERROR", "Fehler beim Speichern des Artikels."), vbExclamation, MODULE_NAME
        Exit Function
    End If

    If ArticleNoExists(articleNo, articleId) Then
        MsgBox modFwTranslationRuntime.ResolveText("MSG.ARTICLE_DUPLICATE_NO", "Artikel-Nr. existiert bereits."), vbExclamation, MODULE_NAME
        Exit Function
    End If

    ValidateArticleForm = True
    Exit Function

ErrorHandler:
    ValidateArticleForm = False
    modErrorHandler.HandleError MODULE_NAME, "ValidateArticleForm", Err
End Function

Public Sub PrepareArticleForSave(ByVal formInstance As Access.Form)
    On Error GoTo ErrorHandler

    Dim articleNo As String
    Dim articleName As String
    Dim articleTypeCode As String
    Dim unitCode As String
    Dim vatCode As String
    Dim barcode As String
    Dim DescriptionText As String

    If formInstance Is Nothing Then
        Exit Sub
    End If

    articleNo = UCase$(ResolveFieldText(formInstance, FIELD_ARTICLE_NO))
    articleName = ResolveFieldText(formInstance, FIELD_ARTICLE_NAME)
    articleTypeCode = UCase$(ResolveFieldText(formInstance, FIELD_ARTICLE_TYPE_CODE))
    unitCode = UCase$(ResolveFieldText(formInstance, FIELD_UNIT_CODE))
    vatCode = UCase$(ResolveFieldText(formInstance, FIELD_VAT_CODE))
    barcode = ResolveFieldText(formInstance, FIELD_BARCODE)
    DescriptionText = ResolveFieldText(formInstance, FIELD_DESCRIPTION_TEXT)

    If LenB(articleNo) = 0 Then
        articleNo = GenerateNextArticleNo()
    End If

    If LenB(articleTypeCode) = 0 Then
        articleTypeCode = DEFAULT_ARTICLE_TYPE_CODE
    End If

    SetFieldValueIfPresent formInstance, FIELD_ARTICLE_NO, articleNo
    SetFieldValueIfPresent formInstance, FIELD_ARTICLE_NAME, articleName
    SetFieldValueIfPresent formInstance, FIELD_ARTICLE_TYPE_CODE, articleTypeCode
    SetFieldValueIfPresent formInstance, FIELD_UNIT_CODE, unitCode
    SetFieldValueIfPresent formInstance, FIELD_VAT_CODE, vatCode
    SetFieldValueIfPresent formInstance, FIELD_BARCODE, barcode
    SetFieldValueIfPresent formInstance, FIELD_DESCRIPTION_TEXT, DescriptionText

    Exit Sub

ErrorHandler:
    modErrorHandler.HandleError MODULE_NAME, "PrepareArticleForSave", Err
End Sub

Public Function ArticleNoExists(ByVal articleNo As String, Optional ByVal excludeArticleId As Long = 0) As Boolean
    On Error GoTo ErrorHandler

    Dim criteria As String

    articleNo = UCase$(Trim$(articleNo))
    criteria = "UCase(Nz([" & FIELD_ARTICLE_NO & "],'')) = " & SqlText(articleNo)

    If excludeArticleId > 0 Then
        criteria = criteria & " AND [" & FIELD_ARTICLE_ID & "] <> " & CStr(excludeArticleId)
    End If

    ArticleNoExists = (DCount("*", TABLE_ARTICLE, criteria) > 0)
    Exit Function

ErrorHandler:
    ArticleNoExists = False
    modErrorHandler.HandleError MODULE_NAME, "ArticleNoExists", Err
End Function

Public Function GenerateNextArticleNo() As String
    On Error GoTo ErrorHandler

    Dim nextNumber As Long
    Dim candidateNo As String

    nextNumber = modDaoHelper.NzLong(DMax(FIELD_ARTICLE_ID, TABLE_ARTICLE), 0) + 1
    If nextNumber <= 0 Then
        nextNumber = 1
    End If

    Do
        candidateNo = ARTICLE_NO_PREFIX & Format$(nextNumber, ARTICLE_NO_FORMAT)
        If Not ArticleNoExists(candidateNo) Then
            GenerateNextArticleNo = candidateNo
            Exit Function
        End If
        nextNumber = nextNumber + 1
    Loop

ErrorHandler:
    If LenB(GenerateNextArticleNo) = 0 Then
        GenerateNextArticleNo = ARTICLE_NO_PREFIX & Format$(1, ARTICLE_NO_FORMAT)
    End If
    modErrorHandler.HandleError MODULE_NAME, "GenerateNextArticleNo", Err
End Function

Public Function ResolveNextArticleNo() As String
    ResolveNextArticleNo = GenerateNextArticleNo()
End Function

Public Function ResolveArticleDisplayName(ByVal formInstance As Access.Form) As String
    Dim articleNo As String
    Dim articleName As String

    articleNo = ResolveFieldText(formInstance, FIELD_ARTICLE_NO)
    articleName = ResolveFieldText(formInstance, FIELD_ARTICLE_NAME)

    If LenB(articleName) > 0 And LenB(articleNo) > 0 Then
        ResolveArticleDisplayName = articleNo & " - " & articleName
    ElseIf LenB(articleName) > 0 Then
        ResolveArticleDisplayName = articleName
    Else
        ResolveArticleDisplayName = articleNo
    End If
End Function

Public Function ResolveArticleId(ByVal formInstance As Access.Form) As Long
    On Error GoTo SafeExit

    If formInstance Is Nothing Then
        Exit Function
    End If

    If HasFormControl(formInstance, FIELD_ARTICLE_ID) Then
        ResolveArticleId = modDaoHelper.NzLong(formInstance.Controls(FIELD_ARTICLE_ID).Value, 0)
    ElseIf HasFormControl(formInstance, "txtArticleId") Then
        ResolveArticleId = modDaoHelper.NzLong(formInstance.Controls("txtArticleId").Value, 0)
    End If

SafeExit:
End Function

Private Sub ConfigureComboIfPresent( _
    ByVal formInstance As Access.Form, _
    ByVal ControlName As String, _
    ByVal rowSource As String, _
    ByVal columnCount As Integer, _
    ByVal columnWidths As String)
    On Error GoTo SafeExit

    If Not HasFormControl(formInstance, ControlName) Then
        Exit Sub
    End If

    formInstance.Controls(ControlName).RowSourceType = "Table/Query"
    formInstance.Controls(ControlName).rowSource = rowSource
    formInstance.Controls(ControlName).BoundColumn = 1
    formInstance.Controls(ControlName).columnCount = columnCount
    formInstance.Controls(ControlName).columnWidths = columnWidths

SafeExit:
End Sub

Private Function ResolveFieldText(ByVal formInstance As Access.Form, ByVal fieldName As String) As String
    On Error GoTo SafeExit

    ResolveFieldText = Trim$(modDaoHelper.NzString(ResolveRawFieldValue(formInstance, fieldName)))

SafeExit:
End Function

Private Function ResolveLongFieldValue(ByVal formInstance As Access.Form, ByVal fieldName As String) As Long
    On Error GoTo SafeExit

    ResolveLongFieldValue = modDaoHelper.NzLong(ResolveRawFieldValue(formInstance, fieldName), 0)

SafeExit:
End Function

Private Function ResolveRawFieldValue(ByVal formInstance As Access.Form, ByVal fieldName As String) As Variant
    On Error GoTo SafeExit

    If formInstance Is Nothing Then
        Exit Function
    End If

    If HasFormControl(formInstance, fieldName) Then
        ResolveRawFieldValue = formInstance.Controls(fieldName).Value
        Exit Function
    End If

    If HasFormControl(formInstance, "txt" & ConvertFieldNameToPascal(fieldName)) Then
        ResolveRawFieldValue = formInstance.Controls("txt" & ConvertFieldNameToPascal(fieldName)).Value
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

Private Function EscapeLikeValue(ByVal valueText As String) As String
    valueText = Replace(valueText, "'", "''")
    valueText = Replace(valueText, "[", "[[]")
    valueText = Replace(valueText, "*", "[*]")
    valueText = Replace(valueText, "?", "[?]")
    valueText = Replace(valueText, "#", "[#]")
    EscapeLikeValue = valueText
End Function