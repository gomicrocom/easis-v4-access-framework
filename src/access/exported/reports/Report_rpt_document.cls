VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rpt_document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Function BuildTenantSenderLine() As String
    Dim tenantName As String
    Dim tenantStreet As String
    Dim tenantZip As String
    Dim tenantCity As String
    
    tenantName = Nz(modTenantRepository.GetTenantParameter("TENANT_NAME", ""), "")
    tenantStreet = Nz(modTenantRepository.GetTenantParameter("TENANT_STREET", ""), "")
    tenantZip = Nz(modTenantRepository.GetTenantParameter("TENANT_ZIP_CODE", ""), "")
    tenantCity = Nz(modTenantRepository.GetTenantParameter("TENANT_CITY", ""), "")
    
    BuildTenantSenderLine = JoinNonEmpty(" | ", tenantName, tenantStreet, Trim$(tenantZip & " " & tenantCity))
End Function

Private Function JoinNonEmpty(ByVal Separator As String, ParamArray values() As Variant) As String
    Dim i As Long
    Dim valueText As String
    Dim result As String
    
    For i = LBound(values) To UBound(values)
        valueText = Trim$(Nz(values(i), ""))
        If Len(valueText) > 0 Then
            If Len(result) > 0 Then
                result = result & Separator
            End If
            result = result & valueText
        End If
    Next i
    
    JoinNonEmpty = result
End Function

Private Sub Berichtskopf_Format(Cancel As Integer, FormatCount As Integer)
    Dim senderLine As String
    Dim windowPosition As String
    Dim hasDelivery As Boolean
    Dim isDeliveryNoteValue  As Boolean
    
    ApplyTranslations
    senderLine = BuildTenantSenderLine()
    windowPosition = UCase$(Trim$(Nz(modTenantRepository.GetTenantParameter("ADDRESS_WINDOW_POSITION", "LEFT"), "LEFT")))
    
    hasDelivery = Nz(Me!has_delivery_address, False)
    isDeliveryNoteValue = IsDeliveryNote()
    
    Me!txtSenderLeft = vbNullString
    Me!txtSenderRight = vbNullString
    Me!txtSenderLeft.Visible = False
    Me!txtSenderRight.Visible = False
    
    If windowPosition = "RIGHT" Then
        Me!txtSenderRight = senderLine
        Me!txtSenderRight.Visible = True
    Else
        Me!txtSenderLeft = senderLine
        Me!txtSenderLeft.Visible = True
    End If
    
    Me!txtAddressLeft = vbNullString
    Me!txtAddressRight = vbNullString
    Me!txtAddressLeft.Visible = False
    Me!txtAddressRight.Visible = False
    
    If windowPosition = "RIGHT" Then
        If hasDelivery Then
            If isDeliveryNoteValue Then
                Me!txtAddressRight = BuildDeliveryAddressBlock()
                Me!txtAddressLeft = BuildBillingAddressBlock()
            Else
                Me!txtAddressRight = BuildBillingAddressBlock()
                Me!txtAddressLeft = BuildDeliveryAddressBlock()
            End If
            
            Me!txtAddressRight.Visible = True
            Me!txtAddressLeft.Visible = True
        Else
            Me!txtAddressRight = BuildBillingAddressBlock()
            Me!txtAddressRight.Visible = True
        End If
    Else
        If hasDelivery Then
            If isDeliveryNoteValue Then
                Me!txtAddressLeft = BuildDeliveryAddressBlock()
                Me!txtAddressRight = BuildBillingAddressBlock()
            Else
                Me!txtAddressLeft = BuildBillingAddressBlock()
                Me!txtAddressRight = BuildDeliveryAddressBlock()
            End If
            
            Me!txtAddressLeft.Visible = True
            Me!txtAddressRight.Visible = True
        Else
            Me!txtAddressLeft = BuildBillingAddressBlock()
            Me!txtAddressLeft.Visible = True
        End If
    End If
End Sub


Private Sub ApplyTranslations()
    Dim LanguageCode As String
    Dim documentType As String

    LanguageCode = Nz(GetControlValue("txtLanguageCode", "DE-CH"), "DE-CH")
    documentType = UCase$(Nz(GetControlValue("txtDocumentType", ""), ""))

    modTranslationService.SetCurrentLanguage LanguageCode

    If HasControl("txtDocumentTitle") Then
        Me!txtDocumentTitle = modTranslationService.T("DOCUMENT." & documentType, documentType)
    End If

    SetCaptionIfPresent "lblDocumentNo", modTranslationService.T("REPORT.DOCUMENT_NO", "Nummer")
    SetCaptionIfPresent "lblDocumentDate", modTranslationService.T("REPORT.DATE", "Datum")
    SetCaptionIfPresent "lblDocumentType", modTranslationService.T("REPORT.DOCUMENT_TYPE", "Dokumenttyp")
    SetCaptionIfPresent "lblStatus", modTranslationService.T("REPORT.STATUS", "Status")
    SetCaptionIfPresent "lblVatSummary", modTranslationService.T("REPORT.VAT_SUMMARY", "MwSt.-Aufschlüsselung")

    SetCaptionIfPresent "lblPos", modTranslationService.T("REPORT.POS", "Pos.")
    SetCaptionIfPresent "lblDescription", modTranslationService.T("REPORT.DESCRIPTION", "Beschreibung")
    SetCaptionIfPresent "lblQuantity", modTranslationService.T("REPORT.QUANTITY", "Anz.")
    SetCaptionIfPresent "lblUnit", modTranslationService.T("REPORT.UNIT", "Einh.")
    SetCaptionIfPresent "lblUnitPrice", modTranslationService.T("REPORT.UNIT_PRICE", "Preis")
    SetCaptionIfPresent "lblLineDiscount", modTranslationService.T("REPORT.DISCOUNT", "Rabatt")
    SetCaptionIfPresent "lblLineSurcharge", modTranslationService.T("REPORT.SURCHARGE", "Zuschlag")
    SetCaptionIfPresent "lblLineNetAmount", modTranslationService.T("REPORT.NET_AMOUNT", "Netto")
    SetCaptionIfPresent "lblLineVatAmount", modTranslationService.T("REPORT.VAT", "MwSt.")
    SetCaptionIfPresent "lblLineGrossAmount", modTranslationService.T("REPORT.GROSS_AMOUNT", "Brutto")

    SetCaptionIfPresent "lblSubtotal", modTranslationService.T("REPORT.SUBTOTAL", "Zwischensumme")
    SetCaptionIfPresent "lblHeaderDiscount", modTranslationService.T("REPORT.HEADER_DISCOUNT", "Kopfrabatt")
    SetCaptionIfPresent "lblHeaderSurcharge", modTranslationService.T("REPORT.HEADER_SURCHARGE", "Kopfzuschlag")
    SetCaptionIfPresent "lblNetAmount", modTranslationService.T("REPORT.NET_AMOUNT", "Netto")
    SetCaptionIfPresent "lblVatAmount", modTranslationService.T("REPORT.VAT", "MwSt.")
    SetCaptionIfPresent "lblTotalGross", modTranslationService.T("REPORT.TOTAL", "Total")
    SetCaptionIfPresent "lblPaymentTerms", modTranslationService.T("REPORT.PAYMENT_TERMS", "Zahlungsbedingungen")
    SetCaptionIfPresent "lblNoVat", modTranslationService.T("REPORT.NO_VAT", "Keine MwSt.")
End Sub

Private Sub Detailbereich_Format(Cancel As Integer, FormatCount As Integer)
    
    Dim IsDeliveryNote As Boolean
    IsDeliveryNote = (UCase$(Nz(Me!txtDocumentType, "")) = "DELIVERY_NOTE")

    Me!txtUnitPrice.Visible = Not IsDeliveryNote
    Me!txtLineDiscount.Visible = Not IsDeliveryNote
    Me!txtLineSurcharge.Visible = Not IsDeliveryNote
    Me!txtLineNet.Visible = Not IsDeliveryNote
    Me!txtLineVat.Visible = Not IsDeliveryNote
    Me!txtLineGross.Visible = Not IsDeliveryNote

End Sub

Private Sub grpDocument_Id_footer_Format(Cancel As Integer, FormatCount As Integer)
    
    Dim showVat As Boolean
    Dim isDeliveryNoteValue  As Boolean

    isDeliveryNoteValue = IsDeliveryNote()
    showVat = ShouldShowVatBlock()

    Me!srpt_document_vat_summary.Visible = showVat And Not IsDeliveryNote

    Me!lblPaymentTerms.Visible = (Not isDeliveryNoteValue And Len(Trim$(Nz(Me!txtPaymentTermsRaw, ""))) > 0)
    Me!txtPaymentTerms.Visible = Me!lblPaymentTerms.Visible
    
    If Me!txtPaymentTerms.Visible Then
        Me!txtPaymentTerms = Nz(Me!txtPaymentTermsRaw, "")
    End If

    Me!lblSubtotal.Visible = Not IsDeliveryNote
    Me!txtSubtotal.Visible = Not IsDeliveryNote
    
    Me!lblHeaderDiscount.Visible = Not IsDeliveryNote
    Me!txtHeaderDiscount.Visible = Not IsDeliveryNote
    
    Me!lblHeaderSurcharge.Visible = Not IsDeliveryNote
    Me!txtHeaderSurcharge.Visible = Not IsDeliveryNote
    
    Me!lblNetAmount.Visible = Not IsDeliveryNote
    Me!txtNetAmount.Visible = Not IsDeliveryNote
    
    Me!lblVatAmount.Visible = Not IsDeliveryNote
    Me!txtVatAmount.Visible = Not IsDeliveryNote
    
    Me!lblTotalGross.Visible = Not IsDeliveryNote
    Me!txtTotalGross.Visible = Not IsDeliveryNote
    
    Me!linTotalTop.Visible = Not IsDeliveryNote
    
    If Not IsDeliveryNote Then
        Me!txtTotalGross = FormatCurrencyAmount( _
            Nz(Me!txtGrossAmount, 0), _
            Nz(Me!txtCurrencyCode, "CHF") _
        )
    End If

End Sub

Private Sub grpDocument_Id_header_Format(Cancel As Integer, FormatCount As Integer)
    
    Dim IsDeliveryNote As Boolean
    IsDeliveryNote = (UCase$(Nz(Me!txtDocumentType, "")) = "DELIVERY_NOTE")

    Me!lblUnitPrice.Visible = Not IsDeliveryNote
    Me!lblLineDiscount.Visible = Not IsDeliveryNote
    Me!lblLineSurcharge.Visible = Not IsDeliveryNote

End Sub

Private Sub Report_NoData(Cancel As Integer)
    MsgBox "Kein Dokument für den angegebenen Filter gefunden.", vbExclamation
    Cancel = True
End Sub

Private Function BuildBillingAddressBlock() As String
    Dim result As String
    
    AppendLine result, Nz(Me!billing_display_name, "")
    AppendLine result, Nz(Me!billing_street, "")
    AppendLine result, BuildZipCityLine( _
        Nz(Me!billing_zip_code, ""), _
        Nz(Me!billing_city, ""), _
        Nz(Me!billing_country_code, "") _
    )
    
    BuildBillingAddressBlock = result
End Function

Private Function BuildDeliveryAddressBlock() As String
    Dim result As String
    
    AppendLine result, Nz(Me!delivery_display_name, "")
    AppendLine result, Nz(Me!delivery_street, "")
    AppendLine result, BuildZipCityLine( _
        Nz(Me!delivery_zip_code, ""), _
        Nz(Me!delivery_city, ""), _
        Nz(Me!delivery_country_code, "") _
    )
    
    BuildDeliveryAddressBlock = result
End Function

Private Function BuildZipCityLine( _
    ByVal ZipCode As String, _
    ByVal City As String, _
    ByVal CountryCode As String _
) As String
    Dim normalizedCountryCode As String
    Dim Prefix As String
    
    normalizedCountryCode = UCase$(Trim$(Nz(CountryCode, "")))
    
    If Len(normalizedCountryCode) > 0 _
       And normalizedCountryCode <> TenantCountryCode() Then
        Prefix = normalizedCountryCode & "-"
    End If
    
    BuildZipCityLine = Trim$(Prefix & Trim$(Nz(ZipCode, "")) & " " & Trim$(Nz(City, "")))
End Function

Private Sub AppendLine(ByRef text As String, ByVal line As String)
    line = Trim$(line)
    If Len(line) = 0 Then Exit Sub
    
    If Len(text) > 0 Then
        text = text & vbCrLf
    End If
    
    text = text & line
End Sub

Private Function TenantCountryCode() As String
    TenantCountryCode = UCase$(Trim$(Nz(modTenantRepository.GetTenantParameter("SENDER_COUNTRY_CODE", "CH"), "CH")))
End Function

Private Function IsDeliveryNote() As Boolean
    IsDeliveryNote = (UCase$(Nz(Me!txtDocumentType, "")) = "DELIVERY_NOTE")
End Function

Private Function IsNoVatDocument() As Boolean
    IsNoVatDocument = _
        UCase$(Nz(Me!txtVatMode, "")) = "NONE" _
        Or UCase$(Nz(Me!txtVatMode, "")) = "EXPORT"
End Function

Private Function ShouldShowVatBlock() As Boolean
    ShouldShowVatBlock = _
        Not IsDeliveryNote() _
        And Not IsNoVatDocument() _
        And Nz(Me!txtVatAmount, 0) <> 0
End Function




Private Function HasControl(ByVal ControlName As String) As Boolean
    On Error GoTo SafeExit
    
    Dim ctl As Control
    
    Set ctl = Me.Controls(ControlName)
    HasControl = True
    Exit Function
    
SafeExit:
    HasControl = False
End Function

Private Sub SetCaptionIfPresent(ByVal ControlName As String, ByVal CaptionText As String)
    On Error GoTo SafeExit
    
    If HasControl(ControlName) Then
        Me.Controls(ControlName).Caption = CaptionText
    End If
    
SafeExit:
End Sub

Private Function GetControlValue(ByVal ControlName As String, Optional ByVal DefaultValue As Variant) As Variant
    On Error GoTo SafeExit
    
    If HasControl(ControlName) Then
        GetControlValue = Me.Controls(ControlName).Value
    Else
        GetControlValue = DefaultValue
    End If
    Exit Function
    
SafeExit:
    GetControlValue = DefaultValue
End Function



