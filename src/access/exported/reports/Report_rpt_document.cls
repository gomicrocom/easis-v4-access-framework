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

Private Function JoinNonEmpty(ByVal Separator As String, ParamArray Values() As Variant) As String
    Dim i As Long
    Dim valueText As String
    Dim result As String
    
    For i = LBound(Values) To UBound(Values)
        valueText = Trim$(Nz(Values(i), ""))
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
    Dim isDeliveryNote As Boolean
    
    senderLine = BuildTenantSenderLine()
    windowPosition = UCase$(Trim$(Nz(modTenantRepository.GetTenantParameter("ADDRESS_WINDOW_POSITION", "LEFT"), "LEFT")))
    
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

    ' --- Adresse ---
    hasDelivery = Nz(Me!has_delivery_address, False)
    isDeliveryNote = (UCase$(Nz(Me!document_type, "")) = "DELIVERY_NOTE")
    
    Me!txtAddressLeft = vbNullString
    Me!txtAddressRight = vbNullString
    
    Me!txtAddressLeft.Visible = False
    Me!txtAddressRight.Visible = False
    
    If windowPosition = "RIGHT" Then
        
        If hasDelivery Then
            
            If isDeliveryNote Then
                ' Fenster = Lieferadresse
                Me!txtAddressRight = BuildDeliveryAddressBlock()
                Me!txtAddressLeft = BuildBillingAddressBlock()
            Else
                ' Fenster = Rechnungsadresse
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
            
            If isDeliveryNote Then
                ' Fenster = Lieferadresse
                Me!txtAddressLeft = BuildDeliveryAddressBlock()
                Me!txtAddressRight = BuildBillingAddressBlock()
            Else
                ' Fenster = Rechnungsadresse
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
Private Function BuildBillingAddressBlock() As String
    Dim result As String
    
    AppendLine result, Nz(Me!billing_display_name, "")
    AppendLine result, Nz(Me!billing_street, "")
    AppendLine result, Nz(Me!billing_zip_city, "")
    
    BuildBillingAddressBlock = result
End Function

Private Sub AppendLine(ByRef text As String, ByVal line As String)
    line = Trim$(line)
    If Len(line) = 0 Then Exit Sub
    
    If Len(text) > 0 Then
        text = text & vbCrLf
    End If
    
    text = text & line
End Sub
Private Function BuildDeliveryAddressBlock() As String
    Dim result As String
    
    AppendLine result, Nz(Me!delivery_display_name, "")
    AppendLine result, Nz(Me!delivery_street, "")
    AppendLine result, Nz(Me!delivery_zip_city, "")
    
    BuildDeliveryAddressBlock = result
End Function
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
    Dim isNoVat As Boolean
    isNoVat = (Nz(Me!vat_mode, "") = "NONE")

    Me!subVatSummary.Visible = Not isNoVat
    Me!lblNoVat.Visible = isNoVat
End Sub


