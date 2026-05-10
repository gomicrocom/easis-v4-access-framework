VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_srpt_document_positions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Berichtskopf_Format(Cancel As Integer, FormatCount As Integer)

    Dim IsDeliveryNote As Boolean
    IsDeliveryNote = (UCase$(Nz(Me!txtDocumentType, "")) = "DELIVERY_NOTE")

    Me!lblUnitPrice.Visible = Not IsDeliveryNote
    Me!lblLineDiscount.Visible = Not IsDeliveryNote
    Me!lblLineSurcharge.Visible = Not IsDeliveryNote
    Me!lblLineNet.Visible = Not IsDeliveryNote
    Me!lblLineVat.Visible = Not IsDeliveryNote
    Me!lblLineGross.Visible = Not IsDeliveryNote

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
