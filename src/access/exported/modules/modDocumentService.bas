Option Compare Database
Option Explicit

'===============================================================================
' Module    : modDocumentService
' Purpose   : Business helpers for document line calculations and validation.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modDocumentService"

Private Const DOC_TYPE_OFFER As String = "OFFER"
Private Const DOC_TYPE_ORDER_EX As String = "ORDER"
Private Const DOC_TYPE_DELIVERY As String = "DELIVERY"
Private Const DOC_TYPE_INVOICE_EX As String = "INVOICE"
Private Const DOC_TYPE_RECEIPT_EX As String = "RECEIPT"
Private Const DOC_TYPE_PROFORMA As String = "PROFORMA"

Public Function CalculateDocumentLineNet(ByVal Quantity As Double, ByVal UnitPrice As Currency) As Currency
    On Error GoTo ErrorHandler

    CalculateDocumentLineNet = RoundCurrency(CCur(Quantity * CDbl(UnitPrice)))
    Exit Function

ErrorHandler:
    CalculateDocumentLineNet = 0
    modErrorHandler.HandleError MODULE_NAME, "CalculateDocumentLineNet", Err
End Function

Public Function CalculateDocumentLineGross(ByVal Quantity As Double, ByVal UnitPrice As Currency, ByVal VatRate As Double, ByVal VatMode As String) As Currency
    On Error GoTo ErrorHandler

    Dim baseAmount As Currency
    Dim normalizedVatMode As String

    baseAmount = CalculateDocumentLineNet(Quantity, UnitPrice)
    normalizedVatMode = modVatHandler.NormalizeVatMode(VatMode)

    Select Case normalizedVatMode
        Case "INCLUSIVE"
            CalculateDocumentLineGross = baseAmount
        Case "EXCLUSIVE"
            CalculateDocumentLineGross = modVatHandler.CalculateGrossFromNet(baseAmount, VatRate)
        Case Else
            CalculateDocumentLineGross = baseAmount
    End Select
    Exit Function

ErrorHandler:
    CalculateDocumentLineGross = 0
    modErrorHandler.HandleError MODULE_NAME, "CalculateDocumentLineGross", Err
End Function

Public Function CalculateDocumentLineVat(ByVal Quantity As Double, ByVal UnitPrice As Currency, ByVal VatRate As Double, ByVal VatMode As String) As Currency
    On Error GoTo ErrorHandler

    Dim baseAmount As Currency

    baseAmount = CalculateDocumentLineNet(Quantity, UnitPrice)
    CalculateDocumentLineVat = modVatHandler.CalculateVatAmount(baseAmount, VatRate, VatMode)
    Exit Function

ErrorHandler:
    CalculateDocumentLineVat = 0
    modErrorHandler.HandleError MODULE_NAME, "CalculateDocumentLineVat", Err
End Function

Public Function ValidateDocumentType(ByVal DocumentTypeCode As String) As Boolean
    Dim normalizedType As String

    normalizedType = UCase$(Trim$(DocumentTypeCode))

    Select Case normalizedType
        Case DOC_TYPE_OFFER, DOC_TYPE_ORDER_EX, DOC_TYPE_DELIVERY, DOC_TYPE_INVOICE_EX, DOC_TYPE_RECEIPT_EX, DOC_TYPE_PROFORMA
            ValidateDocumentType = True
        Case Else
            ValidateDocumentType = False
            If LenB(normalizedType) > 0 Then
                modLoggingHandler.LogWarning MODULE_NAME & ".ValidateDocumentType", _
                    "Unsupported document type '" & normalizedType & "'."
            End If
    End Select
End Function

Public Function CalculateDocumentTotalsFromPositions(ByVal NetSum As Currency, ByVal VatSum As Currency, ByVal GrossSum As Currency) As Boolean
    On Error GoTo ErrorHandler

    If NetSum < 0 Or VatSum < 0 Or GrossSum < 0 Then
        CalculateDocumentTotalsFromPositions = False
    Else
        CalculateDocumentTotalsFromPositions = True
    End If
    Exit Function

ErrorHandler:
    CalculateDocumentTotalsFromPositions = False
    modErrorHandler.HandleError MODULE_NAME, "CalculateDocumentTotalsFromPositions", Err
End Function

Private Function RoundCurrency(ByVal Amount As Currency) As Currency
    RoundCurrency = CCur(Round(CDbl(Amount), 2))
End Function