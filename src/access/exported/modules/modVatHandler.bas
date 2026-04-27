Attribute VB_Name = "modVatHandler"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modVatHandler
' Purpose   : Central VAT mode and VAT rate handling for framework calculations.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

Private Const MODULE_NAME As String = "modVatHandler"

Private Const VAT_MODE_NONE As String = "NONE"
Private Const VAT_MODE_INCLUSIVE As String = "INCLUSIVE"
Private Const VAT_MODE_EXCLUSIVE As String = "EXCLUSIVE"

Private Const PARAM_VAT_MODE As String = "VAT_MODE"
Private Const PARAM_VAT_RATE As String = "VAT_RATE"

Public Function GetVatMode() As String
    On Error GoTo ErrorHandler

    Dim rawVatMode As String
    Dim normalizedVatMode As String

    rawVatMode = modTenantRepository.GetTenantParameter(PARAM_VAT_MODE, VAT_MODE_NONE)
    normalizedVatMode = NormalizeVatMode(rawVatMode)

    If Not IsValidVatMode(normalizedVatMode) Then
        modLoggingHandler.LogWarning MODULE_NAME & ".GetVatMode", _
            "Invalid VAT mode '" & rawVatMode & "' found. Falling back to '" & VAT_MODE_NONE & "'."
        normalizedVatMode = VAT_MODE_NONE
    End If

    GetVatMode = normalizedVatMode
    Exit Function

ErrorHandler:
    GetVatMode = VAT_MODE_NONE
    modErrorHandler.HandleError MODULE_NAME, "GetVatMode", Err
End Function

Public Function GetVatRate() As Double
    On Error GoTo ErrorHandler

    Dim rawVatRate As String

    rawVatRate = Trim$(modTenantRepository.GetTenantParameter(PARAM_VAT_RATE, "0"))

    If LenB(rawVatRate) = 0 Or Not IsNumeric(rawVatRate) Then
        If LenB(rawVatRate) > 0 Then
            modLoggingHandler.LogWarning MODULE_NAME & ".GetVatRate", _
                "Invalid VAT rate '" & rawVatRate & "' found. Falling back to 0."
        End If
        GetVatRate = 0#
    Else
        GetVatRate = CDbl(rawVatRate)
    End If

    Exit Function

ErrorHandler:
    GetVatRate = 0#
    modErrorHandler.HandleError MODULE_NAME, "GetVatRate", Err
End Function

Public Function NormalizeVatMode(ByVal VatMode As String) As String
    NormalizeVatMode = UCase$(Trim$(VatMode))
End Function

Public Function IsValidVatMode(ByVal VatMode As String) As Boolean
    Select Case NormalizeVatMode(VatMode)
        Case VAT_MODE_NONE, VAT_MODE_INCLUSIVE, VAT_MODE_EXCLUSIVE
            IsValidVatMode = True
        Case Else
            IsValidVatMode = False
    End Select
End Function

Public Function CalculateGrossFromNet(ByVal NetAmount As Currency, ByVal VatRate As Double) As Currency
    On Error GoTo ErrorHandler

    Dim vatFactor As Double

    vatFactor = GetVatFactor(VatRate)
    CalculateGrossFromNet = RoundCurrency(CCur(CDbl(NetAmount) * (1# + vatFactor)))

    Exit Function

ErrorHandler:
    CalculateGrossFromNet = 0
    modErrorHandler.HandleError MODULE_NAME, "CalculateGrossFromNet", Err
End Function

Public Function CalculateNetFromGross(ByVal GrossAmount As Currency, ByVal VatRate As Double) As Currency
    On Error GoTo ErrorHandler

    Dim vatFactor As Double

    vatFactor = GetVatFactor(VatRate)

    If vatFactor <= 0 Then
        CalculateNetFromGross = RoundCurrency(GrossAmount)
    Else
        CalculateNetFromGross = RoundCurrency(CCur(CDbl(GrossAmount) / (1# + vatFactor)))
    End If

    Exit Function

ErrorHandler:
    CalculateNetFromGross = 0
    modErrorHandler.HandleError MODULE_NAME, "CalculateNetFromGross", Err
End Function

Public Function CalculateVatAmount(ByVal baseAmount As Currency, ByVal VatRate As Double, ByVal VatMode As String) As Currency
    On Error GoTo ErrorHandler

    Dim normalizedMode As String
    Dim vatFactor As Double
    Dim NetAmount As Currency
    Dim GrossAmount As Currency

    normalizedMode = NormalizeVatMode(VatMode)
    vatFactor = GetVatFactor(VatRate)

    Select Case normalizedMode
        Case VAT_MODE_NONE
            CalculateVatAmount = 0

        Case VAT_MODE_INCLUSIVE
            GrossAmount = RoundCurrency(baseAmount)
            NetAmount = CalculateNetFromGross(GrossAmount, VatRate)
            CalculateVatAmount = RoundCurrency(GrossAmount - NetAmount)

        Case VAT_MODE_EXCLUSIVE
            NetAmount = RoundCurrency(baseAmount)
            CalculateVatAmount = RoundCurrency(CCur(CDbl(NetAmount) * vatFactor))

        Case Else
            CalculateVatAmount = 0
    End Select

    Exit Function

ErrorHandler:
    CalculateVatAmount = 0
    modErrorHandler.HandleError MODULE_NAME, "CalculateVatAmount", Err
End Function

Private Function GetVatFactor(ByVal VatRate As Double) As Double
    If VatRate <= 0 Then
        GetVatFactor = 0#
    Else
        GetVatFactor = VatRate / 100#
    End If
End Function

Private Function RoundCurrency(ByVal Amount As Currency) As Currency
    RoundCurrency = CCur(Round(CDbl(Amount), 2))
End Function

