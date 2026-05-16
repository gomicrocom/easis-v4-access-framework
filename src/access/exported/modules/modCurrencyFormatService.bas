Attribute VB_Name = "modCurrencyFormatService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modCurrencyFormatService
' Purpose   : Provides centralized currency formatting based on ISO 4217
'             reference data from sys_be.accdb.
' Author    : Codex
' Version   : 0.1.1
'===============================================================================

Private Const MODULE_NAME As String = "modCurrencyFormatService"

Private Type tCurrencyFormat
    Symbol As String
    SymbolPosition As String
    DecimalSeparator As String
    ThousandSeparator As String
    MinorUnit As Integer
End Type

Public Function FormatCurrencyAmount( _
    ByVal Amount As Currency, _
    ByVal CurrencyCode As String _
) As String
    On Error GoTo ErrorHandler

    Dim fmt As tCurrencyFormat
    Dim valueText As String

    CurrencyCode = NormalizeCurrencyCode(CurrencyCode)

    fmt = GetCurrencyFormat(CurrencyCode)
    valueText = FormatNumberCustom(Amount, fmt)

    If UCase$(fmt.SymbolPosition) = "SUFFIX" Then
        FormatCurrencyAmount = valueText & " " & fmt.Symbol
    Else
        FormatCurrencyAmount = fmt.Symbol & " " & valueText
    End If

    Exit Function

ErrorHandler:
    FormatCurrencyAmount = CStr(Amount)
    modLoggingHandler.LogError MODULE_NAME & ".FormatCurrencyAmount", _
        "Failed to format currency amount. CurrencyCode=" & CurrencyCode, Err.Number
    modErrorHandler.HandleError MODULE_NAME, "FormatCurrencyAmount", Err
End Function

Public Function NormalizeCurrencyCode(ByVal CurrencyCode As String) As String
    On Error GoTo ErrorHandler

    NormalizeCurrencyCode = UCase$(Trim$(Nz(CurrencyCode, vbNullString)))

    If Len(NormalizeCurrencyCode) = 0 Then
        NormalizeCurrencyCode = "CHF"
    End If

    Exit Function

ErrorHandler:
    NormalizeCurrencyCode = "CHF"
    modLoggingHandler.LogError MODULE_NAME & ".NormalizeCurrencyCode", _
        "Failed to normalize currency code.", Err.Number
    modErrorHandler.HandleError MODULE_NAME, "NormalizeCurrencyCode", Err
End Function

Private Function GetCurrencyFormat(ByVal CurrencyCode As String) As tCurrencyFormat
    On Error GoTo ErrorHandler

    Dim rs As DAO.Recordset
    Dim sql As String

    CurrencyCode = NormalizeCurrencyCode(CurrencyCode)

    sql = "SELECT Symbol, SymbolPosition, DecimalSeparator, ThousandSeparator, MinorUnit " & _
          "FROM ref_currency " & _
          "WHERE CurrencyCode = '" & Replace(CurrencyCode, "'", "''") & "'"

    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)

    If Not rs.EOF Then
        GetCurrencyFormat.Symbol = Nz(rs!Symbol, CurrencyCode)
        GetCurrencyFormat.SymbolPosition = Nz(rs!SymbolPosition, "PREFIX")
        GetCurrencyFormat.DecimalSeparator = Nz(rs!DecimalSeparator, ".")
        GetCurrencyFormat.ThousandSeparator = Nz(rs!ThousandSeparator, ",")
        GetCurrencyFormat.MinorUnit = CInt(Nz(rs!MinorUnit, 2))
    Else
        GetCurrencyFormat = GetCurrencyFormatFallback(CurrencyCode)
    End If

    rs.Close
    Set rs = Nothing

    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    On Error GoTo 0

    GetCurrencyFormat = GetCurrencyFormatFallback(CurrencyCode)

    modLoggingHandler.LogError MODULE_NAME & ".GetCurrencyFormat", _
        "Failed to read currency format. CurrencyCode=" & CurrencyCode, Err.Number
    modErrorHandler.HandleError MODULE_NAME, "GetCurrencyFormat", Err
End Function

Private Function GetCurrencyFormatFallback(ByVal CurrencyCode As String) As tCurrencyFormat
    On Error GoTo ErrorHandler

    CurrencyCode = NormalizeCurrencyCode(CurrencyCode)

    GetCurrencyFormatFallback.Symbol = CurrencyCode
    GetCurrencyFormatFallback.SymbolPosition = "PREFIX"
    GetCurrencyFormatFallback.DecimalSeparator = "."
    GetCurrencyFormatFallback.ThousandSeparator = ","
    GetCurrencyFormatFallback.MinorUnit = 2

    Exit Function

ErrorHandler:
    GetCurrencyFormatFallback.Symbol = "CHF"
    GetCurrencyFormatFallback.SymbolPosition = "PREFIX"
    GetCurrencyFormatFallback.DecimalSeparator = "."
    GetCurrencyFormatFallback.ThousandSeparator = ","
    GetCurrencyFormatFallback.MinorUnit = 2
End Function

Private Function FormatNumberCustom( _
    ByVal Amount As Currency, _
    ByRef fmt As tCurrencyFormat _
) As String
    On Error GoTo ErrorHandler

    Dim raw As String
    Dim parts() As String
    Dim intPart As String
    Dim decPart As String
    Dim signText As String
    Dim absAmount As Currency

    If fmt.MinorUnit < 0 Then
        fmt.MinorUnit = 2
    End If

    If Amount < 0 Then
        signText = "-"
        absAmount = Abs(Amount)
    Else
        signText = vbNullString
        absAmount = Amount
    End If

    raw = Format$(absAmount, "0." & String$(fmt.MinorUnit, "0"))

    parts = Split(raw, ".")
    intPart = parts(0)

    If UBound(parts) >= 1 Then
        decPart = parts(1)
    Else
        decPart = vbNullString
    End If

    intPart = AddThousands(intPart, fmt.ThousandSeparator)

    If fmt.MinorUnit > 0 Then
        FormatNumberCustom = signText & intPart & fmt.DecimalSeparator & decPart
    Else
        FormatNumberCustom = signText & intPart
    End If

    Exit Function

ErrorHandler:
    FormatNumberCustom = CStr(Amount)
    modLoggingHandler.LogError MODULE_NAME & ".FormatNumberCustom", _
        "Failed to format number.", Err.Number
    modErrorHandler.HandleError MODULE_NAME, "FormatNumberCustom", Err
End Function

Private Function AddThousands(ByVal valueText As String, ByVal sep As String) As String
    On Error GoTo ErrorHandler

    Dim result As String
    Dim i As Long
    Dim count As Long

    valueText = Trim$(Nz(valueText, vbNullString))

    If Len(valueText) <= 3 Then
        AddThousands = valueText
        Exit Function
    End If

    For i = Len(valueText) To 1 Step -1
        result = Mid$(valueText, i, 1) & result
        count = count + 1

        If count = 3 And i > 1 Then
            result = sep & result
            count = 0
        End If
    Next i

    AddThousands = result
    Exit Function

ErrorHandler:
    AddThousands = valueText
    modLoggingHandler.LogError MODULE_NAME & ".AddThousands", _
        "Failed to apply thousand separator.", Err.Number
    modErrorHandler.HandleError MODULE_NAME, "AddThousands", Err
End Function

