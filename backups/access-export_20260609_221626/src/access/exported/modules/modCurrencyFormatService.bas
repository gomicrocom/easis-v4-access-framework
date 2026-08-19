Attribute VB_Name = "modCurrencyFormatService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modCurrencyFormatService
' Purpose   : Provides centralized currency formatting based on ISO 4217
'             reference data from sys_be.accdb.
' Author    : Codex
' Version   : 0.1.2
'===============================================================================

Private Const MODULE_NAME As String = "modCurrencyFormatService"
Private Const TABLE_REF_CURRENCY As String = "ref_currency"
Private Const FIELD_CURRENCY_CODE As String = "currency_code"
Private Const FIELD_MINOR_UNIT As String = "minor_unit"
Private Const FIELD_CURRENCY_SYMBOL As String = "currency_symbol"
Private Const FIELD_SYMBOL_POSITION As String = "symbol_position"
Private Const FIELD_DECIMAL_SEPARATOR As String = "decimal_separator"
Private Const FIELD_THOUSAND_SEPARATOR As String = "thousand_separator"

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
    Dim currentFieldName As String

    CurrencyCode = NormalizeCurrencyCode(CurrencyCode)

    sql = "SELECT " & _
          FIELD_CURRENCY_SYMBOL & ", " & _
          FIELD_SYMBOL_POSITION & ", " & _
          FIELD_DECIMAL_SEPARATOR & ", " & _
          FIELD_THOUSAND_SEPARATOR & ", " & _
          FIELD_MINOR_UNIT & " " & _
          "FROM " & TABLE_REF_CURRENCY & " " & _
          "WHERE " & FIELD_CURRENCY_CODE & " = '" & Replace(CurrencyCode, "'", "''") & "'"

    modLoggingHandler.LogInfo MODULE_NAME & ".GetCurrencyFormat", _
        "Query table=" & TABLE_REF_CURRENCY & "; sql=" & sql

    currentFieldName = "<openrecordset>"
    Set rs = currentDb.OpenRecordset(sql, dbOpenSnapshot)

    If Not rs.EOF Then
        currentFieldName = FIELD_CURRENCY_SYMBOL
        GetCurrencyFormat.Symbol = Nz(rs.Fields(FIELD_CURRENCY_SYMBOL).Value, CurrencyCode)

        currentFieldName = FIELD_SYMBOL_POSITION
        GetCurrencyFormat.SymbolPosition = Nz(rs.Fields(FIELD_SYMBOL_POSITION).Value, "PREFIX")

        currentFieldName = FIELD_DECIMAL_SEPARATOR
        GetCurrencyFormat.DecimalSeparator = Nz(rs.Fields(FIELD_DECIMAL_SEPARATOR).Value, ".")

        currentFieldName = FIELD_THOUSAND_SEPARATOR
        GetCurrencyFormat.ThousandSeparator = Nz(rs.Fields(FIELD_THOUSAND_SEPARATOR).Value, ",")

        currentFieldName = FIELD_MINOR_UNIT
        GetCurrencyFormat.MinorUnit = CInt(Nz(rs.Fields(FIELD_MINOR_UNIT).Value, 2))
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
        "Failed to read currency format. table=" & TABLE_REF_CURRENCY & _
        "; sql=" & sql & _
        "; field=" & currentFieldName & _
        "; currency_code=" & CurrencyCode, Err.Number
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