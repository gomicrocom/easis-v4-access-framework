Attribute VB_Name = "modCurrencyFormatService"
Option Compare Database
Option Explicit

'===============================================================================
' Module    : modCurrencyFormatService
' Purpose   : Provides centralized currency formatting based on ISO 4217
'             reference data from sys_be.accdb.
' Author    : Codex
' Version   : 0.1.0
'===============================================================================

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

    Dim fmt As tCurrencyFormat
    Dim valueText As String
    
    fmt = GetCurrencyFormat(CurrencyCode)
    
    valueText = FormatNumberCustom(Amount, fmt)
    
    If fmt.SymbolPosition = "PREFIX" Then
        FormatCurrencyAmount = fmt.Symbol & " " & valueText
    Else
        FormatCurrencyAmount = valueText & " " & fmt.Symbol
    End If

End Function

Private Function GetCurrencyFormat(ByVal CurrencyCode As String) As tCurrencyFormat

    Dim rs As DAO.Recordset
    Dim sql As String
    
    sql = "SELECT Symbol, SymbolPosition, DecimalSeparator, ThousandSeparator, MinorUnit " & _
          "FROM refCurrencies WHERE CurrencyCode = '" & CurrencyCode & "'"
    
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rs.EOF Then
        GetCurrencyFormat.Symbol = Nz(rs!Symbol, CurrencyCode)
        GetCurrencyFormat.SymbolPosition = Nz(rs!SymbolPosition, "PREFIX")
        GetCurrencyFormat.DecimalSeparator = Nz(rs!DecimalSeparator, ".")
        GetCurrencyFormat.ThousandSeparator = Nz(rs!ThousandSeparator, ",")
        GetCurrencyFormat.MinorUnit = Nz(rs!MinorUnit, 2)
    Else
        ' Fallback
        GetCurrencyFormat.Symbol = CurrencyCode
        GetCurrencyFormat.SymbolPosition = "PREFIX"
        GetCurrencyFormat.DecimalSeparator = "."
        GetCurrencyFormat.ThousandSeparator = ","
        GetCurrencyFormat.MinorUnit = 2
    End If
    
    rs.Close

End Function

Private Function FormatNumberCustom( _
    ByVal Amount As Currency, _
    ByRef fmt As tCurrencyFormat _
) As String

    Dim raw As String
    Dim parts() As String
    Dim intPart As String
    Dim decPart As String
    
    raw = Format$(Amount, "0." & String(fmt.MinorUnit, "0"))
    
    parts = Split(raw, ".")
    intPart = parts(0)
    
    If UBound(parts) >= 1 Then
        decPart = parts(1)
    Else
        decPart = ""
    End If
    
    intPart = AddThousands(intPart, fmt.ThousandSeparator)
    
    If fmt.MinorUnit > 0 Then
        FormatNumberCustom = intPart & fmt.DecimalSeparator & decPart
    Else
        FormatNumberCustom = intPart
    End If

End Function

Private Function AddThousands(ByVal valueText As String, ByVal sep As String) As String

    Dim result As String
    Dim i As Long
    Dim count As Long
    
    For i = Len(valueText) To 1 Step -1
        result = Mid$(valueText, i, 1) & result
        count = count + 1
        
        If count = 3 And i > 1 Then
            result = sep & result
            count = 0
        End If
    Next i
    
    AddThousands = result

End Function
