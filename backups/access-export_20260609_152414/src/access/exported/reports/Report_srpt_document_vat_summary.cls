Option Compare Database
Option Explicit

Private Sub Report_Open(Cancel As Integer)
    ApplyTranslations
End Sub

Private Sub ApplyTranslations()
    Me!lblVatRate.Caption = modTranslationService.T("REPORT.VAT_RATE", "Satz")
    Me!lblPositionCount.Caption = modTranslationService.T("REPORT.POSITION_COUNT", "Anz.")
    Me!lblBaseAmount.Caption = modTranslationService.T("REPORT.VAT_BASE", "Basis")
    Me!lblVatAmount.Caption = modTranslationService.T("REPORT.VAT_AMOUNT", "MwSt.")
End Sub