Attribute VB_Name = "modExportHelper"
Option Compare Database
Option Explicit

Sub ExportAllModules()
    Dim obj As AccessObject
    Dim path As String
    
    ' Ordnerpfad festlegen, wohin exportiert werden soll
    path = "C:\Users\gomic\OneDrive\Documents\GitHub\easis-v4-access-framework\src\access\exported\modules\"
    
    ' Ordner erstellen, falls nicht vorhanden
    If Dir(path, vbDirectory) = "" Then MkDir path
    
    ' Standardmodule
    For Each obj In CurrentProject.AllModules
        If Not obj.Name = "modExportHelper" Then Application.SaveAsText acModule, obj.Name, path & obj.Name & ".bas"
    Next obj
    
    ' Klassenmodule
    For Each obj In CurrentProject.AllModules
        ' Klassenmodule haben oft den Typ acClassModule
        ' Application.SaveAsText acClassModule, obj.Name, path & obj.Name & ".cls"
    Next obj
    
    MsgBox "Alle Module wurden exportiert nach: " & path
End Sub



Public Sub runTestPdfReport()
Dim p As String
Debug.Print modPdfExportService.ExportDocumentToPdf(1, p)
Debug.Print p
End Sub
