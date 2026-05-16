Attribute VB_Name = "modExportHelper"
Option Compare Database
Option Explicit

Sub ExportAllModules()
    Dim Obj As accessObject
    Dim path As String
    
    ' Ordnerpfad festlegen, wohin exportiert werden soll
    path = "C:\Users\gomic\OneDrive\Documents\GitHub\easis-v4-access-framework\src\access\exported\modules\"
    
    ' Ordner erstellen, falls nicht vorhanden
    If Dir(path, vbDirectory) = "" Then MkDir path
    
    ' Standardmodule
    For Each Obj In CurrentProject.AllModules
        If Not Obj.Name = "modExportHelper" Then Application.SaveAsText acModule, Obj.Name, path & Obj.Name & ".bas"
    Next Obj
    
    ' Klassenmodule
    For Each Obj In CurrentProject.AllModules
        ' Klassenmodule haben oft den Typ acClassModule
        ' Application.SaveAsText acClassModule, obj.Name, path & obj.Name & ".cls"
    Next Obj
    
    MsgBox "Alle Module wurden exportiert nach: " & path
End Sub

