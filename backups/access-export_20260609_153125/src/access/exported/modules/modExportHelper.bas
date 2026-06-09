Attribute VB_Name = "modExportHelper"
Option Compare Database
Option Explicit

Public Sub ExportAllModulesUtf8()
    Dim comp As Object
    Dim stream As Object
    Dim exportPath As String
    Dim fileExt As String
    
    ' Ordnerpfad festlegen (muss existieren)
    exportPath = CurrentProject.path & "\export\"
    If Dir(exportPath, vbDirectory) = "" Then MkDir exportPath

    ' Alle Komponenten des aktuellen VBA-Projekts durchlaufen
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        ' Dateiendung je nach Typ bestimmen
        Select Case comp.Type
            Case 1: fileExt = ".bas"  ' Standard-Modul
            Case 2: fileExt = ".cls"  ' Klasse / Formular-Code
            Case 3: fileExt = ".frm"  ' UserForm (falls vorhanden)
            Case Else: fileExt = ".cls"
        End Select
        
        ' Code nur exportieren, wenn Zeilen vorhanden sind
        If comp.CodeModule.CountOfLines > 0 Then
            Set stream = CreateObject("ADODB.Stream")
            With stream
                .Type = 2 ' adTypeText
                .Charset = "utf-8"
                .Open
                ' Gesamten Text des Moduls in den Stream schreiben
                .WriteText comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
                .SaveToFile exportPath & comp.Name & fileExt, 2 ' Überschreiben erlaubt
                .Close
             End With
        End If
    Next comp
    
    MsgBox "Alle Module erfolgreich als UTF-8 exportiert!", vbInformation
End Sub

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