Attribute VB_Name = "Backup"

Public Sub BackupWorkbook_Shared(TargetWorkbook As Workbook)
    On Error GoTo ErrorHandler
    
    Debug.Print "=== BackupWorkbook_Shared START ==="
    
    Dim savedate As Date
    Dim savetime As Date
    Dim formattime As String
    Dim formatdate As String
    Dim backupfolder As String
    
    savedate = Date
    savetime = Time
    formattime = Format(savetime, "hh.mm.ss")
    formatdate = Format(savedate, "dd-MM-yyyy")
    
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' Naètení cesty z Settings
    Dim path1 As String, path2 As String
    
    On Error Resume Next
    path1 = Trim(TargetWorkbook.Sheets("Settings").Range("B11").value)
    path2 = Trim(TargetWorkbook.Sheets("Settings").Range("B12").value)
    On Error GoTo ErrorHandler
         
    Debug.Print "Path1: " & path1
    Debug.Print "Path2: " & path2
    
    ' Kontrola cest
    If Len(Dir(path1, vbDirectory)) > 0 Then
        backupfolder = path1
        Debug.Print "Použita Path1"
    ElseIf Len(Dir(path2, vbDirectory)) > 0 Then
        backupfolder = path2
        Debug.Print "Použita Path2"
    Else
        Debug.Print "CHYBA: Žádná cesta k BackUp složce není dostupná"
        MsgBox "Cesta k záložnímu souboru nebyla nalezena!" & vbCrLf & _
               "Path1: " & path1 & vbCrLf & _
               "Path2: " & path2, vbExclamation
        GoTo CleanUp
    End If
    
    ' Ujistit se, že cesta konèí backslashem
    If Right(backupfolder, 1) <> "\" Then
        backupfolder = backupfolder & "\"
    End If
    
    ' Uložení kopie
    Dim backupFileName As String
    backupFileName = backupfolder & formatdate & " " & formattime & " " & TargetWorkbook.Name
    
    Debug.Print "Ukládám backup: " & backupFileName
    TargetWorkbook.SaveCopyAs fileName:=backupFileName
    Debug.Print "Backup úspìšnì uložen"
    
CleanUp:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Debug.Print "=== BackupWorkbook_Shared END ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "CHYBA v BackupWorkbook_Shared: " & Err.Description
    MsgBox "Došlo k chybì pøi ukládání zálohy: " & Err.Description, vbCritical
    GoTo CleanUp
End Sub

Sub ExportAllModules()
    Dim vbComp As Object
    Dim exportPath As String

    ' Nastavte cestu, kam chcete exportovat
    exportPath = "C:\Temp\VBA_Export\"

    ' Vytvoøit složku, pokud neexistuje
    On Error Resume Next
    MkDir exportPath
    On Error GoTo 0

    ' Projít všechny komponenty v projektu
    For Each vbComp In ThisWorkbook.VBProject.VBComponents

        Select Case vbComp.Type
            Case 1 ' Module
                vbComp.Export exportPath & vbComp.Name & ".bas"
                Debug.Print "Exportován: " & vbComp.Name & ".bas"

            Case 2 ' Class Module
                vbComp.Export exportPath & vbComp.Name & ".cls"
                Debug.Print "Exportován: " & vbComp.Name & ".cls"

            Case 3 ' UserForm
                vbComp.Export exportPath & vbComp.Name & ".frm"
                Debug.Print "Exportován: " & vbComp.Name & ".frm"

            Case 100 ' Document (ThisWorkbook, Sheet1, atd.)
                vbComp.Export exportPath & vbComp.Name & ".cls"
                Debug.Print "Exportován: " & vbComp.Name & ".cls"
        End Select
    Next vbComp

    MsgBox "Export dokonèen do: " & exportPath, vbInformation
End Sub
