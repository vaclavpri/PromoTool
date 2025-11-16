Attribute VB_Name = "ID"
'Public Function GenerateID(targetWorkbook As Workbook)
'    Dim ws As Worksheet
'    Set ws = targetWorkbook.Sheets("IDList")
'
'    ' Pøeète aktuální èíslo z buòky A2
'    Dim currentID As Long
'    currentID = ws.Cells(2, "A").value
'
'    ' Zvýší èíslo o 1
'    Dim newID As Long
'    newID = currentID + 1
'
'    ' Zapíše nové èíslo zpìt do buòky A2
'    ws.Cells(2, "A").value = newID
'
'    ' Vrátí nové ID
'    GenerateID = newID
'End Function

Public Function GenerateID(TargetWorkbook As Workbook) As String
    
    Dim wsSettings As Worksheet
    Set wsSettings = TargetWorkbook.Sheets("Settings")
    
    ' Pøeète hodnoty z listu Settings
    Dim prefix As String
    Dim counterValue As Long
    
    prefix = wsSettings.Range("B9").value ' První 3 znaky
    counterValue = wsSettings.Range("C9").value ' Èíslo, které se pøevede na 5místné
    
    ' Kontrola délky prefixu
    If Len(prefix) <> 3 Then
        MsgBox "Hodnota v B9 musí mít pøesnì 3 znaky!", vbExclamation
        GenerateID = ""
        Exit Function
    End If
    
    ' Pøeète aktuální èíslo z buòky A2
    Dim currentID As Long
    currentID = wsSettings.Cells(9, "C").value
    
    ' Zvýší èíslo o 1
    Dim newID As Long
    newID = currentID + 1
    
    ' Zapíše nové èíslo zpìt do buòky A2
    wsSettings.Cells(9, "C").value = newID
    
    ' Vytvoøí PromoID: 3 znaky z B9 + 5místné èíslo z C9 (s nulami na zaèátku)
    Dim formattedCounter As String
    formattedCounter = Format(counterValue, "00000") ' Pøevede na 5místné èíslo (00001, 00123, 12345 atd.)
    
    ' Vrátí nové ID ve formátu: PREFIX + 5MÍSTNÉ_ÈÍSLO (celkem 8 znakù)
    GenerateID = prefix & formattedCounter
    
End Function
