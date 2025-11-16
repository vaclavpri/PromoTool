Attribute VB_Name = "RealPrice_Module"

Public Sub RealPrice(TargetWorkbook As Workbook, selFirstRow As Long, selRowCount As Long)
    On Error GoTo ErrorHandler
    
    Debug.Print "=== RealPrice START ==="
    Debug.Print "Parametry: firstRow=" & selFirstRow & ", rowCount=" & selRowCount
    
    Application.ScreenUpdating = False
    
    ' Kontrola parametrù
    If selFirstRow = 0 Or selRowCount = 0 Then
        MsgBox "Neplatné parametry!", vbCritical
        Exit Sub
    End If
    
    Dim textList As Worksheet
    Dim promoplanList As Worksheet
    Set textList = TargetWorkbook.Sheets("Text")
    Set promoplanList = TargetWorkbook.Sheets("Promoplan")
    Debug.Print "Listy OK"
    
    ' Urèení sloupcù v Text listu
    Dim colFamily As Long, colWeeks As Long, colPromo As Long, colHero As Long, colRealPrice As Long, colPromoID As Long
    colFamily = textList.Range("tFamily").Column
    colWeeks = textList.Range("tWeeks").Column
    colPromo = textList.Range("tPromo").Column
    colHero = textList.Range("tHero").Column
    colRealPrice = textList.Range("tRealPromoPrice").Column
    colPromoID = textList.Range("tPromoID").Column
    Debug.Print "Sloupce: Family=" & colFamily & ", Weeks=" & colWeeks & ", Promo=" & colPromo
    
    ' Naètení dat z výbìru do pole
    Dim textData As Variant
    ReDim textData(1 To selRowCount, 1 To 6)
    
    Dim i As Long
    For i = 1 To selRowCount
        textData(i, 1) = textList.Cells(selFirstRow + i - 1, colFamily).value      ' Family
        textData(i, 2) = textList.Cells(selFirstRow + i - 1, colWeeks).value       ' Weeks
        textData(i, 3) = textList.Cells(selFirstRow + i - 1, colPromo).value       ' Promo
        textData(i, 4) = textList.Cells(selFirstRow + i - 1, colHero).value        ' Hero
        textData(i, 5) = textList.Cells(selFirstRow + i - 1, colRealPrice).value   ' RealPromoPrice
        textData(i, 6) = textList.Cells(selFirstRow + i - 1, colPromoID).value     ' PromoID
    Next i
    Debug.Print "Data z Text naètena: " & selRowCount & " øádkù"
    
    ' ZMÌNA: Seskupit øádky podle PromoID
    Dim promoGroups As Object
    Set promoGroups = CreateObject("Scripting.Dictionary")
    
    Dim promoID As String
    For i = 1 To selRowCount
        promoID = Trim(CStr(textData(i, 6)))
        If promoID <> "" Then
            If Not promoGroups.Exists(promoID) Then
                promoGroups.Add promoID, i & "," ' První øádek pro tuto promoci
            Else
                promoGroups(promoID) = promoGroups(promoID) & i & ","
            End If
        End If
    Next i
    
    Debug.Print "Poèet unikátních promocí: " & promoGroups.Count
    
    ' Dynamické nalezení øádku WeekRow v Promoplan
    Dim weekRowNumber As Long
    weekRowNumber = FindWeekRow(promoplanList)
    
    If weekRowNumber = 0 Then
        MsgBox "Øádek s komentáøem 'WeekRow' nebyl nalezen v Promoplan!", vbCritical
        Exit Sub
    End If
    Debug.Print "WeekRow: " & weekRowNumber
    
    ' Najít sloupec s hodnotou 1 (první týden)
    Dim firstWeekColumn As Long
    firstWeekColumn = FindFirstWeekColumn(promoplanList, weekRowNumber)
    
    If firstWeekColumn = 0 Then
        MsgBox "Sloupec s hodnotou 1 (první týden) nebyl nalezen!", vbCritical
        Exit Sub
    End If
    Debug.Print "FirstWeekColumn: " & firstWeekColumn
    
    ' Najít sloupec s komentáøem "Fami"
    Dim famiColumnNumber As Long
    famiColumnNumber = FindFamiColumn(promoplanList, weekRowNumber)
    
    If famiColumnNumber = 0 Then
        MsgBox "Sloupec s komentáøem 'Fami' nebyl nalezen!", vbCritical
        Exit Sub
    End If
    Debug.Print "FamiColumn: " & famiColumnNumber
    
    ' Najít poslední øádek a sloupec v Promoplan
    Dim lastRow As Long
    Dim lastColumn As Long
    lastRow = promoplanList.Cells(promoplanList.rows.Count, famiColumnNumber).End(xlUp).row
    lastColumn = promoplanList.Cells(weekRowNumber, promoplanList.Columns.Count).End(xlToLeft).Column
    Debug.Print "Promoplan rozsah: LastRow=" & lastRow & ", LastColumn=" & lastColumn
    
    ' Naètení dat z Promoplan
    Dim famiData As Variant
    Dim weeksData As Variant
    
    If lastRow > weekRowNumber Then
        famiData = promoplanList.Range(promoplanList.Cells(weekRowNumber + 1, famiColumnNumber), _
                                       promoplanList.Cells(lastRow, famiColumnNumber)).value
    Else
        MsgBox "Žádná data family v Promoplan!", vbExclamation
        Exit Sub
    End If
    
    If lastColumn >= firstWeekColumn Then
        weeksData = promoplanList.Range(promoplanList.Cells(weekRowNumber, firstWeekColumn), _
                                        promoplanList.Cells(weekRowNumber, lastColumn)).value
    Else
        MsgBox "Žádná data týdnù v Promoplan!", vbExclamation
        Exit Sub
    End If
    
    Debug.Print "Promoplan data naètena: " & UBound(famiData, 1) & " family, " & UBound(weeksData, 2) & " týdnù"
    
    ' ZMÌNA: Zpracovat každou promoci samostatnì
    Dim matchCount As Long
    matchCount = 0
    
    Dim promoKey As Variant
    For Each promoKey In promoGroups.Keys
        Debug.Print "=== Zpracovávám PromoID: " & promoKey & " ==="
        
        ' Získat øádky pro tuto promoci
        Dim rowIndices() As String
        rowIndices = Split(Left(promoGroups(promoKey), Len(promoGroups(promoKey)) - 1), ",")
        
        ' Použít první øádek pro urèení týdne
        Dim firstRowIndex As Long
        firstRowIndex = CLng(rowIndices(0))
        
        ' Urèení týdne
        Dim w As Long
        Dim weekText As String
        Dim FirstWeek As Long
        Dim dashPos As Long
        
        weekText = CStr(textData(firstRowIndex, 2))
        
        If Len(weekText) < 3 Then
            w = CLng(weekText)
        Else
            dashPos = InStr(1, weekText, "-")
            If dashPos > 0 Then
                FirstWeek = CLng(Left(weekText, dashPos - 1))
                w = FirstWeek
            Else
                w = CLng(weekText)
            End If
        End If
        Debug.Print "  Hledaný týden: " & w
        
        ' Zápis do Promoplan
        Dim weekCol As Long
        Dim famiRow As Long
        Dim a As Long
        Dim rowIdx As Long
        
        For weekCol = 1 To UBound(weeksData, 2)
            If CLng(weeksData(1, weekCol)) = w Then
                Debug.Print "  Našel týden " & w & " ve sloupci " & (firstWeekColumn + weekCol - 1)
                
                For famiRow = 1 To UBound(famiData, 1)
                    ' Projít všechny øádky této promoce
                    For a = LBound(rowIndices) To UBound(rowIndices)
                        rowIdx = CLng(rowIndices(a))
                        
                        ' Kontrola: Family se shoduje A Hero je "A"
                        If CStr(famiData(famiRow, 1)) = CStr(textData(rowIdx, 1)) And _
                           UCase(CStr(textData(rowIdx, 4))) = "A" Then
                            
                            ' Zápis do buòky
                            Dim targetRow As Long
                            Dim targetCol As Long
                            targetRow = weekRowNumber + famiRow
                            targetCol = firstWeekColumn + weekCol - 1
                            
                            Dim cellValue As String
                            cellValue = textData(rowIdx, 3) & " " & textData(rowIdx, 5)
                            
                            promoplanList.Cells(targetRow, targetCol).value = cellValue
                            matchCount = matchCount + 1
                            
                            Debug.Print "  Zapsal: '" & cellValue & "' na [" & targetRow & "," & targetCol & "] pro Family: " & textData(rowIdx, 1)
                        End If
                    Next a
                Next famiRow
                
                Exit For ' Týden nalezen
            End If
        Next weekCol
    Next promoKey
    
    Application.ScreenUpdating = True
    
    Debug.Print "=== RealPrice END === Zapsáno: " & matchCount & " bunìk"
    
    If matchCount > 0 Then
        MsgBox "RealPrice dokonèeno! Zapsáno " & matchCount & " promocí do Promoplan.", vbInformation
    Else
        MsgBox "Žádné shody nebyly nalezeny. Zkontrolujte Family a Hero hodnoty.", vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Debug.Print "CHYBA na øádku: " & Erl
    Debug.Print "Err " & Err.Number & ": " & Err.Description
    MsgBox "Chyba v RealPrice: " & Err.Description, vbCritical
    Stop
    Resume
End Sub

' Pomocná funkce pro nalezení sloupce s hodnotou 1
Private Function FindFirstWeekColumn(ws As Worksheet, weekRowNumber As Long) As Long
    Dim col As Long
    Dim lastCol As Long
    
    lastCol = ws.Cells(weekRowNumber, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        If ws.Cells(weekRowNumber, col).value = 1 Then
            FindFirstWeekColumn = col
            Exit Function
        End If
    Next col
    
    FindFirstWeekColumn = 0
End Function

' Pomocná funkce pro nalezení sloupce s komentáøem "Fami"
Private Function FindFamiColumn(ws As Worksheet, rowNumber As Long) As Long
    Dim cell As Range
    Dim lastCol As Long
    
    lastCol = ws.Cells(rowNumber, ws.Columns.Count).End(xlToLeft).Column
    
    For Each cell In ws.Range(ws.Cells(rowNumber, 1), ws.Cells(rowNumber, lastCol))
        If Not cell.comment Is Nothing Then
            If InStr(1, cell.comment.Text, "Fami", vbTextCompare) > 0 Then
                FindFamiColumn = cell.Column
                Exit Function
            End If
        End If
    Next cell
    
    FindFamiColumn = 0
End Function

