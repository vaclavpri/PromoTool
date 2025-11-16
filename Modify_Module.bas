Attribute VB_Name = "Modify_Module"
' ===================================================================
' Modify_Shared - Pøesun promoce do jiného týdne (bez kopírování)
' ===================================================================
Public Sub Modify_Shared(TargetWorkbook As Workbook, SelectedRange As Range)
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Modify_Shared START ==="
    Debug.Print "Vybraný rozsah: " & SelectedRange.Address
    Debug.Print "Barva první buòky: " & SelectedRange.Cells(1).Interior.Color
    
    ' 1. ODEMKNUTÍ
    Call UnlockText(TargetWorkbook)
    Call RemoveFilterIfApplied(TargetWorkbook)
    
    ' 2. ZÍSKÁNÍ PromoID
    If SelectedRange.Cells(1).comment Is Nothing Then
        MsgBox "Vybraná buòka nemá komentáø s PromoID!", vbCritical
        Call LockText(TargetWorkbook)
        Exit Sub
    End If
    
    Dim promoID As String
    promoID = Left(SelectedRange.Cells(1).comment.Text, 8)
    Debug.Print "PromoID: " & promoID
    
    ' 3. POÈET VYBRANÝCH BUNÌK
    Dim selectedCellsCount As Long
    selectedCellsCount = SelectedRange.Cells.Count
    Debug.Print "Poèet vybraných bunìk: " & selectedCellsCount
    
    ' 4. NAJÍT EXISTUJÍCÍ BUÒKY
    Dim ws As Worksheet
    Set ws = SelectedRange.Worksheet
    
    Dim existingCells As Collection
    Set existingCells = FindCellsWithPromoID(ws, promoID)
    Debug.Print "Poèet existujících bunìk: " & existingCells.Count
    
    ' 5. ZJIŠTÌNÍ TYPU
    Dim textList As Worksheet
    Set textList = TargetWorkbook.Sheets("Text")
    
    Dim promoIDRange As Variant, promoRange As Variant
    promoIDRange = textList.Range("tPromoID").value
    promoRange = textList.Range("tPromo").value
    
    Dim promoType As String
    promoType = GetPromoTypeFromPromo(promoID, promoIDRange, promoRange)
    
    If promoType = "" Then
        MsgBox "Typ promoce pro ID " & promoID & " nebyl nalezen!", vbCritical
        Call LockText(TargetWorkbook)
        Exit Sub
    End If
    
    Debug.Print "Typ promoce: " & promoType
    
    ' 6. NAÈTENÍ DAT TÝDNÙ
    Call WeeksArray(TargetWorkbook, SelectedRange)
    
    ' 7. VYTVOØENÍ PROMO OBJEKTU
    Dim Promo As Object
    Set Promo = CreatePromoInstance()
    
    If Promo Is Nothing Then
        MsgBox "Nepodaøilo se vytvoøit Promo instanci!", vbCritical
        Call LockText(TargetWorkbook)
        Exit Sub
    End If
   
'    ' 8. VÝPOÈET NOVÝCH DATUMÙ
'    If Not userFormObj.SetupPromoByListBoxValue(promoType, SelectedRange, Promo) Then
'        MsgBox "Chyba pøi výpoètu nových datumù!", vbCritical
'        Call LockText(TargetWorkbook)
'        Exit Sub
'    End If
    ' 8. VÝPOÈET NOVÝCH DATUMÙ
    Debug.Print "Nastavuji datumy pro typ: " & promoType
    
    ' VOLAT PØÍMO SetupPromoByListBoxValue ze SharedCode (ne pøes UserForm)
    If Not SetupPromoByListBoxValue_Shared(promoType, SelectedRange, Promo, TargetWorkbook) Then
        MsgBox "Chyba pøi výpoètu nových datumù!", vbCritical
        Call LockText(TargetWorkbook)
        Exit Sub
    End If

Debug.Print "Nové datumy vypoèítány"
    ' 9. AKTUALIZACE DATUMÙ V LISTU TEXT
    Call UpdatePromoDates(textList, promoID, promoIDRange, Promo)
    
    ' 10. ÚPRAVA BUNÌK A BAREV - POUZE pokud se poèet zmìnil!
    If selectedCellsCount > existingCells.Count Then
        ' PØIDÁNÍ BUNÌK
        Debug.Print "Pøidávám buòky..."
        Call AddCellsToPromo(SelectedRange, existingCells, promoID)
        
    ElseIf selectedCellsCount < existingCells.Count Then
        ' ODEBRÁNÍ BUNÌK
        Debug.Print "Odebírám buòky..."
        Call RemoveCellsFromPromo(SelectedRange, existingCells)
        
    Else
        ' STEJNÝ POÈET - nedìlat nic, uživatel už buòky pøesunul myší
        Debug.Print "Stejný poèet bunìk - žádné zmìny v buòkách (pøesunuto myší)"
    End If
    
    ' 11. FINÁLNÍ ÚPRAVY
    Application.CutCopyMode = False
    Call ApplyFilterToRow2(TargetWorkbook)
    Call SortIt(TargetWorkbook)
    Call rColor(TargetWorkbook)
    
    ' 12. ZAMKNUTÍ
    Call LockText(TargetWorkbook)
    
    Debug.Print "=== Modify_Shared END ==="
    MsgBox "Promoce aktualizována.", vbInformation
    Exit Sub
    
ErrorHandler:
    Debug.Print "CHYBA: " & Err.Description
    MsgBox "Chyba: " & Err.Description, vbCritical
    Call LockText(TargetWorkbook)
End Sub

' ===================================================================
' POMOCNÉ FUNKCE PRO MODIFY
' ===================================================================

' Najít všechny buòky s daným PromoID v komentáøích
Private Function FindCellsWithPromoID(ws As Worksheet, promoID As String) As Collection
    Set FindCellsWithPromoID = New Collection
    
    Dim cell As Range
    Dim comment As comment
    
    ' Projít všechny buòky s komentáøi
    For Each comment In ws.Comments
        If Not comment.Parent Is Nothing Then
            If Left(comment.Text, 8) = promoID Then
                FindCellsWithPromoID.Add comment.Parent
                Debug.Print "Nalezena buòka s ID: " & comment.Parent.Address
            End If
        End If
    Next comment
End Function

' Zjištìní typu promoce (pro Modify)
Private Function GetPromoTypeFromPromo_Modify(promoID As String, promoIDRange As Variant, promoRange As Variant) As String
    Dim i As Long
    For i = 1 To UBound(promoIDRange, 1)
        If CStr(promoIDRange(i, 1)) = CStr(promoID) Then
            GetPromoTypeFromPromo_Modify = Trim(promoRange(i, 1))
            Exit Function
        End If
    Next i
    GetPromoTypeFromPromo_Modify = ""
End Function


' ===================================================================
' FUNKCE - Najde øádek podle PromoID
' ===================================================================
Private Sub UpdatePromoDates(textList As Worksheet, promoID As String, promoIDRange As Variant, PromoObj As Object)
    Dim rowsUpdated As Long
    rowsUpdated = 0
    
    Debug.Print "UpdatePromoDates: Hledám øádky s PromoID: " & promoID
    
    ' MÍSTO procházení array, najdeme pøímo øádky v listu
    Dim lastRow As Long
    lastRow = textList.Cells(textList.rows.Count, textList.Range("tPromoID").Column).End(xlUp).row
    
    Dim promoIDColumn As Long
    promoIDColumn = textList.Range("tPromoID").Column
    
    Dim currentRow As Long
    For currentRow = 2 To lastRow  ' Zaèínáme od øádku 2 (po hlavièce)
        If CStr(textList.Cells(currentRow, promoIDColumn).value) = CStr(promoID) Then
            
            Debug.Print "Nalezen a aktualizuji øádek: " & currentRow
            
            ' Aktualizovat TENTO øádek
            With textList
                ' Datumy
                .Cells(currentRow, .Range("tNakupOd").Column).value = PromoObj.StartPurchase
                .Cells(currentRow, .Range("tNakupDo").Column).value = PromoObj.EndPurchase
                .Cells(currentRow, .Range("tAkceOd").Column).value = PromoObj.startAkce
                .Cells(currentRow, .Range("tAkceDo").Column).value = PromoObj.endAkce
                .Cells(currentRow, .Range("tSortFrom").Column).value = PromoObj.sortFrom
                .Cells(currentRow, .Range("tSortTo").Column).value = PromoObj.sortTo
                
                ' Týdny
                .Cells(currentRow, .Range("tWeeks").Column).value = PromoObj.WeekRange
                .Cells(currentRow, .Range("tWeeksT").Column).value = PromoObj.weekRangeT
                
                ' Komentáø
                Dim currentComment As String
                currentComment = .Cells(currentRow, .Range("tCom").Column).value
                
                Debug.Print "  Starý komentáø: " & currentComment
                
                Dim parts() As String
                parts = Split(currentComment, " ")
                
                If UBound(parts) >= 0 Then
                    Dim prefix As String
                    prefix = parts(0)
                    .Cells(currentRow, .Range("tCom").Column).value = prefix & " " & PromoObj.WeekRange
                    Debug.Print "  Nový komentáø: " & prefix & " " & PromoObj.WeekRange
                End If
                
                Debug.Print "  Nové datumy:"
                Debug.Print "    NakupOd: " & PromoObj.StartPurchase
                Debug.Print "    NakupDo: " & PromoObj.EndPurchase
                Debug.Print "    AkceOd: " & PromoObj.startAkce
                Debug.Print "    AkceDo: " & PromoObj.endAkce
                Debug.Print "    Týdny: " & PromoObj.WeekRange
            End With
            
            rowsUpdated = rowsUpdated + 1
        End If
    Next currentRow
    
    Debug.Print "UpdatePromoDates: Celkem aktualizováno øádkù: " & rowsUpdated
    
    If rowsUpdated = 0 Then
        Debug.Print "VAROVÁNÍ: Nebyl aktualizován žádný øádek pro PromoID: " & promoID
    End If
End Sub

' ===================================================================
' Pøidání nových bunìk (oznaèeno VÍCE než je aktuálnì)
' ===================================================================
Private Sub AddCellsToPromo(SelectedRange As Range, existingCells As Collection, promoID As String)
    Dim cell As Range
    Dim existingCell As Range
    Dim isExisting As Boolean
    
    ' Získat barvu z první buòky výbìru
    Dim targetColor As Long
    targetColor = SelectedRange.Cells(1).Interior.Color
    Debug.Print "Cílová barva: " & targetColor
    
    ' Pro každou novì vybranou buòku
    For Each cell In SelectedRange
        isExisting = False
        
        ' Zkontrolovat, zda už má komentáø s tímto ID
        For Each existingCell In existingCells
            If cell.Address = existingCell.Address Then
                isExisting = True
                Exit For
            End If
        Next existingCell
        
        ' Pokud je to NOVÁ buòka
        If Not isExisting Then
            ' Pøidat komentáø
            If Not cell.comment Is Nothing Then
                cell.comment.Delete
            End If
            cell.AddComment CStr(promoID)
            
            ' Nastavit barvu
            cell.Interior.Color = targetColor
            
            Debug.Print "Pøidán komentáø a barva do: " & cell.Address
        End If
    Next cell
End Sub

' ===================================================================
' Odebrání bunìk (oznaèeno MÉNÌ než je aktuálnì)
' ===================================================================
Private Sub RemoveCellsFromPromo(SelectedRange As Range, existingCells As Collection)
    Debug.Print "=== RemoveCellsFromPromo START ==="
    Debug.Print "Poèet existujících bunìk: " & existingCells.Count
    Debug.Print "Poèet vybraných bunìk: " & SelectedRange.Cells.Count
    
    Dim existingCell As Range
    Dim cell As Range
    Dim isSelected As Boolean
    
    ' Pro každou existující buòku
    For Each existingCell In existingCells
        Debug.Print "  Kontroluji buòku: " & existingCell.Address
        isSelected = False
        
        ' Zkontrolovat, zda je v novém výbìru
        For Each cell In SelectedRange.Cells
            If cell.Address = existingCell.Address Then
                isSelected = True
                Debug.Print "    Je ve výbìru"
                Exit For
            End If
        Next cell
        
        ' Pokud NENÍ ve výbìru
        If Not isSelected Then
            Debug.Print "    NENÍ ve výbìru - odebírám"
            
            ' Odstranit komentáø
            If Not existingCell.comment Is Nothing Then
                existingCell.comment.Delete
                Debug.Print "    Odstranìn komentáø z: " & existingCell.Address
            End If
            
            ' Odstranit barvu (nastavit na bílou)
            existingCell.Interior.ColorIndex = xlNone
            Debug.Print "    Odstranìna barva z: " & existingCell.Address
        End If
    Next existingCell
    
    Debug.Print "=== RemoveCellsFromPromo END ==="
End Sub
