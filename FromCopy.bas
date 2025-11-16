Attribute VB_Name = "FromCopy"

Public Sub FromCopy4_Shared(TargetWorkbook As Workbook, SelectedRange As Range)
    On Error GoTo ErrorHandler
    
    Debug.Print "=== FromCopy4_Shared START ==="
    
    ' 1. ODEMKNUTÍ
    Call UnlockText(TargetWorkbook)
    Call RemoveFilterIfApplied(TargetWorkbook)
    
    ' 2. ZÍSKÁNÍ STARÉHO PromoID
    If SelectedRange.Cells(1).comment Is Nothing Then
        MsgBox "Vybraná buòka nemá komentáø s PromoID!", vbCritical
        Call LockText(TargetWorkbook)
        Exit Sub
    End If
    
    Dim oldPromoID As String
    oldPromoID = Left(SelectedRange.Cells(1).comment.Text, 8)
    Debug.Print "Starý PromoID: " & oldPromoID
    
    ' 3. REFERENCE
    Dim textList As Worksheet
    Set textList = TargetWorkbook.Sheets("Text")
    
    Dim lastRow As Long
    lastRow = textList.Cells(textList.rows.Count, 1).End(xlUp).row
    
    ' 4. NAÈTENÍ DAT
    Dim promoIDRange As Variant, promoRange As Variant
    promoIDRange = textList.Range("tPromoID").value
    promoRange = textList.Range("tPromo").value
    
    ' 5. ZJIŠTÌNÍ TYPU PROMOCE
    Debug.Print "Zjišuji typ promoce..."
    Dim promoType As String
    promoType = GetPromoTypeFromPromo(oldPromoID, promoIDRange, promoRange)
    
    If promoType = "" Then
        MsgBox "Typ promoce pro ID " & oldPromoID & " nebyl nalezen!", vbCritical
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
    
    ' 8. VÝPOÈET DATUMÙ - ZMÌNA: Volat sdílenou funkci
    Debug.Print "Nastavuji datumy pro typ: " & promoType
    If Not SetupPromoByListBoxValue_Shared(promoType, SelectedRange, Promo, TargetWorkbook, False) Then
        MsgBox "Chyba pøi výpoètu nových datumù!", vbCritical
        Call LockText(TargetWorkbook)
        Exit Sub
    End If
    
    Debug.Print "Nové datumy vypoèítány"
    
    ' 9. GENEROVÁNÍ NOVÉHO PromoID
    Dim newPromoID As String
    newPromoID = GenerateID(TargetWorkbook)
    Debug.Print "Nové PromoID: " & newPromoID
    
    ' 10. KOPÍROVÁNÍ ØÁDKÙ
    Call CopyPromoRows(textList, oldPromoID, promoIDRange, promoRange, lastRow, Promo, newPromoID, TargetWorkbook)
    
    ' 11. AKTUALIZACE KOMENTÁØÙ
    Call UpdateComments_Shared(SelectedRange, newPromoID)
    
    ' 12. FINÁLNÍ ÚPRAVY
    Application.CutCopyMode = False
    Call ApplyFilterToRow2(TargetWorkbook)
    Call SortIt(TargetWorkbook)
    Call rColor(TargetWorkbook)
    
    ' 13. ZAMKNUTÍ
    Call LockText(TargetWorkbook)
    
    Debug.Print "=== FromCopy4_Shared END ==="
    MsgBox "Promoce zkopírována. Nové PromoID: " & newPromoID, vbInformation
    Exit Sub
    
ErrorHandler:
    Debug.Print "CHYBA: " & Err.Description
    MsgBox "Chyba: " & Err.Description, vbCritical
    Call LockText(TargetWorkbook)
End Sub

' ===================================================================
' POMOCNÉ FUNKCE
' ===================================================================

Public Function GetPromoTypeFromPromo(promoID As String, promoIDRange As Variant, promoRange As Variant) As String
    Dim i As Long
    For i = 1 To UBound(promoIDRange, 1)
        If CStr(promoIDRange(i, 1)) = CStr(promoID) Then
            GetPromoTypeFromPromo = Trim(promoRange(i, 1))
            Debug.Print "Nalezen typ: '" & GetPromoTypeFromPromo & "'"
            Exit Function
        End If
    Next i
    GetPromoTypeFromPromo = ""
End Function

Private Sub CopyPromoRows(textList As Worksheet, oldPromoID As String, promoIDRange As Variant, promoRange As Variant, lastRow As Long, PromoObj As Object, newPromoID As String, TargetWorkbook As Workbook)
    Dim i As Long, j As Long
    j = 1
    
    For i = 1 To UBound(promoIDRange, 1)
        If CStr(promoIDRange(i, 1)) = CStr(oldPromoID) Then
            textList.rows(i).Copy Destination:=textList.Cells(lastRow + j, 1)
            
            With textList
                .Cells(lastRow + j, textList.Range("tCom").Column).value = promoRange(i, 1) & " " & PromoObj.WeekRange
                .Cells(lastRow + j, textList.Range("tNakupOd").Column).value = PromoObj.StartPurchase
                .Cells(lastRow + j, textList.Range("tNakupDo").Column).value = PromoObj.EndPurchase
                .Cells(lastRow + j, textList.Range("tAkceOd").Column).value = PromoObj.startAkce
                .Cells(lastRow + j, textList.Range("tAkceDo").Column).value = PromoObj.endAkce
                .Cells(lastRow + j, textList.Range("tSortFrom").Column).value = PromoObj.sortFrom
                .Cells(lastRow + j, textList.Range("tSortTo").Column).value = PromoObj.sortTo
                .Cells(lastRow + j, textList.Range("tWeeks").Column).value = PromoObj.WeekRange
                .Cells(lastRow + j, textList.Range("tWeeksT").Column).value = PromoObj.weekRangeT
                .Cells(lastRow + j, textList.Range("tPromoID").Column).value = newPromoID
            End With
            
            Debug.Print "Zkopírován øádek " & i & " -> " & (lastRow + j)
            j = j + 1
        End If
    Next i
    
    Debug.Print "Celkem zkopírováno: " & (j - 1) & " øádkù"
End Sub

Private Sub UpdateComments_Shared(SelectedRange As Range, newID As String)
    Dim cell As Range
    For Each cell In SelectedRange
        If Not cell.comment Is Nothing Then cell.comment.Delete
        cell.AddComment CStr(newID)
    Next cell
End Sub
