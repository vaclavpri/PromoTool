Attribute VB_Name = "Clear"
' ===================================================================
' Clear1 - mazání promocí
' ===================================================================
Public Sub Clear1(TargetWorkbook As Workbook)
    ' Získání PromoID z oznaèených bunìk
    Dim promoIDsToDelete As Collection
    Set promoIDsToDelete = New Collection
    
    Dim cell As Range
    Dim promoID As String
    
    ' Získání všech PromoID z komentáøù oznaèených bunìk
    On Error Resume Next
    For Each cell In TargetWorkbook.Application.Selection
        If Not cell.comment Is Nothing Then
            ' Získat prvních 8 znakù z komentáøe (PromoID)
            promoID = Left(cell.comment.Text, 8)
            
            If Len(promoID) = 8 Then  ' Kontrola délky 8 znakù
                promoIDsToDelete.Add promoID, promoID
            End If
        End If
    Next cell
    On Error GoTo 0
    
    If promoIDsToDelete.Count = 0 Then
        MsgBox "V oznaèených buòkách nebyla nalezena žádná PromoID v komentáøích.", vbInformation
        Exit Sub
    End If
    
    ' POTVRZENÍ PØED SMAZÁNÍM
    Dim response As VbMsgBoxResult
    response = MsgBox("Opravdu chcete smazat " & promoIDsToDelete.Count & " promocí?", _
                      vbYesNo + vbQuestion, "Potvrzení smazání")
    
    If response = vbNo Then Exit Sub
    
    ' Smazání všech promocí s danými PromoID
    Dim searchValue As String
    Dim i As Long
    
    For i = 1 To promoIDsToDelete.Count
        searchValue = promoIDsToDelete(i)
        Call DeletePromoByID(TargetWorkbook, searchValue)
    Next i
    
    MsgBox "Bylo odstranìno " & promoIDsToDelete.Count & " promocí.", vbInformation
    
End Sub

Private Sub DeletePromoByID(TargetWorkbook As Workbook, searchValue As String)
    Call UnlockText(TargetWorkbook)
    Call RemoveFilterIfApplied(TargetWorkbook)
    
    Dim textList As Worksheet
    Dim CrmList As Worksheet
    Set textList = TargetWorkbook.Sheets("Text")
    Set CrmList = TargetWorkbook.Sheets("CRM")
    
    ' Naètení dat z Text listu do pole
    Dim textLastRow As Long
    Dim textData As Variant
    Dim textPromoIDColumn As Long
    
    textPromoIDColumn = textList.Range("tPromoID").Column
    textLastRow = textList.Cells(textList.rows.Count, textPromoIDColumn).End(xlUp).row
    
    If textLastRow < 3 Then Exit Sub
    
    textData = textList.Range(textList.Cells(3, textPromoIDColumn), _
                              textList.Cells(textLastRow, textPromoIDColumn)).value
    
    Dim rowsToDelete As Collection
    Set rowsToDelete = New Collection
    
    Dim i As Long
    For i = 1 To UBound(textData, 1)
        If CStr(textData(i, 1)) = searchValue Then
            rowsToDelete.Add i + 2
        End If
    Next i
    
    ' Zmìna statusu na CRM listu
    Dim cLastRow As Long
    Dim cData As Variant
    Dim cIDColumn As Long
    Dim cStatusColumn As Long
    
    cIDColumn = CrmList.Range("cIDakce").Column
    cStatusColumn = CrmList.Range("cStatus").Column
    cLastRow = CrmList.Cells(CrmList.rows.Count, cIDColumn).End(xlUp).row
    
    If cLastRow >= 1 Then
        cData = CrmList.Range(CrmList.Cells(1, cIDColumn), _
                              CrmList.Cells(cLastRow, cStatusColumn)).value
        
        Dim changed As Boolean
        changed = False
        
        For i = 1 To UBound(cData, 1)
            If CStr(cData(i, 1)) = searchValue Then
                cData(i, cStatusColumn - cIDColumn + 1) = "Cancelled"
                changed = True
            End If
        Next i
        
        If changed Then
            CrmList.Range(CrmList.Cells(1, cIDColumn), _
                         CrmList.Cells(cLastRow, cStatusColumn)).value = cData
        End If
    End If
    
    ' Smazání øádkù
    Dim textDeleteCount As Long
    textDeleteCount = 0
    
    If rowsToDelete.Count > 0 Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
        For i = rowsToDelete.Count To 1 Step -1
            textList.rows(rowsToDelete(i)).Delete
            textDeleteCount = textDeleteCount + 1
        Next i
        
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    End If
    
    ' Vyèištìní VŠECH bunìk se stejným PromoID v komentáøi
    Call ClearAllCellsWithPromoID(TargetWorkbook, searchValue)
    
End Sub

Private Sub ClearAllCellsWithPromoID(TargetWorkbook As Workbook, searchValue As String)
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim cell As Range
    Dim promoIDFromComment As String
    
    ' Projít VŠECHNY listy v sešitu
    For Each ws In TargetWorkbook.Worksheets
        ' Projít všechny buòky s komentáøem na aktuálním listu
        For Each cell In ws.Cells.SpecialCells(xlCellTypeComments)
            If Not cell.comment Is Nothing Then
                ' Získat prvních 8 znakù z komentáøe
                promoIDFromComment = Left(cell.comment.Text, 8)
                
                ' Pokud se shoduje s hledanou hodnotou
                If promoIDFromComment = searchValue Then
                    cell.ClearContents
                    cell.Interior.ColorIndex = 0
                    cell.ClearComments
                End If
            End If
        Next cell
    Next ws
    
    On Error GoTo 0
End Sub

' Volat na konci po zpracování všech PromoID
Public Sub FinalizeAfterDelete(TargetWorkbook As Workbook)
    Call SortIt(TargetWorkbook)
    Call rColor(TargetWorkbook)
    Call LockText(TargetWorkbook)
End Sub

