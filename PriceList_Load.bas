Attribute VB_Name = "PriceList_Load"
Public Sub LoadPriceList(TargetWorkbook As Workbook)
    Call UnlockPriceList(TargetWorkbook)
    
    Dim targetSheet As Worksheet
    Dim sourceRange As String
    Dim rgTarget As Range
    Dim rgFilter As Range
    Dim rgCol As Range
    Dim rgResult As Range
    Dim Criterion As Long
    
    Set targetSheet = TargetWorkbook.Sheets("PriceList")
    Set rgTarget = targetSheet.Range("A1:CR500")
    
    ' Definice bunìk pro naèítání nastavení
    Dim CellPath1 As String
    Dim CellPath2 As String
    Dim CellFile As String
    Dim CellSheetName As String
    CellPath1 = "B2"
    CellPath2 = "B3"
    CellFile = "B4"
    CellSheetName = "B5"
    
    ' Naète hodnoty z Settings
    Dim path1 As String
    Dim path2 As String
    Dim fileName As String
    Dim sheetName As String
    
    On Error Resume Next
    path1 = TargetWorkbook.Sheets("Settings").Range(CellPath1).value
    path2 = TargetWorkbook.Sheets("Settings").Range(CellPath2).value
    fileName = TargetWorkbook.Sheets("Settings").Range(CellFile).value
    sheetName = TargetWorkbook.Sheets("Settings").Range(CellSheetName).value
    On Error GoTo 0
    
    ' Kontrola cest - sestavení celé cesty
    Dim fullPath As String
    Dim basePath As String
    
    If Len(Dir(path1 & fileName)) > 0 Then
        basePath = path1
    ElseIf Len(Dir(path2 & fileName)) > 0 Then
        basePath = path2
    Else
        MsgBox "Ceník nebyl nalezen na žádné z uvedených cest."
        Exit Sub
    End If
    
    ' Sestavení sourceRange - SPRÁVNÝ FORMÁT
    sourceRange = "='" & basePath & "[" & fileName & "]" & sheetName & "'!$A$1:$CR$500"
    
    rgTarget.FormulaArray = sourceRange
    rgTarget.formula = rgTarget.value
    
    ' Pøidání názvù sloupcù do pojmenovaných rozsahù na øádek 4
    Dim ws As Worksheet
    Set ws = TargetWorkbook.Sheets("PriceList")
    
    On Error Resume Next
    ' Zapíše názvy do bunìk pojmenovaných rozsahù (øádek 4)
    ws.Range("family").Cells(4).value = "Family"
    ws.Range("CustomerID").Cells(4).value = "CustomerID"
    ws.Range("Brand").Cells(4).value = "Brand"
    On Error GoTo 0
    
    ' Filtrace a odstranìní øádkù
    Criterion = 0
    Set rgCol = targetSheet.Range("G4", targetSheet.Cells(targetSheet.rows.Count, "G").End(xlUp))
    rgCol.AutoFilter 1, Criterion
    
    On Error Resume Next
    Set rgFilter = rgCol.Offset(1).Resize(rgCol.rows.Count - 1).SpecialCells(xlCellTypeVisible)
    If Not rgFilter Is Nothing Then
        Set rgResult = rgFilter.EntireRow
        rgCol.Parent.AutoFilterMode = False
        rgResult.Delete
    Else
        rgCol.Parent.AutoFilterMode = False
    End If
    On Error GoTo 0
    
    ' Naètení dat do kolekce Dictionary
    Call ProductsArray(TargetWorkbook)
    Call FamilyArray(TargetWorkbook)
    
    ' Kód pro doplnìní nul
    Dim addZero(1 To 4) As String
    addZero(1) = "80686021902"
    addZero(2) = "80686021919"
    addZero(3) = "80686043492"
    addZero(4) = "87000652286"
    
    ' Získání kolekce produktù
    Dim productsCol As Collection
    Set productsCol = GetProductsCollection()
    
    ' Procházení dat pomocí kolekce
    Dim rowData As Object
    Dim currentEAN As String
    Dim originalEAN As String
    Dim needsZero As Boolean
    Dim j As Long
    Dim i As Long
    Dim familyValue As Variant
    Dim customerIDValue As Variant
    Dim brandValue As Variant
    
    Debug.Print "=== LoadPriceList Start ==="
    Debug.Print "Products collection count: " & productsCol.Count
    
    i = 5 ' Zaèínáme od øádku 5 (první datový øádek)
    For Each rowData In productsCol
        If rowData.Exists("ean") Then
            originalEAN = Trim(CStr(rowData("ean")))
            currentEAN = originalEAN
            
            ' Zkontroluje, zda EAN potøebuje pøidat nulu
            needsZero = False
            For j = LBound(addZero) To UBound(addZero)
                If currentEAN = addZero(j) Then
                    needsZero = True
                    Exit For
                End If
            Next j
            
            ' Formátování EAN v PriceList
            ws.Range("EAN")(i).NumberFormat = "@"
            If needsZero Then
                ws.Range("EAN")(i).value = "'0" & currentEAN
                ' Pro vyhledávání použijeme EAN s nulou
                currentEAN = "0" & currentEAN
            Else
                ws.Range("EAN")(i).value = "'" & currentEAN
            End If
            
            ' Aktualizuje nové sloupce pomocí vyhledávání v FamilyList podle EAN
            familyValue = GetFamilyByEAN(currentEAN)
            customerIDValue = GetCustomerIDByEAN(currentEAN)
            brandValue = GetBrandByEAN(currentEAN)
            
            ' Zápis hodnot do PriceList
            ws.Range("family")(i).value = familyValue
            ws.Range("CustomerID")(i).value = customerIDValue
            ws.Range("Brand")(i).value = brandValue
            
            ' Debug pro první 5 øádkù
            If i <= 9 Then
                Debug.Print "Row " & i & ": EAN=" & currentEAN & ", Family=" & familyValue & _
                           ", CustomerID=" & customerIDValue & ", Brand=" & brandValue
            End If
        Else
            If i <= 9 Then
                Debug.Print "Row " & i & ": ean column not found"
            End If
        End If
        
        i = i + 1
    Next rowData
    
    Debug.Print "=== LoadPriceList End: Processed " & (i - 5) & " rows ==="
    
    Call LockPriceList(TargetWorkbook)
End Sub


