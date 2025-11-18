Attribute VB_Name = "UpdatePrices"
' ===================================================================
' UpdatePrices_Shared - S automatickým přidáváním/mazáním produktů
' ===================================================================
Public Sub UpdatePrices_Shared(TargetWorkbook As Workbook, SelectedRange As Range, fcType As String, countryCode As String)
    On Error GoTo ErrorHandler
    
    Debug.Print "=== UpdatePrices_Shared START ==="

    ' Načíst countryCode z Settings
    countryCode = GetCountryCode(TargetWorkbook)

    If fcType = "" Then fcType = "AFC"  ' Default
    
    Debug.Print "FC Type: " & fcType
    Debug.Print "Country Code: " & countryCode
    
    ' 1. ODEMKNUT�
    Call UnlockText(TargetWorkbook)
    Call RemoveFilterIfApplied(TargetWorkbook)
    
    ' 2. KONTROLA V�BÍRU
    If SelectedRange Is Nothing Then
        MsgBox "Není vybrán žádné rozsah!", vbCritical
        Call LockText(TargetWorkbook)
        Exit Sub
    End If
    
    ' 3. NA�TENí DAT Z PRICELIST
    Call ProductsArray(TargetWorkbook)
    
    Dim productsCol As Collection
    Set productsCol = GetProductsCollection()
    
    If productsCol Is Nothing Or productsCol.Count = 0 Then
        MsgBox "Nepodařilo se načíst data z PriceList!", vbCritical
        Call LockText(TargetWorkbook)
        Exit Sub
    End If
    
    Debug.Print "Načteno produktů: " & productsCol.Count
    
    ' 4. REFERENCE NA LIST TEXT
    Dim textList As Worksheet
    Set textList = TargetWorkbook.Sheets("Text")
    
    ' 5. ZJIŠTĚNÍ VYBRANÝCH ŘÁDKŮ
    Dim rowsToUpdate As Collection
    Set rowsToUpdate = GetSelectedRows(SelectedRange, textList)
    
    If rowsToUpdate.Count = 0 Then
        MsgBox "Nebyly vybrány žádné platní řádky!", vbExclamation
        Call LockText(TargetWorkbook)
        Exit Sub
    End If
    
    Debug.Print "Požet řídkí k aktualizaci: " & rowsToUpdate.Count
    
   ' 6. SESKUPIT říDKY PODLE PromoID
    Dim promoGroups As Object
    Set promoGroups = GroupRowsByPromoID(textList, rowsToUpdate)
    
    Debug.Print "Požet unikítních promocí: " & promoGroups.Count

    ' 7. ZPRACOVAT KA�DOU PROMOCI
    Dim promoID As Variant
    Dim updatedCount As Long, addedCount As Long, deletedCount As Long, notFoundCount As Long
    updatedCount = 0
    addedCount = 0
    deletedCount = 0
    notFoundCount = 0
    
    For Each promoID In promoGroups.Keys
        Debug.Print "Zpracov�vým PromoID: " & promoID
        
        ' Pževést string řídkí na Collection
        Dim promoRows As Collection
        Set promoRows = StringToCollection(CStr(promoGroups(promoID)))
        
        Debug.Print "  Požet řídkí v této promoci: " & promoRows.Count
        
        ' ZMĚNA: Pžedat fcType a countryCode
        Dim result As Variant
        result = ProcessPromoGroup(textList, promoRows, productsCol, CStr(promoID), fcType, countryCode)
        
        updatedCount = updatedCount + result(0)
        addedCount = addedCount + result(1)
        deletedCount = deletedCount + result(2)
        notFoundCount = notFoundCount + result(3)
    Next promoID
        
    ' 8. FIN�LNí ÚPRAVY
    Call ApplyFilterToRow2(TargetWorkbook)
    Call SortIt(TargetWorkbook)
    Call rColor(TargetWorkbook)
    Call LockText(TargetWorkbook)
        
    Debug.Print "=== UpdatePrices_Shared END ==="
    MsgBox "Aktualizace dokonžena:" & vbCrLf & _
           "- Aktualizováno: " & updatedCount & " řídk�" & vbCrLf & _
           "- Přidáno: " & addedCount & " řídk�" & vbCrLf & _
           "- Smazáno: " & deletedCount & " řídk�" & vbCrLf & _
           "- Nenalezeno v PriceList: " & notFoundCount & " řídk�", vbInformation
    Exit Sub
    
ErrorHandler:
    Debug.Print "CHYBA: " & Err.Description
    MsgBox "Chyba: " & Err.Description, vbCritical
    Call LockText(TargetWorkbook)
End Sub

' ===================================================================
' GroupRowsByPromoID - S explicitním zacházením s typy
' ===================================================================
Private Function GroupRowsByPromoID(textList As Worksheet, rowsToUpdate As Collection) As Object
    Set GroupRowsByPromoID = CreateObject("Scripting.Dictionary")
    
    Dim promoIDCol As Long
    promoIDCol = GetColumnSafe(textList, "tPromoID")
    
    Dim rowNum As Variant
    Dim promoID As String
    Dim promoIDKey As Variant  ' Pro použití jako klíč v Dictionary
    Dim currentValue As Variant
    
    For Each rowNum In rowsToUpdate
        promoID = Trim(CStr(textList.Cells(CLng(rowNum), promoIDCol).value))
        
        If promoID <> "" Then
            promoIDKey = promoID  ' Pževod na Variant
            
            If Not GroupRowsByPromoID.Exists(promoIDKey) Then
                GroupRowsByPromoID.Add promoIDKey, CStr(rowNum)
            Else
                currentValue = GroupRowsByPromoID.Item(promoIDKey)
                GroupRowsByPromoID.Item(promoIDKey) = CStr(currentValue) & "," & CStr(rowNum)
            End If
        End If
    Next rowNum
End Function

' ===================================================================
' Pomocná funkce - Pževod stringu na Collection
' ===================================================================
Private Function StringToCollection(rowsString As String) As Collection
    Set StringToCollection = New Collection
    
    Dim rowsArray() As String
    rowsArray = Split(rowsString, ",")
    
    Dim i As Long
    For i = LBound(rowsArray) To UBound(rowsArray)
        StringToCollection.Add CLng(Trim(rowsArray(i)))
    Next i
End Function

' ===================================================================
' Zpracování jední promoce (skupina řídkí se stejným PromoID)
' ===================================================================
Private Function ProcessPromoGroup(textList As Worksheet, promoRows As Collection, productsCol As Collection, promoID As String, fcType As String, countryCode As String) As Variant
    Dim updatedCount As Long, addedCount As Long, deletedCount As Long, notFoundCount As Long
    updatedCount = 0
    addedCount = 0
    deletedCount = 0
    notFoundCount = 0
    
    ' 1. ZÍSKAT FAMILY a tVyber z prvního řádku
    Dim firstRow As Long
    firstRow = promoRows(1)
    
    Dim familyValue As String
    familyValue = Trim(textList.Cells(firstRow, GetColumnSafe(textList, "tFamily")).value)
    
    Dim vyberValue As String
    vyberValue = Trim(textList.Cells(firstRow, GetColumnSafe(textList, "tVyber")).value)
    
    Debug.Print "  Family: " & familyValue & ", tVyber: " & vyberValue
    
    ' 2. NAJÍT V�ECHNY PRODUKTY V T�TO FAMILY V PRICELIST
    Dim familyProducts As Collection
    Set familyProducts = GetProductsByFamily(productsCol, familyValue)
    
    Debug.Print "  Produktí v PriceList pro Family " & familyValue & ": " & familyProducts.Count
    Debug.Print "  Produktí v Text: " & promoRows.Count
    
    ' 3. ZÍSKAT EXISTUJÍCÍ PRODUKTY V PROMOCI
    Dim existingProducts As Object  ' Dictionary: productName -> rowNumber
    Set existingProducts = CreateObject("Scripting.Dictionary")
    
    Dim rowNum As Variant
    For Each rowNum In promoRows
        Dim productName As String
        productName = Trim(textList.Cells(CLng(rowNum), GetColumnSafe(textList, "tProduct")).value)
        If productName <> "" Then
            existingProducts.Add productName, CLng(rowNum)
        End If
    Next rowNum
    
    ' 4. AKTUALIZOVAT EXISTUJÍCÍ + NAJÍT PRODUKTY K PŘIDÁNÍ/SMAZÁNÍ
    Dim productsToAdd As Collection
    Set productsToAdd = New Collection
    
    Dim productsToDelete As Collection
    Set productsToDelete = New Collection
    
    ' Projít produkty v PriceList
' Projít produkty v PriceList
Dim productRow As Object
For Each productRow In familyProducts
    Dim productFullName As String
    
    ' Sestavit productFullName podle countryCode
    If UCase(Trim(countryCode)) = "SVK" Then
        productFullName = productRow("material_name")
    Else
        productFullName = productRow("material_name") & " " & productRow("volume_l")
    End If
    
    Debug.Print "    Kontroluji produkt z PriceList: " & productFullName
    
    If existingProducts.Exists(productFullName) Then
        ' Produkt existuje í AKTUALIZOVAT
        Debug.Print "      Existuje v Text - AKTUALIZUJI"
        Dim targetRow As Long
        targetRow = existingProducts(productFullName)
        
        If UpdateSingleRow(textList, targetRow, productsCol, countryCode) Then
            updatedCount = updatedCount + 1
        Else
            notFoundCount = notFoundCount + 1
        End If
        
        ' Odebrat z existingProducts
        existingProducts.Remove productFullName
    Else
        ' Produkt NEEXISTUJE v Text
        Debug.Print "      Neexistuje v Text"
        Debug.Print "      tVyber: " & vyberValue
        
        If UCase(vyberValue) = "N" Then
            Debug.Print "      >>> PŘIDÁM (tVyber = N)"
            productsToAdd.Add productRow
        Else
            Debug.Print "      Nepřidávým (tVyber = " & vyberValue & ")"
        End If
    End If
Next productRow

Debug.Print "  Produkty k přidání: " & productsToAdd.Count
    
    ' Co zbylo v existingProducts = produkty K SMAZÁNÍ (nejsou v PriceList)
    Dim productToDelete As Variant
    For Each productToDelete In existingProducts.Keys
        productsToDelete.Add existingProducts(productToDelete)
    Next productToDelete
    
    ' 5. PŘIDAT NOVÉ PRODUKTY (pouze pokud tVyber = "N")
    If UCase(vyberValue) = "N" And productsToAdd.Count > 0 Then
        Debug.Print "  Přidávám " & productsToAdd.Count & " nových produktů..."
        addedCount = AddNewProducts(textList, productsToAdd, firstRow, promoID, countryCode)
    End If
    
    ' 6. SMAZAT NEEXISTUJÍCÍ PRODUKTY
    If productsToDelete.Count > 0 Then
        Debug.Print "  Mažu " & productsToDelete.Count & " neexistujících produktů..."
        deletedCount = DeleteProducts(textList, productsToDelete)
    End If
    
    ' Vrítit statistiky
    ProcessPromoGroup = Array(updatedCount, addedCount, deletedCount, notFoundCount)
End Function

' ===================================================================
' Získání všech produktů z dané Family
' ===================================================================
Private Function GetProductsByFamily(productsCol As Collection, familyValue As String) As Collection
    Set GetProductsByFamily = New Collection
    
    Dim rowData As Object
    For Each rowData In productsCol
        If Trim(rowData("Family")) = Trim(familyValue) Then
            GetProductsByFamily.Add rowData
        End If
    Next rowData
End Function

' ===================================================================
' Přidání nových produktů
' ===================================================================
Private Function AddNewProducts(textList As Worksheet, productsToAdd As Collection, templateRow As Long, promoID As String, countryCode As String) As Long
    AddNewProducts = 0
    
    ' OPRAVA: Načíst fcType z tFCtype se správnou kontrolou
    Dim fcType As String
    Dim fcTypeCol As Long
    fcTypeCol = GetColumnSafe(textList, "tFCtype")
    
    Debug.Print "  Sloupec tFCtype: " & fcTypeCol
    
    If fcTypeCol > 0 Then
        fcType = Trim(CStr(textList.Cells(templateRow, fcTypeCol).value))
        Debug.Print "  Hodnota v templateRow " & templateRow & ": '" & fcType & "'"
    Else
        fcType = ""
        Debug.Print "  VAROVÁNÍ�: Sloupec tFCtype nenalezen!"
    End If
    
    ' Default hodnota pokud je prázdné
    If fcType = "" Or fcType = "0" Then
        fcType = "AFC"
        Debug.Print "  Použit default fcType: AFC"
    End If
    
    Debug.Print "  fcType pro noví produkty: " & fcType
    
    ' Najít poslední řádek
    Dim lastRow As Long
    lastRow = textList.Cells(textList.rows.Count, GetColumnSafe(textList, "tProduct")).End(xlUp).row
    
    Dim newRow As Long
    newRow = lastRow + 1
    
    Dim productRow As Object
    For Each productRow In productsToAdd
        ' Zkopírovat formát z templateRow
        textList.rows(templateRow).Copy
        textList.rows(newRow).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        
        ' Sestavit productName podle countryCode
        Dim productName As String
        If UCase(Trim(countryCode)) = "SVK" Then
            productName = productRow("material_name")
        Else
            productName = productRow("material_name") & " " & productRow("volume_l")
        End If
        
        Debug.Print "    Přidávám: " & productName
        
        ' Základní údaje
        Call WriteToColumnSafe(textList, newRow, "tProduct", productName)
        Call WriteToColumnSafe(textList, newRow, "tCustomerID", productRow("CustomerID"))
        Call WriteToColumnSafe(textList, newRow, "tEAN", "'" & productRow("ean"))
        Call WriteToColumnSafe(textList, newRow, "tPackageSize", productRow("volume_l"))
        Call WriteToColumnSafe(textList, newRow, "tStockID", productRow("sap_id"))
        Call WriteToColumnSafe(textList, newRow, "tBrand", productRow("Brand"))
        Call WriteToColumnSafe(textList, newRow, "tFamily", productRow("Family"))
        Call WriteToColumnSafe(textList, newRow, "tCategory", productRow("category"))
        Call WriteToColumnSafe(textList, newRow, "tPromoID", promoID)
        
        ' Ceny - získat z GetPromoPriceData
        Dim familyValue As String
        familyValue = Trim(textList.Cells(templateRow, GetColumnSafe(textList, "tFamily")).value)
        
        Dim priceType As String
        priceType = Trim(textList.Cells(templateRow, GetColumnSafe(textList, "tPriceType")).value)
        If priceType = "" Then priceType = "ANCD"
        
        ' Vol�ní GetPromoPriceData s fcType
        Dim result As Variant
        result = GetPromoPriceData(familyValue, priceType, productRow, fcType)
        
        Call WriteToColumnSafe(textList, newRow, "tPromoPrice", result(0))
        Call WriteToColumnSafe(textList, newRow, "tPriceType", result(1))
        Call WriteToColumnSafe(textList, newRow, "tZStype", result(2))
        Call WriteToColumnSafe(textList, newRow, "tAFC", result(3))
        Call WriteToColumnSafe(textList, newRow, "tKomp", result(4))
        Call WriteToColumnSafe(textList, newRow, "tC1l", result(5))
        Call WriteToColumnSafe(textList, newRow, "tZS", result(6))
        Call WriteToColumnSafe(textList, newRow, "tPriorita", result(7))
        Call WriteToColumnSafe(textList, newRow, "tFCtype", result(8))
        
        Call WriteToColumnSafe(textList, newRow, "tFC", productRow("ncd_invoice"))
        Call WriteToColumnSafe(textList, newRow, "tNCD", productRow("ncd_inc_vat"))
        
        ' Zkopírovat další údaje z templateRow (datumy, týdny, atd.)
        Call CopyPromotionData(textList, templateRow, newRow)
        
        ' Zkopírovat tDiff, tVol, tOfftakeTotal, tC1Total z JINÉHO řádku ve Family
        Call CopyFamilySpecificData(textList, newRow, familyValue, promoID)
        
        AddNewProducts = AddNewProducts + 1
        newRow = newRow + 1
    Next productRow
End Function

' ===================================================================
' Kopírování dat specifických pro Family - S VZORCI
' ===================================================================
Private Sub CopyFamilySpecificData(textList As Worksheet, targetRow As Long, familyValue As String, promoID As String)
    On Error Resume Next
    
    Debug.Print "      Kopářuji Family-specific data pro řádek " & targetRow
    
    ' Najít jiní řádek se stejnou Family a PromoID (ale ne targetRow)
    Dim sourceRow As Long
    sourceRow = FindFamilySourceRow(textList, familyValue, promoID, targetRow)
    
    If sourceRow = 0 Then
        Debug.Print "      VAROVÁNÍ�: Nenalezen zdrojoví řádek pro Family " & familyValue
        Exit Sub
    End If
    
    Debug.Print "      Zdrojoví řádek: " & sourceRow
    
    ' Zkopírovat VZORCE (ne hodnoty)
    Call CopyFormulaToRow(textList, sourceRow, targetRow, "tDiff")
    Call CopyFormulaToRow(textList, sourceRow, targetRow, "tVol")
    Call CopyFormulaToRow(textList, sourceRow, targetRow, "tOfftakeTotal")
    Call CopyFormulaToRow(textList, sourceRow, targetRow, "tC1Total")
    
    ' tPriceType je hodnota (ne vzorec), kopářovat normální
    Call WriteToColumnSafe(textList, targetRow, "tPriceType", _
        textList.Cells(sourceRow, GetColumnSafe(textList, "tPriceType")).value)
    
    Debug.Print "      Family data zkopářována"
    
    On Error GoTo 0
End Sub

' ===================================================================
' Kopírování vzorce s úpravou řísla řádku
' ===================================================================
Private Sub CopyFormulaToRow(ws As Worksheet, sourceRow As Long, targetRow As Long, rangeName As String)
    On Error Resume Next
    
    Dim sourceCol As Long
    sourceCol = GetColumnSafe(ws, rangeName)
    
    If sourceCol = 0 Or sourceCol = 1 Then
        Debug.Print "        VAROVÁNÍ�: Sloupec " & rangeName & " nebyl nalezen"
        Exit Sub
    End If
    
    Dim sourceCell As Range
    Set sourceCell = ws.Cells(sourceRow, sourceCol)
    
    Dim targetCell As Range
    Set targetCell = ws.Cells(targetRow, sourceCol)
    
    ' Zkontrolovat, zda zdrojoví buňka obsahuje vzorec
    If sourceCell.HasFormula Then
        ' Zkopírovat vzorec a upravit říslo řádku
        Dim originalFormula As String
        Dim adjustedFormula As String
        
        originalFormula = sourceCell.formula
        
        ' Nahradit všechny výskyty zdrojovího řádku cílovým řídkem
        adjustedFormula = ReplaceRowReferences(originalFormula, sourceRow, targetRow)
        
        targetCell.formula = adjustedFormula
        Debug.Print "        " & rangeName & ": Zkopářován a upraven vzorec"
        Debug.Print "          P�vodní: " & originalFormula
        Debug.Print "          Nov�: " & adjustedFormula
    Else
        ' Pokud není vzorec, zkopírovat hodnotu
        targetCell.value = sourceCell.value
        Debug.Print "        " & rangeName & ": Zkopářována hodnota (žádné vzorec)"
    End If
    
    On Error GoTo 0
End Sub

' ===================================================================
' Nahrazení odkazí na řádky ve vzorci
' ===================================================================
Private Function ReplaceRowReferences(formula As String, oldRow As Long, newRow As Long) As String
    Dim result As String
    result = formula
    
    ' Nahradit absolutní odkazy: $A$10 í $A$50
    result = Replace(result, "$" & oldRow, "$" & newRow)
    
    ' Nahradit relativní odkazy: A10 í A50
    ' Projít všechny písmena sloupců (A-Z, AA-ZZ)
    Dim col As String
    Dim i As Long
    
    ' Jednop�smenní sloupce (A-Z)
    For i = 65 To 90  ' ASCII A-Z
        col = Chr(i)
        result = Replace(result, col & oldRow, col & newRow)
        result = Replace(result, col & "$" & oldRow, col & "$" & newRow)
    Next i
    
    ' Dvoum�stní sloupce (AA-AZ, BA-BZ, CA-CZ)
    Dim firstChar As String, secondChar As String
    For i = 65 To 90
        firstChar = Chr(i)
        For j = 65 To 90
            secondChar = Chr(j)
            col = firstChar & secondChar
            result = Replace(result, col & oldRow, col & newRow)
            result = Replace(result, col & "$" & oldRow, col & "$" & newRow)
        Next j
    Next i
    
    ReplaceRowReferences = result
End Function

' ===================================================================
' Najít zdrojoví řádek se stejnou Family a PromoID
' ===================================================================
Private Function FindFamilySourceRow(textList As Worksheet, familyValue As String, promoID As String, excludeRow As Long) As Long
    FindFamilySourceRow = 0
    
    Dim lastRow As Long
    lastRow = textList.Cells(textList.rows.Count, GetColumnSafe(textList, "tProduct")).End(xlUp).row
    
    Dim familyCol As Long, promoIDCol As Long
    familyCol = GetColumnSafe(textList, "tFamily")
    promoIDCol = GetColumnSafe(textList, "tPromoID")
    
    Dim i As Long
    For i = 3 To lastRow  ' Za�ít od řádku 3 (pžeskočit header)
        If i <> excludeRow Then  ' Pžeskočit cíloví řádek
            If Trim(textList.Cells(i, familyCol).value) = Trim(familyValue) And _
               Trim(CStr(textList.Cells(i, promoIDCol).value)) = Trim(promoID) Then
                ' Na�li jsme vhodní řádek
                FindFamilySourceRow = i
                Exit Function
            End If
        End If
    Next i
End Function

' ===================================================================
' Kopírování promoce dat z template řádku
' ===================================================================
Private Sub CopyPromotionData(textList As Worksheet, fromRow As Long, toRow As Long)
    ' Zkopírovat datumy a týdny
    Call WriteToColumnSafe(textList, toRow, "tAkceOd", textList.Cells(fromRow, GetColumnSafe(textList, "tAkceOd")).value)
    Call WriteToColumnSafe(textList, toRow, "tAkceDo", textList.Cells(fromRow, GetColumnSafe(textList, "tAkceDo")).value)
    Call WriteToColumnSafe(textList, toRow, "tNakupOd", textList.Cells(fromRow, GetColumnSafe(textList, "tNakupOd")).value)
    Call WriteToColumnSafe(textList, toRow, "tNakupDo", textList.Cells(fromRow, GetColumnSafe(textList, "tNakupDo")).value)
    Call WriteToColumnSafe(textList, toRow, "tSortFrom", textList.Cells(fromRow, GetColumnSafe(textList, "tSortFrom")).value)
    Call WriteToColumnSafe(textList, toRow, "tSortTo", textList.Cells(fromRow, GetColumnSafe(textList, "tSortTo")).value)
    Call WriteToColumnSafe(textList, toRow, "tWeeks", textList.Cells(fromRow, GetColumnSafe(textList, "tWeeks")).value)
    Call WriteToColumnSafe(textList, toRow, "tWeeksT", textList.Cells(fromRow, GetColumnSafe(textList, "tWeeksT")).value)
    Call WriteToColumnSafe(textList, toRow, "tTypAkce", textList.Cells(fromRow, GetColumnSafe(textList, "tTypAkce")).value)
    Call WriteToColumnSafe(textList, toRow, "tPromo", textList.Cells(fromRow, GetColumnSafe(textList, "tPromo")).value)
    Call WriteToColumnSafe(textList, toRow, "tCom", textList.Cells(fromRow, GetColumnSafe(textList, "tCom")).value)
    Call WriteToColumnSafe(textList, toRow, "tVyber", textList.Cells(fromRow, GetColumnSafe(textList, "tVyber")).value)
    Call WriteToColumnSafe(textList, toRow, "tCustomer", textList.Cells(fromRow, GetColumnSafe(textList, "tCustomer")).value)
    Call WriteToColumnSafe(textList, toRow, "tHero", "N")  ' Noví produkt není hero
    Call WriteToColumnSafe(textList, toRow, "tPotvrzeno", textList.Cells(fromRow, GetColumnSafe(textList, "tPotvrzeno")).value)
    Call WriteToColumnSafe(textList, toRow, "tFCname", textList.Cells(fromRow, GetColumnSafe(textList, "tFCname")).value)
End Sub

' ===================================================================
' Smazání produktů
' ===================================================================
Private Function DeleteProducts(textList As Worksheet, rowsToDelete As Collection) As Long
    DeleteProducts = 0
    
    ' Se�adit řádky od nejvyříího k nejniříímu (aby se při mazání neposouvaly indexy)
    Dim sortedRows As Collection
    Set sortedRows = SortRowsDescending(rowsToDelete)
    
    Dim rowNum As Variant
    For Each rowNum In sortedRows
        Debug.Print "    Mažu řádek: " & rowNum
        textList.rows(CLng(rowNum)).Delete
        DeleteProducts = DeleteProducts + 1
    Next rowNum
End Function

' ===================================================================
' Seřazení řídkí sestupně
' ===================================================================
Private Function SortRowsDescending(rows As Collection) As Collection
    Set SortRowsDescending = New Collection
    
    ' Pževést na array a seřadit
    Dim arr() As Long
    ReDim arr(1 To rows.Count)
    
    Dim i As Long
    For i = 1 To rows.Count
        arr(i) = rows(i)
    Next i
    
    ' Bubble sort (sestupně)
    Dim j As Long, temp As Long
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    ' Pževést zpět do Collection
    For i = 1 To UBound(arr)
        SortRowsDescending.Add arr(i)
    Next i
End Function

' ===================================================================
' PŮVODNÍ FUNKCE (beze změn)
' ===================================================================

' GetSelectedRows - beze změn
Private Function GetSelectedRows(SelectedRange As Range, textList As Worksheet) As Collection
    Set GetSelectedRows = New Collection
    
    Dim cell As Range
    Dim rowNum As Long
    Dim addedRows As Object
    Set addedRows = CreateObject("Scripting.Dictionary")
    
    For Each cell In SelectedRange
        If cell.Worksheet.Name = textList.Name Then
            rowNum = cell.row
            If rowNum >= 3 And Not addedRows.Exists(rowNum) Then
                If Trim(textList.Cells(rowNum, GetColumnSafe(textList, "tProduct")).value) <> "" Then
                    GetSelectedRows.Add rowNum
                    addedRows.Add rowNum, True
                End If
            End If
        End If
    Next cell
End Function

' UpdateSingleRow
Private Function UpdateSingleRow(textList As Worksheet, rowNum As Long, productsCol As Collection, countryCode As String) As Boolean
    On Error GoTo ErrorHandler
    UpdateSingleRow = False
    
    Dim productName As String
    productName = textList.Cells(rowNum, GetColumnSafe(textList, "tProduct")).value
    If Trim(productName) = "" Then Exit Function
    
    ' PŘIDÁNO: Načíst fcType z řádku
    Dim fcType As String
    fcType = Trim(textList.Cells(rowNum, GetColumnSafe(textList, "tFCname")).value)
    Debug.Print "    fcType pro řádek " & rowNum & ": " & fcType
    
    Dim priceType As String
    priceType = Trim(textList.Cells(rowNum, GetColumnSafe(textList, "tPriceType")).value)
    If priceType = "" Then priceType = "ANCD"
    
    Dim familyValue As String
    familyValue = Trim(textList.Cells(rowNum, GetColumnSafe(textList, "tFamily")).value)
    
    Dim productRow As Object
    Set productRow = FindProductInCollection(productsCol, productName, countryCode)
    If productRow Is Nothing Then Exit Function
    
    ' ZMĚNA: GetPromoPriceData s fcType
    Dim result As Variant
    result = GetPromoPriceData(familyValue, priceType, productRow, fcType)
    
    Call WriteToColumnSafe(textList, rowNum, "tPromoPrice", result(0))
    Call WriteToColumnSafe(textList, rowNum, "tZStype", result(2))
    Call WriteToColumnSafe(textList, rowNum, "tAFC", result(3))
    Call WriteToColumnSafe(textList, rowNum, "tKomp", result(4))
    Call WriteToColumnSafe(textList, rowNum, "tC1l", result(5))
    Call WriteToColumnSafe(textList, rowNum, "tZS", result(6))
    Call WriteToColumnSafe(textList, rowNum, "tPriorita", result(7))
    Call WriteToColumnSafe(textList, rowNum, "tFCtype", result(8))
    
    Call WriteToColumnSafe(textList, rowNum, "tFC", productRow("ncd_invoice"))
    Call WriteToColumnSafe(textList, rowNum, "tNCD", productRow("ncd_inc_vat"))
    Call WriteToColumnSafe(textList, rowNum, "tCustomerID", productRow("CustomerID"))
    Call WriteToColumnSafe(textList, rowNum, "tEAN", "'" & productRow("ean"))
    Call WriteToColumnSafe(textList, rowNum, "tStockID", productRow("sap_id"))
    Call WriteToColumnSafe(textList, rowNum, "tBrand", productRow("Brand"))
    
    UpdateSingleRow = True
    Exit Function
    
ErrorHandler:
    UpdateSingleRow = False
End Function

' FindProductInCollection - beze změn
Private Function FindProductInCollection(productsCol As Collection, productName As String, countryCode As String) As Object
    Set FindProductInCollection = Nothing
    Dim rowData As Object
    For Each rowData In productsCol
        ' Sestavit productName podle countryCode
        Dim fullName As String
        If UCase(Trim(countryCode)) = "SVK" Then
            fullName = Trim(rowData("material_name"))
        Else
            fullName = Trim(rowData("material_name") & " " & rowData("volume_l"))
        End If
        
        If fullName = Trim(productName) Then
            Set FindProductInCollection = rowData
            Exit Function
        End If
    Next rowData
End Function
