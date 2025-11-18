Attribute VB_Name = "Module1"
Public Function ValidateRequiredSelections(promoceIndex As Long, priceIndex As Long) As Boolean 'Kontrola, �e je v�e vybr�no v listboxech
    If promoceIndex = -1 Or priceIndex = -1 Then
        MsgBox "Nejsou vybr�ny v�echny povinn� �daje."
        ValidateRequiredSelections = False
    Else
        ValidateRequiredSelections = True
    End If
End Function

Public Sub CopySelectedProductsToHero(productList As Variant, SelectedItems As Variant, ByRef heroList As Variant, promoceIndex As Long, priceIndex As Long)
    ' Kontrola povinn�ch �daj�
    If promoceIndex = -1 Or priceIndex = -1 Then
        MsgBox "Nejsou vybr�ny v�echny povinn� �daje."
        Exit Sub
    End If
    
    ' Vytvo�� seznam hero produkt�
    Dim heroArray() As String
    Dim heroCount As Long
    Dim i As Long
    
    ' Spo��t� vybran� produkty
    For i = 0 To UBound(SelectedItems)
        If SelectedItems(i) = True Then
            heroCount = heroCount + 1
        End If
    Next i
    
    If heroCount > 0 Then
        ReDim heroArray(0 To heroCount - 1)
        Dim index As Long
        For i = 0 To UBound(SelectedItems)
            If SelectedItems(i) = True Then
                heroArray(index) = productList(i)
                index = index + 1
            End If
        Next i
    End If
    
    heroList = heroArray
End Sub

' Pomocn� funkce pro bezpe�n� z�sk�n� sloupce
Public Function GetColumnSafe(ws As Worksheet, rangeName As String) As Long
    On Error Resume Next
    GetColumnSafe = ws.Range(rangeName).Column
    If Err.Number <> 0 Then
        GetColumnSafe = 1 ' V�choz� sloupec A, pokud range neexistuje
        Err.Clear
    End If
    On Error GoTo 0
End Function

' Pomocn� funkce pro bezpe�n� z�pis do sloupce
Public Sub WriteToColumnSafe(ws As Worksheet, row As Long, rangeName As String, value As Variant)
    On Error Resume Next
    Dim col As Long
    col = ws.Range(rangeName).Column
    If Err.Number = 0 Then
        ws.Cells(row, col).value = value
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Public Function GetCustomer(TargetWorkbook As Workbook) As String
    On Error Resume Next
    GetCustomer = TargetWorkbook.Sheets("Settings").Range("B1").value
    If Err.Number <> 0 Then
        GetCustomer = ""
        Err.Clear
    End If
    On Error GoTo 0
End Function

' Z�sk� countryCode z Settings listu
Public Function GetCountryCode(TargetWorkbook As Workbook) As String
    On Error Resume Next
    GetCountryCode = Trim(TargetWorkbook.Sheets("Settings").Range("B10").value)
    On Error GoTo 0

    If GetCountryCode = "" Then
        GetCountryCode = "CZK"  ' Default
    End If
End Function
' Zjist� vybran� polo�ky v listboxu
Public Function GetSelectedItems(listBox As Object) As Variant
    ' Spo��t� vybran� polo�ky
    Dim selectedCount As Long
    Dim i As Long
    
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) = True Then
            selectedCount = selectedCount + 1
        End If
    Next i
    
    ' Pokud nic nen� vybr�no, vr�t� pr�zdn� pole
    If selectedCount = 0 Then
        GetSelectedItems = Array()  ' Pr�zdn� pole
        Exit Function
    End If
    
    ' Vytvo�� pole a napln� ho
    Dim SelectedItems() As String
    ReDim SelectedItems(0 To selectedCount - 1)
    
    Dim index As Long
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) = True Then
            SelectedItems(index) = listBox.List(i)
            index = index + 1
        End If
    Next i
    
    GetSelectedItems = SelectedItems
End Function
' Hero
Public Function GetHeroItem(listBox As Object) As String
    Dim i As Long
    For i = 0 To listBox.ListCount - 1
        If listBox.Selected(i) = True Then
            GetHeroItem = listBox.List(i)
            Exit Function
        End If
    Next i
    GetHeroItem = ""
End Function
' V�b�r
Public Function GetVyberValue(listBox As Object) As String
    Dim allSelected As Boolean
    allSelected = True
    
    Dim i As Long
    For i = 0 To listBox.ListCount - 1
        If Not listBox.Selected(i) Then
            allSelected = False
            Exit For
        End If
    Next i
    
    If allSelected Then
        GetVyberValue = "N"
    Else
        GetVyberValue = "A"
    End If
End Function
Public Function CreatePromoInstance() As Promo
    Set CreatePromoInstance = New Promo
End Function

Public Function GetPromoPriceData(familyValue As String, selectedPrice As String, productRow As Object, fcType As String) As Variant
    Dim promoValue As Variant
    Dim promoName As String
    Dim zsName As String
    Dim AFCvalue As Variant
    Dim kompValue As Variant
    Dim c1Value As Variant
    Dim ZSvalue As Variant
    Dim Priorita As String
    Dim FCname As String
    
    ' Zkontroluje, zda produkt odpov�d� rodin�
    If familyValue = productRow("Family") Then
        
        ' ZM�NA: Nejd��v zkontrolovat, jestli je to FC
        If UCase(Trim(fcType)) = "FC" Then
            ' Pro FC vr�tit jen z�kladn� hodnoty
            promoValue = productRow("ncd_inc_vat")
            promoName = selectedPrice  ' Pou��t vybranou cenu (ANCD, TANCD...)
            zsName = ""
            AFCvalue = ""
            kompValue = ""
            c1Value = ""
            ZSvalue = ""
            Priorita = "Standard"
            FCname = "FC"
            
        Else
            
            FCname = "AFC" 'fcType
            
            Select Case selectedPrice
                Case "ANCD"
                    promoValue = productRow("ancd_inc_vat")
                    promoName = "ANCD"
                    zsName = "ZSANCD"
                    AFCvalue = productRow("ancd_invoice")
                    kompValue = productRow("ancd_comp_tcogs_czk_pc")
                    c1Value = productRow("promo_c1_l")
                    ZSvalue = productRow("ancd_rebates")
                    Priorita = "Standard"
                    
                Case "TANCD"
                    If productRow("tancd1_inc_vat") <> 0 Then
                        promoValue = productRow("tancd1_inc_vat")
                        promoName = "TANCD"
                        zsName = "ZSTANCD"
                        AFCvalue = productRow("tancd1_invoice")
                        kompValue = productRow("tancd1_comp_tcogs_czk_pc")
                        c1Value = productRow("tancd1_c1_l")
                        ZSvalue = productRow("tancd1_rebate")
                        Priorita = "Taktick�"
                    Else
                        promoValue = productRow("ancd_inc_vat")
                        promoName = "ANCD"
                        zsName = "ZSANCD"
                        AFCvalue = productRow("ancd_invoice")
                        kompValue = productRow("ancd_comp_tcogs_czk_pc")
                        c1Value = productRow("promo_c1_l")
                        ZSvalue = productRow("ancd_rebates")
                        Priorita = "Standard"
                    End If
                    
                Case "TANCD II", "TANCDII"
                    If productRow("tancd2_inc_vat") <> 0 Then
                        promoValue = productRow("tancd2_inc_vat")
                        promoName = "TANCDII"
                        zsName = "ZSTANCDII"
                        AFCvalue = productRow("tancd2_invoice")
                        kompValue = productRow("tancd2_comp_tcogs2_czk_pc")
                        c1Value = productRow("tancd2_c1_l")
                        ZSvalue = productRow("tancd2_rebate")
                        Priorita = "Taktick�"
                    ElseIf productRow("tancd1_inc_vat") <> 0 Then
                        promoValue = productRow("tancd1_inc_vat")
                        promoName = "TANCD"
                        zsName = "ZSTANCD"
                        AFCvalue = productRow("tancd1_invoice")
                        kompValue = productRow("tancd1_comp_tcogs_czk_pc")
                        c1Value = productRow("tancd1_c1_l")
                        ZSvalue = productRow("tancd1_rebate")
                        Priorita = "Taktick�"
                    Else
                        promoValue = productRow("ancd_inc_vat")
                        promoName = "ANCD"
                        zsName = "ZSANCD"
                        AFCvalue = productRow("ancd_invoice")
                        kompValue = productRow("ancd_comp_tcogs_czk_pc")
                        c1Value = productRow("promo_c1_l")
                        ZSvalue = productRow("ancd_rebates")
                        Priorita = "Standard"
                    End If
                    
                Case "TANCD III", "TANCDIII"
                    If productRow("tancd3_inc_vat") <> 0 Then
                        promoValue = productRow("tancd3_inc_vat")
                        promoName = "TANCDIII"
                        zsName = "ZSTANCDIII"
                        AFCvalue = productRow("tancd3_invoice")
                        kompValue = productRow("tancd3_comp_tcogs3_czk_pc")
                        c1Value = productRow("tancd3_c1_l")
                        ZSvalue = productRow("tancd3_rebate")
                        Priorita = "Taktick�"
                    ElseIf productRow("tancd2_inc_vat") <> 0 Then
                        promoValue = productRow("tancd2_inc_vat")
                        promoName = "TANCDII"
                        zsName = "ZSTANCDII"
                        AFCvalue = productRow("tancd2_invoice")
                        kompValue = productRow("tancd2_comp_tcogs2_czk_pc")
                        c1Value = productRow("tancd2_c1_l")
                        ZSvalue = productRow("tancd2_rebate")
                        Priorita = "Taktick�"
                    ElseIf productRow("tancd1_inc_vat") <> 0 Then
                        promoValue = productRow("tancd1_inc_vat")
                        promoName = "TANCD"
                        zsName = "ZSTANCD"
                        AFCvalue = productRow("tancd1_invoice")
                        kompValue = productRow("tancd1_comp_tcogs_czk_pc")
                        c1Value = productRow("tancd1_c1_l")
                        ZSvalue = productRow("tancd1_rebate")
                        Priorita = "Taktick�"
                    Else
                        promoValue = productRow("ancd_inc_vat")
                        promoName = "ANCD"
                        zsName = "ZSANCD"
                        AFCvalue = productRow("ancd_invoice")
                        kompValue = productRow("ancd_comp_tcogs_czk_pc")
                        c1Value = productRow("promo_c1_l")
                        ZSvalue = productRow("ancd_rebates")
                        Priorita = "Standard"
                    End If
                    
                Case Else
                    promoValue = ""
                    promoName = ""
                    zsName = ""
                    AFCvalue = ""
                    kompValue = ""
                    c1Value = ""
                    ZSvalue = ""
                    Priorita = ""
            End Select
        End If
        
    Else
        promoValue = ""
        promoName = ""
        zsName = ""
        AFCvalue = ""
        kompValue = ""
        c1Value = ""
        ZSvalue = ""
        Priorita = ""
        FCname = ""
    End If
        
    GetPromoPriceData = Array(promoValue, promoName, zsName, AFCvalue, kompValue, c1Value, ZSvalue, Priorita, FCname)
End Function

Public Sub FillSelectedProductsToTextList(TargetWorkbook As Workbook, selectedProducts As Variant, familyValue As String, selectedPrice As String, PromoObj As Object, heroProduct As String, promoID As String, vyberValue As String, pcsPlanText As String, isPlan As Boolean, fcType As String, countryCode As String, commentText As String)
    
    Call UnlockText(TargetWorkbook)
    Call RemoveFilterIfApplied(TargetWorkbook)

    ' P�id� pouze vybran� produkty
    Dim productsCol As Collection
    Set productsCol = GetProductsCollection()

    If productsCol Is Nothing Or productsCol.Count = 0 Then
        Call ProductsArray(TargetWorkbook)
        Set productsCol = GetProductsCollection()
    End If

    If productsCol.Count = 0 Then
        MsgBox "��dn� data k na�ten�!"
        Exit Sub
    End If

    Dim textList As Worksheet
    Set textList = TargetWorkbook.Sheets("Text")
    Dim firstEmptyRow As Long
    firstEmptyRow = textList.Cells(textList.rows.Count, GetColumnSafe(textList, "tProduct")).End(xlUp).row + 1

    If firstEmptyRow <= 2 Then firstEmptyRow = 3

    ' Rozd�lit text podle ��rky
    Dim PcsArray() As String
    PcsArray = Split(pcsPlanText, ",")
    Dim j As Long
    j = 0

    ' Projde vybran� produkty
    Dim selectedProduct As Variant
    Dim rowData As Object

    For Each selectedProduct In selectedProducts
        For Each rowData In productsCol
        
            ' Podle countryCode rozhodnout form�t
            Dim productName As String
            If UCase(Trim(countryCode)) = "SVK" Then
                productName = rowData("material_name")
            Else
                productName = rowData("material_name") & " " & rowData("volume_l")
            End If

            If productName = CStr(selectedProduct) Then
                With textList
                    WriteToColumnSafe textList, firstEmptyRow, "tProduct", productName
                    WriteToColumnSafe textList, firstEmptyRow, "tCustomerID", rowData("CustomerID")
                    WriteToColumnSafe textList, firstEmptyRow, "tEAN", "'" & rowData("ean")
                    WriteToColumnSafe textList, firstEmptyRow, "tPackageSize", rowData("volume_l")
                    WriteToColumnSafe textList, firstEmptyRow, "tStockID", rowData("sap_id")
                    WriteToColumnSafe textList, firstEmptyRow, "tBrand", rowData("Brand")
                    WriteToColumnSafe textList, firstEmptyRow, "tCustomer", GetCustomer(TargetWorkbook)
                    WriteToColumnSafe textList, firstEmptyRow, "tFC", rowData("ncd_invoice")
                    WriteToColumnSafe textList, firstEmptyRow, "tNCD", rowData("ncd_inc_vat")
                    WriteToColumnSafe textList, firstEmptyRow, "tFamily", rowData("Family")
                    WriteToColumnSafe textList, firstEmptyRow, "tCategory", rowData("category")
                    WriteToColumnSafe textList, firstEmptyRow, "tVyber", vyberValue


                    WriteToColumnSafe textList, firstEmptyRow, "tAkceOd", PromoObj.startAkce
                    WriteToColumnSafe textList, firstEmptyRow, "tAkceDo", PromoObj.endAkce
                    WriteToColumnSafe textList, firstEmptyRow, "tNakupOd", PromoObj.StartPurchase
                    WriteToColumnSafe textList, firstEmptyRow, "tNakupDo", PromoObj.EndPurchase
                    WriteToColumnSafe textList, firstEmptyRow, "tSortFrom", PromoObj.sortFrom
                    WriteToColumnSafe textList, firstEmptyRow, "tSortTo", PromoObj.sortTo
                    WriteToColumnSafe textList, firstEmptyRow, "tTypAkce", PromoObj.typAkce
                    WriteToColumnSafe textList, firstEmptyRow, "tPromo", PromoObj.promoTyp
                    WriteToColumnSafe textList, firstEmptyRow, "tWeeks", PromoObj.WeekRange
                    WriteToColumnSafe textList, firstEmptyRow, "tWeeksT", PromoObj.weekRangeT
                    WriteToColumnSafe textList, firstEmptyRow, "tCom", PromoObj.promoTyp & " " & PromoObj.WeekRange
                    WriteToColumnSafe textList, firstEmptyRow, "tPromoID", promoID

                    ' Adding price data from the GetPromoPriceData function
                    Dim result As Variant
                    result = GetPromoPriceData(familyValue, selectedPrice, rowData, fcType)

                    WriteToColumnSafe textList, firstEmptyRow, "tPromoPrice", result(0)  ' promoValue
                    WriteToColumnSafe textList, firstEmptyRow, "tPriceType", result(1)   ' promoName
                    WriteToColumnSafe textList, firstEmptyRow, "tZS", result(6)          ' zsValue
                    WriteToColumnSafe textList, firstEmptyRow, "tAFC", result(3)         ' AFCvalue
                    WriteToColumnSafe textList, firstEmptyRow, "tKomp", result(4)        ' kompValue
                    WriteToColumnSafe textList, firstEmptyRow, "tC1l", result(5)         ' c1Value
                    WriteToColumnSafe textList, firstEmptyRow, "tPriorita", result(7)    ' Priorita
                    WriteToColumnSafe textList, firstEmptyRow, "tZStype", result(2)      ' zsName
                    WriteToColumnSafe textList, firstEmptyRow, "tFCname", result(8)      ' FCname
                    
                   ' Hero
                    If productName = heroProduct Then
                        WriteToColumnSafe textList, firstEmptyRow, "tHero", "A"
                    Else
                        WriteToColumnSafe textList, firstEmptyRow, "tHero", "N"
                    End If

                    ' Pozn�mka - writes "Pl�n" if CB_Plan is checked
                    If isPlan Then
                        WriteToColumnSafe textList, firstEmptyRow, "tPotvrzeno", "Pl�n"
                    Else
                        WriteToColumnSafe textList, firstEmptyRow, "tPotvrzeno", ""
                    End If
                    
                    ' Z�pis koment��e z TB_Comment
                    WriteToColumnSafe textList, firstEmptyRow, "tPozn", commentText

                    WriteToColumnSafe textList, firstEmptyRow, "tDiff", "=(" & Range("tRealPromoPrice").Cells(firstEmptyRow, 1).Address(False, False) & "/" & Range("tPromoPrice").Cells(firstEmptyRow, 1).Address(False, False) & ")-1"
                    WriteToColumnSafe textList, firstEmptyRow, "tC1Total", "=" & Range("tC1l").Cells(firstEmptyRow, 1).Address(False, False) & "*" & Range("tOffTakes").Cells(firstEmptyRow, 1).Address(False, False)
                    WriteToColumnSafe textList, firstEmptyRow, "tOffTakeTotal", "=" & Range("tRealPromoPrice").Cells(firstEmptyRow, 1).Address(False, False) & "*" & Range("tOffTakes").Cells(firstEmptyRow, 1).Address(False, False)
                    WriteToColumnSafe textList, firstEmptyRow, "tVol", "=" & Range("tPcsPlan").Cells(firstEmptyRow, 1).Address(False, False) & "*" & Range("tPackageSize").Cells(firstEmptyRow, 1).Address(False, False)

                    If j <= UBound(PcsArray) Then
                        WriteToColumnSafe textList, firstEmptyRow, "tOffTakes", Trim(PcsArray(j))
                    Else
                        WriteToColumnSafe textList, firstEmptyRow, "tOffTakes", ""
                    End If
                    j = j + 1

                End With
                firstEmptyRow = firstEmptyRow + 1
                Exit For
            End If
        Next rowData
    Next selectedProduct

End Sub

Public Function GetPLastRow() As Long
    GetPLastRow = pLastRow
End Function

Public Function GetProductValue(rowIndex As Long, columnName As String) As Variant
    ' Z�sk� produkt podle indexu v kolekci
    Dim productsCol As Collection
    Set productsCol = GetProductsCollection()
    
    ' Kontrola, zda index je platn�
    If rowIndex < 1 Or rowIndex > productsCol.Count Then
        GetProductValue = ""
        Exit Function
    End If
    
    ' Z�sk�n� ��dku produktu
    Dim productRow As Object
    Set productRow = productsCol(rowIndex)
    
    ' Vr�cen� hodnoty podle n�zvu sloupce
    Select Case LCase(columnName)
        Case "family"
            If productRow.Exists("Family") Then
                GetProductValue = productRow("Family")
            Else
                GetProductValue = ""
            End If
        Case "product", "material_name"
            If productRow.Exists("material_name") Then
                GetProductValue = productRow("material_name")
            Else
                GetProductValue = ""
            End If
        Case "packagesize", "volume_l"
            If productRow.Exists("volume_l") Then
                GetProductValue = productRow("volume_l")
            Else
                GetProductValue = ""
            End If
        Case Else
            ' Pokus� se naj�t sloupec podle p�esn�ho n�zvu
            If productRow.Exists(columnName) Then
                GetProductValue = productRow(columnName)
            Else
                GetProductValue = ""
            End If
    End Select
End Function

Public Sub ShowUserForm2(TargetWorkbook As Workbook, selectedAddress As String)
    On Error GoTo ErrorHandler
    
    Debug.Print "=== ShowUserForm2 START ==="
    Debug.Print "TargetWorkbook: " & TargetWorkbook.Name
    Debug.Print "selectedAddress: " & selectedAddress
    
    ' Naj�t aktivn� list v TargetWorkbook
    Dim ws As Worksheet
    Set ws = TargetWorkbook.ActiveSheet
    
    ' P�ev�st adresu zp�t na Range
    Dim SelectedRange As Range
    Set SelectedRange = ws.Range(selectedAddress)
    
    Debug.Print "SelectedRange: " & SelectedRange.Address
    
    ' Vytvo�it novou instanci UserForm2
    Dim uf As UserForm2
    Set uf = New UserForm2
    
    ' P�edat odkazy
    Set uf.TargetWorkbook = TargetWorkbook
    Set uf.SelectedRange = SelectedRange
    
    ' Zobrazit UserForm
    uf.Show
    
    Debug.Print "=== ShowUserForm2 END ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "CHYBA: " & Err.Description
    MsgBox "Chyba v ShowUserForm2: " & Err.Description, vbCritical
End Sub

Public Sub PridejVybraneHeroProdukty(UserForm As Object, selectedPrice As String, PromoObj As Object, heroProduct As String, promoID As String, vyberValue As String, pcsPlanText As String, isPlan As Boolean, TargetWorkbook As Workbook, SelectedRange As Range, fcType As String, countryCode As String, commentText As String)
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== PridejVybraneHeroProdukty START ==="
    
    ' Na��st produkty
    Call ProductsArray(TargetWorkbook)
   
    ' Z�skat vybran� produkty z UserFormu
    Dim selectedProducts As Variant
    selectedProducts = GetSelectedItems(UserForm.LB_Product)
    
    ' Debug
    Debug.Print "TypeName: " & TypeName(selectedProducts)
    If IsArray(selectedProducts) Then
        Debug.Print "Po�et produkt�: " & (UBound(selectedProducts) - LBound(selectedProducts) + 1)
        Dim i As Long
        For i = LBound(selectedProducts) To UBound(selectedProducts)
            Debug.Print "Produkt " & i & ": " & selectedProducts(i)
        Next i
    End If
    
    ' Family value z SelectedRange
    Dim familyValue As String
    familyValue = SelectedRange.Worksheet.Range("C" & SelectedRange.row).value
    Debug.Print "Family: " & familyValue
    Debug.Print "FC Type: " & fcType
    
    ' Na�te countryCode z Settings
    If Trim(countryCode) = "" Then
        countryCode = "CZK"  ' Default
    End If
    
    Debug.Print "Country Code: " & countryCode
    Debug.Print "=== DEBUG P�ED VOL�N�M ==="
    
    Call FillSelectedProductsToTextList(TargetWorkbook, selectedProducts, familyValue, selectedPrice, PromoObj, heroProduct, promoID, vyberValue, pcsPlanText, isPlan, fcType, countryCode, commentText)
    Debug.Print "=== PridejVybraneHeroProdukty END ==="
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "CHYBA v PridejVybraneHeroProdukty: " & Err.Description
End Sub

Public Sub GetWeekIntervalsFromSelection(SelectedRange As Range, ByRef weekInterval As String, ByRef weekIntervalT As String, TargetWorkbook As Workbook)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = SelectedRange.Worksheet

    Dim weekRowNumber As Long, weekRowNumberT As Long
    weekRowNumber = FindWeekRow(ws)

    If weekRowNumber = 0 Then
        Err.Raise vbObjectError + 2, "GetWeekIntervalsFromSelection", "��dek s koment��em 'WeekRow' nebyl nalezen!"
    End If

    weekRowNumberT = weekRowNumber - 1

    Dim m As Long, o As Long, cc As Long
    m = SelectedRange.Cells(1).Column

    With SelectedRange
        cc = .Columns.Count
        o = .Cells(.Count).Column
    End With

    If cc = 1 Then
        weekInterval = ws.Cells(weekRowNumber, m).value
        weekIntervalT = ws.Cells(weekRowNumberT, m).value
    Else
        weekInterval = ws.Cells(weekRowNumber, m).value & "-" & ws.Cells(weekRowNumber, o).value
        weekIntervalT = ws.Cells(weekRowNumberT, m).value & "-" & ws.Cells(weekRowNumberT, o).value
    End If

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "GetWeekIntervalsFromSelection", Err.Description
End Sub

