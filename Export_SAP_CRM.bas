Attribute VB_Name = "Export_SAP_CRM"

Public Sub SAP_CRM(TargetWorkbook As Workbook, SelectedRange As Range)
    
    Call UnlockText(TargetWorkbook)
    Call UnlockSAP(TargetWorkbook)
    
    ' PØIDÁNO: Naèíst countryCode
    Dim countryCode As String
    On Error Resume Next
    countryCode = Trim(TargetWorkbook.Sheets("Settings").Range("B10").value)
    On Error GoTo 0
    
    If countryCode = "" Then countryCode = "CZK"
    
    Debug.Print "Country Code: " & countryCode
    
    ' Naèíst data do kolekce - vždy naèíst znovu
    Call ProductsArray(TargetWorkbook)
    
    ' Získat kolekci
    Dim productsCol As Collection
    On Error Resume Next
    Set productsCol = GetProductsCollection()
    On Error GoTo 0
    
    ' Kontrola, zda se kolekce naèetla
    If productsCol Is Nothing Then
        MsgBox "Chyba: ProductsCollection není inicializována!" & vbCrLf & _
               "Ujistìte se, že list PriceList obsahuje data.", vbCritical
        Call LockText(TargetWorkbook)
        Call LockSAP(TargetWorkbook)
        Exit Sub
    End If
    
    If productsCol.Count = 0 Then
        MsgBox "Žádná data k naètení z PriceList!", vbExclamation
        Call LockText(TargetWorkbook)
        Call LockSAP(TargetWorkbook)
        Exit Sub
    End If
    
    Debug.Print "Products loaded: " & productsCol.Count
    
    ' Nastavení listù
    Dim sapList As Worksheet
    Dim CrmList As Worksheet
    Dim textList As Worksheet
    Dim settingsSheet As Worksheet
    
    Set sapList = TargetWorkbook.Sheets("SAP")
    Set CrmList = TargetWorkbook.Sheets("CRM")
    Set textList = TargetWorkbook.Sheets("Text")
    Set settingsSheet = TargetWorkbook.Sheets("Settings")
    
    Dim customerHierarchy As String
    customerHierarchy = settingsSheet.Range("B6").value
    
    ' Naètení dat z Text listu do pole
    Dim tFirstRow As Long
    Dim tRowCount As Long
    tFirstRow = SelectedRange(1).row
    tRowCount = SelectedRange.rows.Count
    
    ' Urèení sloupcù
    Dim colTypAkce As Long, colPriorita As Long, colStockID As Long
    Dim colNakupOd As Long, colNakupDo As Long, colAkceOd As Long, colAkceDo As Long
    Dim colProduct As Long, colEAN As Long, colAFC As Long, colPromoPrice As Long
    Dim colFamily As Long, colPromoID As Long
    
    colTypAkce = textList.Range("tTypAkce").Column
    colPriorita = textList.Range("tPriorita").Column
    colStockID = textList.Range("tStockID").Column
    colNakupOd = textList.Range("tNakupOd").Column
    colNakupDo = textList.Range("tNakupDo").Column
    colAkceOd = textList.Range("tAkceOd").Column
    colAkceDo = textList.Range("tAkceDo").Column
    colProduct = textList.Range("tProduct").Column
    colEAN = textList.Range("tEAN").Column
    colAFC = textList.Range("tAFC").Column
    colPromoPrice = textList.Range("tPromoPrice").Column
    colFamily = textList.Range("tFamily").Column
    colPromoID = textList.Range("tPromoID").Column
    
    ' Naèíst všechna data z výbìru najednou do pole
    Dim textData As Variant
    ReDim textData(1 To tRowCount, 1 To 13)
    
    Dim i As Long
    For i = 1 To tRowCount
        textData(i, 1) = textList.Cells(tFirstRow + i - 1, colTypAkce).value
        textData(i, 2) = textList.Cells(tFirstRow + i - 1, colPriorita).value
        textData(i, 3) = textList.Cells(tFirstRow + i - 1, colStockID).value
        textData(i, 4) = textList.Cells(tFirstRow + i - 1, colNakupOd).value
        textData(i, 5) = textList.Cells(tFirstRow + i - 1, colNakupDo).value
        textData(i, 6) = textList.Cells(tFirstRow + i - 1, colAkceOd).value
        textData(i, 7) = textList.Cells(tFirstRow + i - 1, colAkceDo).value
        textData(i, 8) = textList.Cells(tFirstRow + i - 1, colProduct).value
        textData(i, 9) = textList.Cells(tFirstRow + i - 1, colEAN).value
        textData(i, 10) = textList.Cells(tFirstRow + i - 1, colAFC).value
        textData(i, 11) = textList.Cells(tFirstRow + i - 1, colPromoPrice).value
        textData(i, 12) = textList.Cells(tFirstRow + i - 1, colFamily).value
        textData(i, 13) = textList.Cells(tFirstRow + i - 1, colPromoID).value
    Next i
    
    ' Vymazání dat z listu SAP (zachovat hlavièky na øádcích 1-3)
    Dim lastSapRow As Long
    lastSapRow = sapList.Cells(sapList.rows.Count, 1).End(xlUp).row
    
    If lastSapRow >= 4 Then
        sapList.rows("4:" & lastSapRow).Delete
    End If
    
    ' Pøíprava dat pro SAP (nový formát)
    Dim sapData() As Variant
    ReDim sapData(1 To tRowCount * productsCol.Count, 1 To 14)
    
    Dim sapRowIndex As Long
    sapRowIndex = 1
    
    ' Counter pro incrementální hodnoty
    Dim rowCounter As Long
    rowCounter = 1
    
    ' Pøíprava dat pro CRM
    Dim crmData() As Variant
    ReDim crmData(1 To tRowCount * productsCol.Count, 1 To 11)
    
    Dim crmRowIndex As Long
    crmRowIndex = 1
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Projít vybrané øádky a produkty
    Dim j As Long
    Dim productRow As Object
    Dim productKey As String
    Dim textKey As String
    Dim VypocetH As Double
    
    For j = 1 To tRowCount
        textKey = textData(j, 12) & textData(j, 8) ' Family & Product
        
        ' Najít odpovídající produkt v kolekci
        For Each productRow In productsCol
            If productRow.Exists("Family") And productRow.Exists("material_name") And productRow.Exists("volume_l") Then
                
                ' ZMÌNA: productKey podle countryCode
                If UCase(Trim(countryCode)) = "SVK" Then
                    productKey = productRow("Family") & productRow("material_name")
                Else
                    productKey = productRow("Family") & productRow("material_name") & " " & productRow("volume_l")
                End If
                
                If textKey = productKey Then
                    ' Pøidat do SAP pole (NOVÝ FORMÁT)
                    sapData(sapRowIndex, 1) = "ZP01"                                    ' A - ConditionType
                    sapData(sapRowIndex, 2) = 922                                       ' C - ConditionTable
                    sapData(sapRowIndex, 3) = "CZ10"                                    ' E - SalesOrganization
                    sapData(sapRowIndex, 4) = 10                                        ' G - DistributionChannel
                    sapData(sapRowIndex, 5) = textData(j, 3)                            ' K - Material (StockID)
                    sapData(sapRowIndex, 6) = customerHierarchy                         ' X - CustomerHierarchy
                    sapData(sapRowIndex, 7) = Format(textData(j, 4), "YYYYMMDD")        ' AE - ValidityStartDate
                    sapData(sapRowIndex, 8) = Format(textData(j, 5), "YYYYMMDD")        ' AF - ValidityEndDate
                    
                    ' Výpoèet hodnoty (BEZ Q1)
                    If productRow.Exists("base_price") And productRow.Exists("special_discount") Then
                        VypocetH = (-1 * textData(j, 10) / productRow("base_price") + 1 - productRow("special_discount") / 100) * 100
                        ' Nastavení hodnoty s desetinnou teèkou a záporným znaménkem
                        sapData(sapRowIndex, 9) = "'" & Replace(CStr(Round(VypocetH, 3) * (-1)), ",", ".") ' AG - ConditionRateValue
                    Else
                        sapData(sapRowIndex, 9) = "'0.000"
                    End If
                    
                    sapData(sapRowIndex, 10) = "%"                                      ' AH - ConditionRateValueUnit
                    sapData(sapRowIndex, 11) = "$$" & Format(rowCounter, "00000000")    ' AA - ConditionRecord
                    sapData(sapRowIndex, 12) = "'01"                                    ' AB - ConditionSequentialNumber
                    sapData(sapRowIndex, 13) = textData(j, 8)                           ' BA - Product název
                    sapData(sapRowIndex, 14) = textData(j, 10)                          ' BB - AFC hodnota
                    
                    sapRowIndex = sapRowIndex + 1
                    rowCounter = rowCounter + 1
                    
                    ' Pøidat do CRM pole
                    crmData(crmRowIndex, 1) = textData(j, 13)     ' cIDakce
                    crmData(crmRowIndex, 2) = textData(j, 8)      ' cNazevProduktu
                    crmData(crmRowIndex, 3) = "'" & textData(j, 9) ' cEAN
                    crmData(crmRowIndex, 4) = "Planned"           ' cStatus
                    crmData(crmRowIndex, 5) = "Tesco"             ' cZakaznik
                    crmData(crmRowIndex, 6) = customerHierarchy   ' cZakaznikSAP
                    crmData(crmRowIndex, 7) = textData(j, 6)      ' cAkceOd
                    crmData(crmRowIndex, 8) = textData(j, 7)      ' cAkceDo
                    crmData(crmRowIndex, 9) = textData(j, 2)      ' cPriorita
                    crmData(crmRowIndex, 10) = textData(j, 1)     ' cTypAkce
                    crmData(crmRowIndex, 11) = textData(j, 11)    ' cPromoCena
                    
                    crmRowIndex = crmRowIndex + 1
                    
                    Exit For ' Produkt nalezen
                End If
            End If
        Next productRow
    Next j
    
    ' Bulk zápis do SAP (NOVÝ FORMÁT) - zaèíná od øádku 4
    If sapRowIndex > 1 Then
        Dim sapStartRow As Long
        sapStartRow = 4 ' Zaèínáme od 4. øádku (øádky 1-3 jsou hlavièky)
        
        For i = 1 To sapRowIndex - 1
            sapList.Cells(sapStartRow + i - 1, 1).value = sapData(i, 1)      ' A - ConditionType
            sapList.Cells(sapStartRow + i - 1, 3).value = sapData(i, 2)      ' C - ConditionTable
            sapList.Cells(sapStartRow + i - 1, 5).value = sapData(i, 3)      ' E - SalesOrganization
            sapList.Cells(sapStartRow + i - 1, 7).value = sapData(i, 4)      ' G - DistributionChannel
            sapList.Cells(sapStartRow + i - 1, 11).value = sapData(i, 5)     ' K - Material
            sapList.Cells(sapStartRow + i - 1, 24).value = sapData(i, 6)     ' X - CustomerHierarchy
            sapList.Cells(sapStartRow + i - 1, 27).value = sapData(i, 11)    ' AA - ConditionRecord
            sapList.Cells(sapStartRow + i - 1, 28).value = sapData(i, 12)    ' AB - ConditionSequentialNumber
            sapList.Cells(sapStartRow + i - 1, 31).value = sapData(i, 7)     ' AE - ValidityStartDate
            sapList.Cells(sapStartRow + i - 1, 32).value = sapData(i, 8)     ' AF - ValidityEndDate
            sapList.Cells(sapStartRow + i - 1, 33).value = sapData(i, 9)     ' AG - ConditionRateValue
            sapList.Cells(sapStartRow + i - 1, 34).value = sapData(i, 10)    ' AH - ConditionRateValueUnit
            sapList.Cells(sapStartRow + i - 1, 53).value = sapData(i, 13)    ' BA - Product název
            sapList.Cells(sapStartRow + i - 1, 54).value = sapData(i, 14)    ' BB - AFC
        Next i
    End If
    
    ' Zápis CSV na list Text
    Dim colCSV As Long
    colCSV = textList.Range("tCSV").Column
    For i = 0 To tRowCount - 1
        textList.Cells(tFirstRow + i, colCSV).value = "ANO"
    Next i
    
    ' Bulk zápis do CRM
    If crmRowIndex > 1 Then
        Dim crmStartRow As Long
        crmStartRow = CrmList.Cells(CrmList.rows.Count, CrmList.Range("cIDakce").Column).End(xlUp).row + 1
        
        Dim colCIDakce As Long, colCNazev As Long, colCEAN As Long, colCStatus As Long
        Dim colCZakaznik As Long, colCZakaznikSAP As Long, colCAkceOd As Long
        Dim colCAkceDo As Long, colCPriorita As Long, colCTypAkce As Long, colCPromoCena As Long
        
        colCIDakce = CrmList.Range("cIDakce").Column
        colCNazev = CrmList.Range("cNazevProduktu").Column
        colCEAN = CrmList.Range("cEAN").Column
        colCStatus = CrmList.Range("cStatus").Column
        colCZakaznik = CrmList.Range("cZakaznik").Column
        colCZakaznikSAP = CrmList.Range("cZakaznikSAP").Column
        colCAkceOd = CrmList.Range("cAkceOd").Column
        colCAkceDo = CrmList.Range("cAkceDo").Column
        colCPriorita = CrmList.Range("cPriorita").Column
        colCTypAkce = CrmList.Range("cTypAkce").Column
        colCPromoCena = CrmList.Range("cPromoCena").Column
        
        For i = 1 To crmRowIndex - 1
            CrmList.Cells(crmStartRow + i - 1, colCIDakce).value = crmData(i, 1)
            CrmList.Cells(crmStartRow + i - 1, colCNazev).value = crmData(i, 2)
            CrmList.Cells(crmStartRow + i - 1, colCEAN).value = crmData(i, 3)
            CrmList.Cells(crmStartRow + i - 1, colCStatus).value = crmData(i, 4)
            CrmList.Cells(crmStartRow + i - 1, colCZakaznik).value = crmData(i, 5)
            CrmList.Cells(crmStartRow + i - 1, colCZakaznikSAP).value = crmData(i, 6)
            CrmList.Cells(crmStartRow + i - 1, colCAkceOd).value = crmData(i, 7)
            CrmList.Cells(crmStartRow + i - 1, colCAkceDo).value = crmData(i, 8)
            CrmList.Cells(crmStartRow + i - 1, colCPriorita).value = crmData(i, 9)
            CrmList.Cells(crmStartRow + i - 1, colCTypAkce).value = crmData(i, 10)
            CrmList.Cells(crmStartRow + i - 1, colCPromoCena).value = crmData(i, 11)
        Next i
    End If
    
    ' Mazání odbìhlých promocí z CRM
    Dim cLastRow As Long
    colCAkceDo = CrmList.Range("cAkceDo").Column
    cLastRow = CrmList.Cells(CrmList.rows.Count, colCAkceDo).End(xlUp).row
    
    If cLastRow > 1 Then
        For i = cLastRow To 2 Step -1
            If CrmList.Cells(i, colCAkceDo).value < Date Then
                CrmList.rows(i).Delete
            End If
        Next i
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Call LockText(TargetWorkbook)
    Call LockSAP(TargetWorkbook)
    
    Debug.Print "CSV: Zpracováno " & (sapRowIndex - 1) & " SAP øádkù a " & (crmRowIndex - 1) & " CRM øádkù"
    
End Sub

Public Sub ExportData(TargetWorkbook As Workbook)
        
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandler
    
    ' Nastavení listù
    Dim sapList As Worksheet
    Dim settingsSheet As Worksheet
    Set sapList = TargetWorkbook.Sheets("SAP")
    Set settingsSheet = TargetWorkbook.Sheets("Settings")
    
    ' Naètení cest ze Settings
    Dim path1 As String
    Dim path2 As String
    Dim sharePointPath As String
    
    path1 = settingsSheet.Range("B7").value
    path2 = settingsSheet.Range("B8").value
    sharePointPath = "https://stockgroup.sharepoint.com/sites/power-apps-data/promotool-automation-uat/Shared%20Documents/landing_v2/"
    
    ' Vytvoøení názvu souboru
    Dim fileName As String
    fileName = sapList.Cells(4, 24).value & "_"
    
    Dim TimeStamp As String
    TimeStamp = Format(Now, "yyyymmdd-hhnnss")
    
    ' Vytvoøení nového sešitu
    Dim Wbk As Workbook
    Set Wbk = Workbooks.Add
    
    ' Nastavení textového formátu
    Wbk.Sheets(1).Cells.NumberFormat = "@"
    
    ' Kopírování dat z listu SAP
    sapList.UsedRange.Copy
    Wbk.Sheets(1).Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' Pøevod èíselných datumù na textový formát
    Dim lastRow As Long
    Dim i As Long
    lastRow = Wbk.Sheets(1).UsedRange.rows.Count
    
    ' Pøevod dat ve sloupcích F a G na èitelný textový formát
    For i = 2 To lastRow
        ' Sloupec F (PLATOD)
        If IsNumeric(Wbk.Sheets(1).Cells(i, 6).value) And Wbk.Sheets(1).Cells(i, 6).value > 0 Then
            Wbk.Sheets(1).Cells(i, 6).value = Format(CDate(Wbk.Sheets(1).Cells(i, 6).value), "dd.mm.yyyy")
        End If
        ' Sloupec G (PLATDO)
        If IsNumeric(Wbk.Sheets(1).Cells(i, 7).value) And Wbk.Sheets(1).Cells(i, 7).value > 0 Then
            Wbk.Sheets(1).Cells(i, 7).value = Format(CDate(Wbk.Sheets(1).Cells(i, 7).value), "dd.mm.yyyy")
        End If
    Next i
    
    ' Odstranìní všech tlaèítek z nového listu
    Dim shp As Shape
    Dim found As Boolean
    
    Do
        found = False
        For Each shp In Wbk.Sheets(1).Shapes
            If Left(shp.Name, 6) = "Button" Then
                shp.Delete
                found = True
                Exit For
            End If
        Next shp
    Loop While found
    
    ' Urèení lokální cesty
    Dim targetFolderPath As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Zkontrolujte cesty ze Settings
    If path1 <> "" And fso.FolderExists(path1) Then
        targetFolderPath = path1
        If Right(targetFolderPath, 1) <> "\" Then targetFolderPath = targetFolderPath & "\"
    ElseIf path2 <> "" And fso.FolderExists(path2) Then
        targetFolderPath = path2
        If Right(targetFolderPath, 1) <> "\" Then targetFolderPath = targetFolderPath & "\"
    Else
        MsgBox "Žádná z cest v Settings (B7, B8) neexistuje!", vbCritical
        GoTo CleanUp
    End If
    
    ' Cesty pro uložení
    Dim wbkSharePointPath As String
    Dim wbkLocalPath As String
    
    wbkSharePointPath = sharePointPath & fileName & TimeStamp & ".xlsx"
    wbkLocalPath = targetFolderPath & fileName & TimeStamp & ".xlsx"
       
    On Error Resume Next
    Wbk.Sheets(1).Protect Password:=GetPassword(), DrawingObjects:=True, Contents:=True, Scenarios:=True
    On Error GoTo ErrorHandler
    
    ' Uložení do SharePointu
    On Error Resume Next
    Wbk.SaveAs _
        fileName:=wbkSharePointPath, _
        FileFormat:=xlOpenXMLWorkbook, _
        CreateBackup:=False
    
    Dim sharePointSaved As Boolean
    sharePointSaved = (Err.Number = 0)
    
    If Not sharePointSaved Then
        Debug.Print "SharePoint uložení selhalo: " & Err.Description
    End If
    On Error GoTo ErrorHandler
    
    ' Uložení do lokální složky
    Wbk.SaveAs _
        fileName:=wbkLocalPath, _
        FileFormat:=xlOpenXMLWorkbook, _
        CreateBackup:=False
    
    Debug.Print "Export dokonèen: " & wbkLocalPath
    
    If sharePointSaved Then
        MsgBox "Export dokonèen!" & vbCrLf & _
               "Lokální: " & wbkLocalPath & vbCrLf & _
               "SharePoint: " & wbkSharePointPath, vbInformation
    Else
        MsgBox "Export dokonèen lokálnì!" & vbCrLf & _
               "Lokální: " & wbkLocalPath & vbCrLf & vbCrLf & _
               "SharePoint uložení selhalo.", vbExclamation
    End If
    
CleanUp:
    On Error Resume Next
    If Not Wbk Is Nothing Then
        Wbk.Saved = True
        Wbk.Close
    End If
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    Debug.Print "Chyba " & Err.Number & ": " & Err.Description
    MsgBox "Chyba pøi exportu: " & Err.Description, vbCritical
    Resume CleanUp
    
End Sub

