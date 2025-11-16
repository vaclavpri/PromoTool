VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Zadej promoci"
   ClientHeight    =   9864.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8748.001
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SelectedRange As Range
Private Promo As Object
Private m_TargetWorkbook As Workbook

Public Property Set TargetWorkbook(wb As Workbook)
    Set m_TargetWorkbook = wb
    Call LoadFCTypesToListBox
End Property

Public Property Get TargetWorkbook() As Workbook
    Set TargetWorkbook = m_TargetWorkbook
End Property

Private Sub CommandButton1_Click()
    Dim n As Integer
    If LB_Promoce.ListIndex = -1 Or LB_Price.ListIndex = -1 Then
        MsgBox "Nejsou vybrány všechny povinné údaje."
    Else
        For n = 0 To LB_Product.ListCount - 1
            If LB_Product.Selected(n) = True Then
                LB_Hero.AddItem LB_Product.List(n)
            End If
        Next n
    End If
End Sub

Private Sub CommandButton2_Click()
    If LB_Hero.ListIndex = -1 Then
        MsgBox "Vyber Hero produkt."
    Else
        InsertPromo  ' Nyní tato procedura existuje níže
    End If
End Sub

Private Sub InsertPromo()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== InsertPromo START ==="
    
    ' Zavolat PromoSet
    Call PromoSet
    
    ' Ovìøení, že Promo je vytvoøený
    If Promo Is Nothing Then
        MsgBox "Chyba: Promo nebyl vytvoøen!"
        Exit Sub
    End If
    
    Debug.Print "Promo typ: " & Promo.promoTyp
    
    ' Generovat PromoID
    Dim promoID As String
    promoID = GenerateID(TargetWorkbook)
    Debug.Print "Vygenerované PromoID: " & promoID
    
    ' Reference na Text list
    Dim textList As Worksheet
    Set textList = TargetWorkbook.Sheets("Text")
    
    Dim firstEmptyRow As Long
    firstEmptyRow = textList.Cells(textList.rows.Count, textList.Range("tProduct").Column).End(xlUp).row + 1
    Debug.Print "firstEmptyRow: " & firstEmptyRow
    
    If firstEmptyRow <= 2 Then
        firstEmptyRow = 3
    End If
    
    Dim Fami As Range
    Set Fami = SelectedRange.Worksheet.Range("C" & SelectedRange.row)
    Debug.Print "Family: " & Fami.value
    
    Dim selectedPrice As String
    selectedPrice = LB_Price.List(LB_Price.ListIndex)
    Debug.Print "Selected Price: " & selectedPrice
    
    ' Objemy
    Dim pcsPlanText As String
    pcsPlanText = TB_PcsPlan.Text
    Debug.Print "PCS Plan Text: " & pcsPlanText
    
    ' Výbìr
    Dim vyberValue As String
    vyberValue = GetVyberValue(Me.LB_Product)
    Debug.Print "Vyber Value: " & vyberValue
    
    ' Hero
    Dim heroProduct As String
    heroProduct = GetHeroItem(Me.LB_Hero)
    heroProduct = Me.LB_Product.List(Me.LB_Product.ListIndex)
    Debug.Print "Hero Product: " & heroProduct
    
    ' Plán
    Dim isPlan As Boolean
    isPlan = CB_Plan.value
    Debug.Print "Is Plan: " & isPlan
    
    ' Naèíst countryCode z Settings
    Dim countryCode As String
    countryCode = GetCountryCode()
    Debug.Print "Country Code: " & countryCode
       
    Dim commentText As String
    commentText = Trim(Me.TB_Comment.value)
    Debug.Print "Comment: " & commentText
       
    Debug.Print "=== PØED VOLÁNÍM PridejVybraneHeroProdukty ==="
    
    ' Zápis do listu Text
    Call PridejVybraneHeroProdukty(Me, selectedPrice, Promo, heroProduct, promoID, vyberValue, pcsPlanText, isPlan, TargetWorkbook, SelectedRange, Me.LB_FC.value, countryCode, commentText)
    
    Debug.Print "=== PO VOLÁNÍ PridejVybraneHeroProdukty ==="
    
    ' PØESUNUTO: Seøazení PØED formátováním
    Debug.Print "=== PØED Seøazením ==="
    Call ApplyFilterToRow2(TargetWorkbook)
    Call SortIt(TargetWorkbook)
    Debug.Print "=== PO Seøazení ==="
    
    ' Barvení øádkù
    Call rColor(TargetWorkbook)
    Debug.Print "=== PO rColor ==="
    
        Debug.Print "=== KONTROLA KOMENTÁØÙ PØED FormatPromoCells ==="
    Dim checkCell As Range
    For Each checkCell In SelectedRange.Cells
        If Not checkCell.comment Is Nothing Then
            Debug.Print "  Buòka " & checkCell.Address & " má komentáø: " & Left(checkCell.comment.Text, 8)
        Else
            Debug.Print "  Buòka " & checkCell.Address & " NEMÁ komentáø"
        End If
    Next checkCell
    Debug.Print "=== KONEC KONTROLY ==="
    
    ' Formátování kostièek
    Dim usePlanColor As Boolean
    usePlanColor = CB_Plan.value
    Call FormatPromoCells(TargetWorkbook, SelectedRange, Promo, promoID, usePlanColor)
        
    Debug.Print "=== PO FormatPromoCells ==="
    
    ' Zavøít UserForm
    Unload Me
    
    MsgBox "Promoce byla úspìšnì vložena! PromoID: " & promoID, vbInformation
    
    Debug.Print "=== InsertPromo END ==="
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "CHYBA v InsertPromo na øádku: " & Erl
    Debug.Print "Chyba " & Err.Number & ": " & Err.Description
    MsgBox "Chyba: " & Err.Description, vbCritical
End Sub

Private Function GetCountryCode() As String
    On Error Resume Next
    GetCountryCode = Trim(TargetWorkbook.Sheets("Settings").Range("B10").value)
    On Error GoTo 0
    
    If GetCountryCode = "" Then
        GetCountryCode = "CZK"  ' Default
    End If
End Function

'Private Sub LoadProducts()
'    On Error GoTo ErrorHandler
'
'    Debug.Print "=== LoadProducts START ==="
'
'    Dim SelectedValue As String
'    Dim i As Long
'    Dim loadPrices As Boolean
'
'    ' Zkontrolovat, že máme SelectedRange
'    If SelectedRange Is Nothing Then
'        Debug.Print "SelectedRange není nastaven"
'        Exit Sub
'    End If
'
'    ' Zjistit, jestli naèítat i ceny
'    loadPrices = (LB_Price.ListIndex >= 0)
'
'    If loadPrices Then
'        SelectedValue = LB_Price.value
'        Debug.Print "Vybraná cena: " & SelectedValue
'    Else
'        Debug.Print "Žádná cena nevybraná - naèítám jen produkty"
'    End If
'
'    ' Získat fcType z LB_FC
'    Dim fcType As String
'    If LB_FC.ListIndex >= 0 Then
'        fcType = LB_FC.value
'        Debug.Print "FC Type: " & fcType
'    Else
'        fcType = "AFC"
'        Debug.Print "FC Type: " & fcType & " (default)"
'    End If
'
'    ' PØIDÁNO: Naèíst countryCode z Settings
'    Dim countryCode As String
'    countryCode = GetCountryCode()
'    Debug.Print "Country Code: " & countryCode
'
'    ' Vyèištìní ListBoxù
'    LB_Product.Clear
'    LB_PriceValues.Clear
'    LB_AFC.Clear
'    LB_ZS.Clear
'
'    ' Naètení Products pole
'    Call ProductsArray(TargetWorkbook)
'
'    Dim selectedFamily As String
'    selectedFamily = SelectedRange.Worksheet.Cells(SelectedRange.row, 3).value
'    Debug.Print "Vybraná rodina: " & selectedFamily
'
'    ' Procházení kolekce produktù
'    Dim rowData As Object
'    For Each rowData In GetProductsCollection()
'
'        If rowData.Exists("Family") And rowData("Family") = selectedFamily Then
'
'            ' Podle countryCode rozhodnout formát productName
'            Dim productName As String
'            If UCase(Trim(countryCode)) = "SVK" Then
'                productName = rowData("material_name")  ' Bez volume_l
'            Else
'                productName = rowData("material_name") & " " & rowData("volume_l")  ' S volume_l (default pro CZK)
'            End If
'
'            LB_Product.AddItem productName
'
'            ' Naèíst ceny jen pokud je vybraná cena
'            If loadPrices Then
'                Dim result As Variant
'                result = GetPromoPriceData(selectedFamily, SelectedValue, rowData, fcType)
'                LB_PriceValues.AddItem result(0)
'                LB_AFC.AddItem result(3)
'                LB_ZS.AddItem result(2)
'            Else
'                ' Pøidat prázdné hodnoty
'                LB_PriceValues.AddItem ""
'                LB_AFC.AddItem ""
'                LB_ZS.AddItem ""
'            End If
'        End If
'    Next rowData
'
'    Debug.Print "Poèet produktù: " & LB_Product.ListCount
'
'    ' Vybrat všechny produkty
'    For i = 0 To LB_Product.ListCount - 1
'        LB_Product.Selected(i) = True
'    Next i
'
'    Debug.Print "=== LoadProducts END ==="
'    Exit Sub
'
'ErrorHandler:
'    Debug.Print "CHYBA v LoadProducts: " & Err.Description
'    MsgBox "Chyba pøi naèítání produktù: " & Err.Description, vbCritical
'End Sub

Private Sub LoadProducts()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== LoadProducts START ==="
    
    Dim SelectedValue As String
    Dim i As Long
    Dim loadPrices As Boolean
    
    ' Zkontrolovat, že máme SelectedRange
    If SelectedRange Is Nothing Then
        Debug.Print "SelectedRange není nastaven"
        Exit Sub
    End If
    
    ' Zjistit, jestli naèítat i ceny
    loadPrices = (LB_Price.ListIndex >= 0)
    
    If loadPrices Then
        SelectedValue = LB_Price.value
        Debug.Print "Vybraná cena: " & SelectedValue
    Else
        Debug.Print "Žádná cena nevybraná - naèítám jen produkty"
    End If
    
    ' Získat fcType z LB_FC
    Dim fcType As String
    If LB_FC.ListIndex >= 0 Then
        fcType = LB_FC.value
        Debug.Print "FC Type: " & fcType
    Else
        fcType = "AFC"
        Debug.Print "FC Type: " & fcType & " (default)"
    End If
    
    ' PØIDÁNO: Naèíst countryCode z Settings
    Dim countryCode As String
    countryCode = GetCountryCode()
    Debug.Print "Country Code: " & countryCode
    
    ' Vyèištìní ListBoxù
    LB_Product.Clear
    LB_PriceValues.Clear
    LB_AFC.Clear
    LB_ZS.Clear
    
    ' Naètení Products pole
    Call ProductsArray(TargetWorkbook)
    
    Dim selectedFamily As String
    selectedFamily = SelectedRange.Worksheet.Cells(SelectedRange.row, 3).value
    Debug.Print "Vybraná rodina: " & selectedFamily
    

    ' Procházení kolekce produktù
    Dim rowData As Object
    For Each rowData In GetProductsCollection()
        
        If rowData.Exists("Family") And rowData("Family") = selectedFamily Then
            
            ' Podle countryCode rozhodnout formát productName
            Dim productName As String
            If UCase(Trim(countryCode)) = "SVK" Then
                productName = rowData("material_name")
            Else
                productName = rowData("material_name") & " " & rowData("volume_l")
            End If
            
            LB_Product.AddItem productName
            
            ' NAÈÍTÁNÍ CEN - TADY PØIDAT DEBUG
            If loadPrices Then
                Dim result As Variant
                Debug.Print "Volám GetPromoPriceData s:"
                Debug.Print "  selectedFamily: " & selectedFamily
                Debug.Print "  SelectedValue: " & SelectedValue
                Debug.Print "  fcType: " & fcType
                
                result = GetPromoPriceData(selectedFamily, SelectedValue, rowData, fcType)
                
                Debug.Print "Výsledek:"
                Debug.Print "  result(0): " & result(0)
                Debug.Print "  result(2): " & result(2)
                Debug.Print "  result(3): " & result(3)
                
                LB_PriceValues.AddItem result(0)
                LB_AFC.AddItem result(3)
                LB_ZS.AddItem result(6)
            Else
                ' Pøidat prázdné hodnoty
                LB_PriceValues.AddItem ""
                LB_AFC.AddItem ""
                LB_ZS.AddItem ""
            End If
        End If
    Next rowData
    
    Debug.Print "Poèet produktù: " & LB_Product.ListCount
    
    ' Vybrat všechny produkty
    For i = 0 To LB_Product.ListCount - 1
        LB_Product.Selected(i) = True
    Next i
    
    Debug.Print "=== LoadProducts END ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "CHYBA v LoadProducts: " & Err.Description
    MsgBox "Chyba pøi naèítání produktù: " & Err.Description, vbCritical
End Sub

Public Sub LoadFCTypesToListBox()
    On Error GoTo ErrorHandler
    
    ' Vyèistit ListBox
    Me.LB_FC.Clear
    
    ' Zkontrolovat, zda máme TargetWorkbook
    If TargetWorkbook Is Nothing Then
        Debug.Print "TargetWorkbook není nastaven!"
        Exit Sub
    End If
    
    ' Zkusit naèíst list PromoConfig z TargetWorkbook (uživatelský soubor)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = TargetWorkbook.Sheets("PromoConfig")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Debug.Print "List 'PromoConfig' nebyl nalezen v " & TargetWorkbook.Name
        Exit Sub
    End If
    
    Debug.Print "List PromoConfig nalezen v: " & TargetWorkbook.Name
    
    ' Najít sloupec FC_Type (N nebo pojmenovaný rozsah)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, "N").End(xlUp).row
    
    Debug.Print "Poslední øádek ve sloupci FC_Type: " & lastRow
    
    If lastRow < 2 Then
        Debug.Print "Ve sloupci FC_Type nejsou žádná data!"
        Exit Sub
    End If
    
    ' Projít všechny hodnoty ve sloupci FC_Type (N) od øádku 2
    Dim i As Long
    Dim fcValue As String
    
    For i = 2 To lastRow
        fcValue = Trim(ws.Cells(i, "N").value)
        If fcValue <> "" Then
            Me.LB_FC.AddItem fcValue
            Debug.Print "  Pøidáno: " & fcValue
        End If
    Next i
    
    Debug.Print "Naèteno " & Me.LB_FC.ListCount & " hodnot do LB_FC"
    
    ' PØIDÁNO: Pokud je jen jedna hodnota, automaticky ji vybrat
    If Me.LB_FC.ListCount = 1 Then
        Me.LB_FC.ListIndex = 0
        Debug.Print "Automaticky vybrána jediná hodnota: " & Me.LB_FC.value
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "CHYBA v LoadFCTypesToListBox: " & Err.Description
    MsgBox "Chyba pøi naèítání FC_Type: " & Err.Description, vbCritical
End Sub
Private Sub UserForm_Initialize()
    ' Nastavit pouze základní vlastnosti ListBoxù
    ' NEPOUŽÍVAT TargetWorkbook nebo SelectedRange zde!
    
    Call LoadFCTypesToListBox
    With LB_Promoce
        .MultiSelect = fmMultiSelectSingle
        .ListStyle = fmListStyleOption
        .AddItem "Leták"
        .AddItem "Leták + Tichá"
        .AddItem "Tichá promoce"
        .AddItem "Titulka"
        .AddItem "Titulka + Tichá"
        .AddItem "WOW Page"
        .AddItem "WOW Page + Tichá"
        .AddItem "WOW okno"
        .AddItem "WOW okno + Tichá"
        .AddItem "1denní"
        .AddItem "Víkendová"
        .AddItem "Vklad"
    End With
    
    With LB_Price
        .MultiSelect = fmMultiSelectSingle
        .ListStyle = fmListStyleOption
        .AddItem "ANCD"
        .AddItem "TANCD"
        .AddItem "TANCD II"
        .AddItem "TANCD III"
    End With
    
    With LB_Product
        .MultiSelect = fmMultiSelectMulti
        .ListStyle = fmListStyleOption
    End With
    
    With LB_Hero
        .MultiSelect = fmMultiSelectSingle
        .ListStyle = fmListStyleOption
    End With
        
    With LB_FC
        .MultiSelect = fmMultiSelectSingle
        .ListStyle = fmListStyleOption
    End With
            
End Sub

Private Sub UserForm_Activate()
    Static initialized As Boolean
    
    If Not initialized Then
        Debug.Print "=== UserForm_Activate - první spuštìní ==="
        
        ' Naèíst FC typy
        If LB_FC.ListCount = 0 Then
            Call LoadFCTypesToListBox
        End If
        
        ' PØIDAT: Vybrat první cenu, pokud je jen jedna
        If LB_Price.ListCount = 1 Then
            LB_Price.ListIndex = 0
            Debug.Print "Automaticky vybrána jediná cena: " & LB_Price.value
        End If
        
        ' Naèíst produkty
        Call LoadProducts
        
        initialized = True
    End If
End Sub

Public Sub LoadData()
    On Error GoTo ErrorHandler
    
    ' Kontrola, že promìnné jsou nastavené
    If TargetWorkbook Is Nothing Then
        MsgBox "TargetWorkbook není nastaven!", vbCritical
        Exit Sub
    End If
    
    If SelectedRange Is Nothing Then
        MsgBox "SelectedRange není nastaven!", vbCritical
        Exit Sub
    End If
    
    ' Naète Products do kolekce
    Call ProductsArray(TargetWorkbook)
    
    ' Získá family hodnotu
    Dim selectedFamily As String
    selectedFamily = SelectedRange.Worksheet.Cells(SelectedRange.row, 3).value
    
    ' Naplní LB_Product produkty z dané family
    LB_Product.Clear
    
    Dim rowData As Object
    Dim productText As String
    
    For Each rowData In ProductsCollection
        If rowData.Exists("Family") Then
            If rowData("Family") = selectedFamily Then
                If rowData.Exists("material_name") And rowData.Exists("volume_l") Then
                    productText = rowData("material_name") & " " & rowData("volume_l")
                    LB_Product.AddItem productText
                End If
            End If
        End If
    Next rowData
    
    ' Oznaèí všechny produkty
    Dim j As Long
    For j = 0 To LB_Product.ListCount - 1
        LB_Product.Selected(j) = True
    Next j
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Chyba v LoadData: " & Err.Description, vbCritical
End Sub

Public Sub PromoSet()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== PromoSet START ==="
    
    ' Naèíst data
    Call ProductsArray(TargetWorkbook)
    Call WeeksArray(TargetWorkbook, SelectedRange)
    
    Debug.Print "Vytváøím Promo instanci..."
    Set Promo = CreatePromoInstance()
    
    If Promo Is Nothing Then
        MsgBox "Nepodaøilo se vytvoøit Promo instanci!", vbCritical
        Exit Sub
    End If
    
    Debug.Print "Promo vytvoøeno: " & Not (Promo Is Nothing)
    
    Dim selectedPromo As String
    selectedPromo = Me.LB_Promoce.value
    
    Debug.Print "Vybraná promoce z ListBoxu: " & selectedPromo
    
    ' ZMÌNA: Volat sdílenou funkci a pøedat flag pro plán
    Dim usePlanColor As Boolean
    usePlanColor = Me.CB_Plan.value
    
    If Not SetupPromoByListBoxValue_Shared(selectedPromo, SelectedRange, Promo, TargetWorkbook, usePlanColor) Then
        MsgBox "Chyba pøi nastavení promoce!", vbCritical
        Exit Sub
    End If
    
    Debug.Print "=== PromoSet END ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "CHYBA v PromoSet: " & Err.Description & " na øádku " & Erl
    MsgBox "Chyba v PromoSet: " & Err.Description
End Sub

Private Sub LB_Price_Change()
    ' Naèíst produkty s novými cenami
    Call LoadProducts
End Sub

Function GetRGBColor(r As Long, g As Long, b As Long) As Variant
    ' Kontrola, zda je CheckBox 'CB_Plan' zaškrtnutý
    If CB_Plan.value = True Then
        ' Vrátí jednotnou šedou barvu
        GetRGBColor = Array(180, 180, 180)
    Else
        ' Vrátí pùvodní barvy
        GetRGBColor = Array(r, g, b)
    End If
End Function

' Wrapper pro GetFamilyByEAN
Public Function CallGetFamilyByEAN(eanValue As String) As Variant
    Call EnsureSharedCodeOpen
    CallGetFamilyByEAN = Application.Run("SharedCode.xlsm!GetFamilyByEAN", eanValue)
End Function

' Wrapper pro GetCustomerIDByEAN
Public Function CallGetCustomerIDByEAN(eanValue As String) As Variant
    Call EnsureSharedCodeOpen
    CallGetCustomerIDByEAN = Application.Run("SharedCode.xlsm!GetCustomerIDByEAN", eanValue)
End Function

' Wrapper pro GetBrandByEAN
Public Function CallGetBrandByEAN(eanValue As String) As Variant
    Call EnsureSharedCodeOpen
    CallGetBrandByEAN = Application.Run("SharedCode.xlsm!GetBrandByEAN", eanValue)
End Function
