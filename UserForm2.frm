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
        MsgBox "Nejsou vybr�ny v�echny povinn� �daje."
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
        InsertPromo  ' Nyn� tato procedura existuje n�e
    End If
End Sub

Private Sub InsertPromo()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== InsertPromo START ==="
    
    ' Zavolat PromoSet
    Call PromoSet
    
    ' Ov��en�, �e Promo je vytvo�en�
    If Promo Is Nothing Then
        MsgBox "Chyba: Promo nebyl vytvo�en!"
        Exit Sub
    End If
    
    Debug.Print "Promo typ: " & Promo.promoTyp
    
    ' Generovat PromoID
    Dim promoID As String
    promoID = GenerateID(TargetWorkbook)
    Debug.Print "Vygenerovan� PromoID: " & promoID
    
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
    
    ' V�b�r
    Dim vyberValue As String
    vyberValue = GetVyberValue(Me.LB_Product)
    Debug.Print "Vyber Value: " & vyberValue
    
    ' Hero
    Dim heroProduct As String
    heroProduct = GetHeroItem(Me.LB_Hero)
    heroProduct = Me.LB_Product.List(Me.LB_Product.ListIndex)
    Debug.Print "Hero Product: " & heroProduct
    
    ' Pl�n
    Dim isPlan As Boolean
    isPlan = CB_Plan.value
    Debug.Print "Is Plan: " & isPlan
    
    ' Na��st countryCode z Settings
    Dim countryCode As String
    countryCode = GetCountryCode()
    Debug.Print "Country Code: " & countryCode
       
    Dim commentText As String
    commentText = Trim(Me.TB_Comment.value)
    Debug.Print "Comment: " & commentText
       
    Debug.Print "=== P�ED VOL�N�M PridejVybraneHeroProdukty ==="
    
    ' Z�pis do listu Text
    Call PridejVybraneHeroProdukty(Me, selectedPrice, Promo, heroProduct, promoID, vyberValue, pcsPlanText, isPlan, TargetWorkbook, SelectedRange, Me.LB_FC.value, countryCode, commentText)
    
    Debug.Print "=== PO VOL�N� PridejVybraneHeroProdukty ==="
    
    ' P�ESUNUTO: Se�azen� P�ED form�tov�n�m
    Debug.Print "=== P�ED Se�azen�m ==="
    Call ApplyFilterToRow2(TargetWorkbook)
    Call SortIt(TargetWorkbook)
    Debug.Print "=== PO Se�azen� ==="
    
    ' Barven� ��dk�
    Call rColor(TargetWorkbook)
    Debug.Print "=== PO rColor ==="
    
        Debug.Print "=== KONTROLA KOMENT��� P�ED FormatPromoCells ==="
    Dim checkCell As Range
    For Each checkCell In SelectedRange.Cells
        If Not checkCell.comment Is Nothing Then
            Debug.Print "  Bu�ka " & checkCell.Address & " m� koment��: " & Left(checkCell.comment.Text, 8)
        Else
            Debug.Print "  Bu�ka " & checkCell.Address & " NEM� koment��"
        End If
    Next checkCell
    Debug.Print "=== KONEC KONTROLY ==="
    
    ' Form�tov�n� kosti�ek
    Dim usePlanColor As Boolean
    usePlanColor = CB_Plan.value
    Call FormatPromoCells(TargetWorkbook, SelectedRange, Promo, promoID, usePlanColor)
        
    Debug.Print "=== PO FormatPromoCells ==="
    
    ' Zav��t UserForm
    Unload Me
    
    MsgBox "Promoce byla �sp�n� vlo�ena! PromoID: " & promoID, vbInformation
    
    Debug.Print "=== InsertPromo END ==="
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "CHYBA v InsertPromo na ��dku: " & Erl
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
'    ' Zkontrolovat, �e m�me SelectedRange
'    If SelectedRange Is Nothing Then
'        Debug.Print "SelectedRange nen� nastaven"
'        Exit Sub
'    End If
'
'    ' Zjistit, jestli na��tat i ceny
'    loadPrices = (LB_Price.ListIndex >= 0)
'
'    If loadPrices Then
'        SelectedValue = LB_Price.value
'        Debug.Print "Vybran� cena: " & SelectedValue
'    Else
'        Debug.Print "��dn� cena nevybran� - na��t�m jen produkty"
'    End If
'
'    ' Z�skat fcType z LB_FC
'    Dim fcType As String
'    If LB_FC.ListIndex >= 0 Then
'        fcType = LB_FC.value
'        Debug.Print "FC Type: " & fcType
'    Else
'        fcType = "AFC"
'        Debug.Print "FC Type: " & fcType & " (default)"
'    End If
'
'    ' P�ID�NO: Na��st countryCode z Settings
'    Dim countryCode As String
'    countryCode = GetCountryCode()
'    Debug.Print "Country Code: " & countryCode
'
'    ' Vy�i�t�n� ListBox�
'    LB_Product.Clear
'    LB_PriceValues.Clear
'    LB_AFC.Clear
'    LB_ZS.Clear
'
'    ' Na�ten� Products pole
'    Call ProductsArray(TargetWorkbook)
'
'    Dim selectedFamily As String
'    selectedFamily = SelectedRange.Worksheet.Cells(SelectedRange.row, 3).value
'    Debug.Print "Vybran� rodina: " & selectedFamily
'
'    ' Proch�zen� kolekce produkt�
'    Dim rowData As Object
'    For Each rowData In GetProductsCollection()
'
'        If rowData.Exists("Family") And rowData("Family") = selectedFamily Then
'
'            ' Podle countryCode rozhodnout form�t productName
'            Dim productName As String
'            If UCase(Trim(countryCode)) = "SVK" Then
'                productName = rowData("material_name")  ' Bez volume_l
'            Else
'                productName = rowData("material_name") & " " & rowData("volume_l")  ' S volume_l (default pro CZK)
'            End If
'
'            LB_Product.AddItem productName
'
'            ' Na��st ceny jen pokud je vybran� cena
'            If loadPrices Then
'                Dim result As Variant
'                result = GetPromoPriceData(selectedFamily, SelectedValue, rowData, fcType)
'                LB_PriceValues.AddItem result(0)
'                LB_AFC.AddItem result(3)
'                LB_ZS.AddItem result(2)
'            Else
'                ' P�idat pr�zdn� hodnoty
'                LB_PriceValues.AddItem ""
'                LB_AFC.AddItem ""
'                LB_ZS.AddItem ""
'            End If
'        End If
'    Next rowData
'
'    Debug.Print "Po�et produkt�: " & LB_Product.ListCount
'
'    ' Vybrat v�echny produkty
'    For i = 0 To LB_Product.ListCount - 1
'        LB_Product.Selected(i) = True
'    Next i
'
'    Debug.Print "=== LoadProducts END ==="
'    Exit Sub
'
'ErrorHandler:
'    Debug.Print "CHYBA v LoadProducts: " & Err.Description
'    MsgBox "Chyba p�i na��t�n� produkt�: " & Err.Description, vbCritical
'End Sub

Private Sub LoadProducts()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== LoadProducts START ==="
    
    Dim SelectedValue As String
    Dim i As Long
    Dim loadPrices As Boolean
    
    ' Zkontrolovat, �e m�me SelectedRange
    If SelectedRange Is Nothing Then
        Debug.Print "SelectedRange nen� nastaven"
        Exit Sub
    End If
    
    ' Zjistit, jestli na��tat i ceny
    loadPrices = (LB_Price.ListIndex >= 0)
    
    If loadPrices Then
        SelectedValue = LB_Price.value
        Debug.Print "Vybran� cena: " & SelectedValue
    Else
        Debug.Print "��dn� cena nevybran� - na��t�m jen produkty"
    End If
    
    ' Z�skat fcType z LB_FC
    Dim fcType As String
    If LB_FC.ListIndex >= 0 Then
        fcType = LB_FC.value
        Debug.Print "FC Type: " & fcType
    Else
        fcType = "AFC"
        Debug.Print "FC Type: " & fcType & " (default)"
    End If
    
    ' P�ID�NO: Na��st countryCode z Settings
    Dim countryCode As String
    countryCode = GetCountryCode()
    Debug.Print "Country Code: " & countryCode
    
    ' Vy�i�t�n� ListBox�
    LB_Product.Clear
    LB_PriceValues.Clear
    LB_AFC.Clear
    LB_ZS.Clear
    
    ' Na�ten� Products pole
    Call ProductsArray(TargetWorkbook)
    
    Dim selectedFamily As String
    selectedFamily = SelectedRange.Worksheet.Cells(SelectedRange.row, 3).value
    Debug.Print "Vybran� rodina: " & selectedFamily
    

    ' Proch�zen� kolekce produkt�
    Dim rowData As Object
    For Each rowData In GetProductsCollection()
        
        If rowData.Exists("Family") And rowData("Family") = selectedFamily Then
            
            ' Podle countryCode rozhodnout form�t productName
            Dim productName As String
            productName = GetProductName(rowData, countryCode)
            
            LB_Product.AddItem productName
            
            ' NA��T�N� CEN - TADY P�IDAT DEBUG
            If loadPrices Then
                Dim result As Variant
                Debug.Print "Vol�m GetPromoPriceData s:"
                Debug.Print "  selectedFamily: " & selectedFamily
                Debug.Print "  SelectedValue: " & SelectedValue
                Debug.Print "  fcType: " & fcType
                
                result = GetPromoPriceData(selectedFamily, SelectedValue, rowData, fcType)
                
                Debug.Print "V�sledek:"
                Debug.Print "  result(0): " & result(0)
                Debug.Print "  result(2): " & result(2)
                Debug.Print "  result(3): " & result(3)
                
                LB_PriceValues.AddItem result(0)
                LB_AFC.AddItem result(3)
                LB_ZS.AddItem result(6)
            Else
                ' P�idat pr�zdn� hodnoty
                LB_PriceValues.AddItem ""
                LB_AFC.AddItem ""
                LB_ZS.AddItem ""
            End If
        End If
    Next rowData
    
    Debug.Print "Po�et produkt�: " & LB_Product.ListCount
    
    ' Vybrat v�echny produkty
    For i = 0 To LB_Product.ListCount - 1
        LB_Product.Selected(i) = True
    Next i
    
    Debug.Print "=== LoadProducts END ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "CHYBA v LoadProducts: " & Err.Description
    MsgBox "Chyba p�i na��t�n� produkt�: " & Err.Description, vbCritical
End Sub

Public Sub LoadFCTypesToListBox()
    On Error GoTo ErrorHandler
    
    ' Vy�istit ListBox
    Me.LB_FC.Clear
    
    ' Zkontrolovat, zda m�me TargetWorkbook
    If TargetWorkbook Is Nothing Then
        Debug.Print "TargetWorkbook nen� nastaven!"
        Exit Sub
    End If
    
    ' Zkusit na��st list PromoConfig z TargetWorkbook (u�ivatelsk� soubor)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = TargetWorkbook.Sheets("PromoConfig")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Debug.Print "List 'PromoConfig' nebyl nalezen v " & TargetWorkbook.Name
        Exit Sub
    End If
    
    Debug.Print "List PromoConfig nalezen v: " & TargetWorkbook.Name
    
    ' Naj�t sloupec FC_Type (N nebo pojmenovan� rozsah)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, "N").End(xlUp).row
    
    Debug.Print "Posledn� ��dek ve sloupci FC_Type: " & lastRow
    
    If lastRow < 2 Then
        Debug.Print "Ve sloupci FC_Type nejsou ��dn� data!"
        Exit Sub
    End If
    
    ' Proj�t v�echny hodnoty ve sloupci FC_Type (N) od ��dku 2
    Dim i As Long
    Dim fcValue As String
    
    For i = 2 To lastRow
        fcValue = Trim(ws.Cells(i, "N").value)
        If fcValue <> "" Then
            Me.LB_FC.AddItem fcValue
            Debug.Print "  P�id�no: " & fcValue
        End If
    Next i
    
    Debug.Print "Na�teno " & Me.LB_FC.ListCount & " hodnot do LB_FC"
    
    ' P�ID�NO: Pokud je jen jedna hodnota, automaticky ji vybrat
    If Me.LB_FC.ListCount = 1 Then
        Me.LB_FC.ListIndex = 0
        Debug.Print "Automaticky vybr�na jedin� hodnota: " & Me.LB_FC.value
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "CHYBA v LoadFCTypesToListBox: " & Err.Description
    MsgBox "Chyba p�i na��t�n� FC_Type: " & Err.Description, vbCritical
End Sub
Private Sub UserForm_Initialize()
    ' Nastavit pouze z�kladn� vlastnosti ListBox�
    ' NEPOU��VAT TargetWorkbook nebo SelectedRange zde!
    
    Call LoadFCTypesToListBox
    With LB_Promoce
        .MultiSelect = fmMultiSelectSingle
        .ListStyle = fmListStyleOption
        .AddItem "Let�k"
        .AddItem "Let�k + Tich�"
        .AddItem "Tich� promoce"
        .AddItem "Titulka"
        .AddItem "Titulka + Tich�"
        .AddItem "WOW Page"
        .AddItem "WOW Page + Tich�"
        .AddItem "WOW okno"
        .AddItem "WOW okno + Tich�"
        .AddItem "1denn�"
        .AddItem "V�kendov�"
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
        Debug.Print "=== UserForm_Activate - prvn� spu�t�n� ==="
        
        ' Na��st FC typy
        If LB_FC.ListCount = 0 Then
            Call LoadFCTypesToListBox
        End If
        
        ' P�IDAT: Vybrat prvn� cenu, pokud je jen jedna
        If LB_Price.ListCount = 1 Then
            LB_Price.ListIndex = 0
            Debug.Print "Automaticky vybr�na jedin� cena: " & LB_Price.value
        End If
        
        ' Na��st produkty
        Call LoadProducts
        
        initialized = True
    End If
End Sub

Public Sub LoadData()
    On Error GoTo ErrorHandler
    
    ' Kontrola, �e prom�nn� jsou nastaven�
    If TargetWorkbook Is Nothing Then
        MsgBox "TargetWorkbook nen� nastaven!", vbCritical
        Exit Sub
    End If
    
    If SelectedRange Is Nothing Then
        MsgBox "SelectedRange nen� nastaven!", vbCritical
        Exit Sub
    End If
    
    ' Na�te Products do kolekce
    Call ProductsArray(TargetWorkbook)
    
    ' Z�sk� family hodnotu
    Dim selectedFamily As String
    selectedFamily = SelectedRange.Worksheet.Cells(SelectedRange.row, 3).value
    
    ' Napln� LB_Product produkty z dan� family
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
    
    ' Ozna�� v�echny produkty
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
    
    ' Na��st data
    Call ProductsArray(TargetWorkbook)
    Call WeeksArray(TargetWorkbook, SelectedRange)
    
    Debug.Print "Vytv���m Promo instanci..."
    Set Promo = CreatePromoInstance()
    
    If Promo Is Nothing Then
        MsgBox "Nepoda�ilo se vytvo�it Promo instanci!", vbCritical
        Exit Sub
    End If
    
    Debug.Print "Promo vytvo�eno: " & Not (Promo Is Nothing)
    
    Dim selectedPromo As String
    selectedPromo = Me.LB_Promoce.value
    
    Debug.Print "Vybran� promoce z ListBoxu: " & selectedPromo
    
    ' ZM�NA: Volat sd�lenou funkci a p�edat flag pro pl�n
    Dim usePlanColor As Boolean
    usePlanColor = Me.CB_Plan.value
    
    If Not SetupPromoByListBoxValue_Shared(selectedPromo, SelectedRange, Promo, TargetWorkbook, usePlanColor) Then
        MsgBox "Chyba p�i nastaven� promoce!", vbCritical
        Exit Sub
    End If
    
    Debug.Print "=== PromoSet END ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "CHYBA v PromoSet: " & Err.Description & " na ��dku " & Erl
    MsgBox "Chyba v PromoSet: " & Err.Description
End Sub

Private Sub LB_Price_Change()
    ' Na��st produkty s nov�mi cenami
    Call LoadProducts
End Sub

Function GetRGBColor(r As Long, g As Long, b As Long) As Variant
    ' Kontrola, zda je CheckBox 'CB_Plan' za�krtnut�
    If CB_Plan.value = True Then
        ' Vr�t� jednotnou �edou barvu
        GetRGBColor = Array(180, 180, 180)
    Else
        ' Vr�t� p�vodn� barvy
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
