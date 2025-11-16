Attribute VB_Name = "SetupPromo"
Public Function SetupPromoByListBoxValue_Shared(listBoxValue As String, SelectedRange As Range, PromoObj As Object, TargetWorkbook As Workbook, Optional usePlanColor As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    
    Debug.Print "=== SetupPromoByListBoxValue_Shared START pro: " & listBoxValue & " ==="
    
    ' 1. NAÈÍST KONFIGURACI Z UŽIVATELSKÉHO SOUBORU
    Dim promoConfig As Object
    Set promoConfig = LoadPromoConfig(TargetWorkbook, listBoxValue)
    
    If promoConfig Is Nothing Then
        MsgBox "Konfigurace pro promoci '" & listBoxValue & "' nebyla nalezena!", vbCritical
        SetupPromoByListBoxValue_Shared = False
        Exit Function
    End If
    
    Debug.Print "Konfigurace naètena, kontrola hodnot:"
    Debug.Print "  PromoName: " & promoConfig("PromoName")
    Debug.Print "  R: " & promoConfig("R")
    Debug.Print "  G: " & promoConfig("G")
    Debug.Print "  B: " & promoConfig("B")
    Debug.Print "  TypAkce: " & promoConfig("TypAkce")
    Debug.Print "  PromoTyp: " & promoConfig("PromoTyp")
    Debug.Print "  StartWeekOffset: " & promoConfig("StartWeekOffset")
    
    ' 2. ZÍSKAT ZÁKLADNÍ DATUMY
    Debug.Print "Získávám StartWeek a EndWeek..."
    Dim StartWeek As Date, endWeek As Date
    StartWeek = GetStartWeek()
    endWeek = GetEndWeek()
    Debug.Print "  StartWeek: " & StartWeek
    Debug.Print "  endWeek: " & endWeek
    
    ' 3. ZÍSKAT WEEK INTERVALS
    Debug.Print "Získávám weekInterval..."
    Dim weekInterval As String, weekIntervalT As String
    Call GetWeekIntervalsFromSelection(SelectedRange, weekInterval, weekIntervalT, TargetWorkbook)
    Debug.Print "  weekInterval: " & weekInterval
    Debug.Print "  weekIntervalT: " & weekIntervalT
    
       
    ' 4. BARVY (z konfigurace nebo plán)
    Dim tempR As Long, tempG As Long, tempB As Long
    If usePlanColor Then
        tempR = 180: tempG = 180: tempB = 180
    Else
        tempR = promoConfig("R")
        tempG = promoConfig("G")
        tempB = promoConfig("B")
    End If
    
    ' 5. VYPOÈÍTAT DATUMY S OFFSETY
Dim calcStartWeek As Date, calcEndWeek As Date
Dim calcStartPurchase As Date, calcEndPurchase As Date
Dim calcSortFrom As Date, calcSortTo As Date

' OPRAVA: Použít DateAdd místo pøímého pøièítání
calcStartWeek = DateAdd("d", promoConfig("StartWeekOffset"), StartWeek)
calcEndWeek = DateAdd("d", promoConfig("EndWeekOffset"), endWeek)
calcStartPurchase = DateAdd("d", promoConfig("StartPurchaseOffset"), StartWeek)
calcEndPurchase = DateAdd("d", promoConfig("EndPurchaseOffset"), endWeek)
calcSortFrom = DateAdd("d", promoConfig("SortFromOffset"), StartWeek)
calcSortTo = DateAdd("d", promoConfig("SortToOffset"), endWeek)

Debug.Print "Vypoèítané datumy:"
Debug.Print "  calcStartWeek: " & calcStartWeek
Debug.Print "  calcEndWeek: " & calcEndWeek

' 6. NASTAVIT PROMO
Debug.Print "Pøed voláním PromoSettings:"
Debug.Print "  TypAkce: '" & promoConfig("TypAkce") & "'"
Debug.Print "  PromoTyp: '" & promoConfig("PromoTyp") & "'"
Debug.Print "  weekInterval: '" & weekInterval & "'"
Debug.Print "  weekIntervalT: '" & weekIntervalT & "'"

On Error Resume Next
Call PromoObj.PromoSettings( _
    promoConfig("TypAkce"), _
    promoConfig("PromoTyp"), _
    weekInterval, _
    weekIntervalT, _
    calcStartWeek, _
    calcEndWeek, _
    calcStartPurchase, _
    calcEndPurchase, _
    calcSortFrom, _
    calcSortTo, _
    tempR, _
    tempG, _
    tempB, _
    vbWhite _
)

If Err.Number <> 0 Then
    Debug.Print "CHYBA pøi volání PromoSettings: " & Err.Description
    Debug.Print "Err.Number: " & Err.Number
    On Error GoTo ErrorHandler
    Err.Raise Err.Number, , Err.Description
End If
On Error GoTo ErrorHandler

Debug.Print "Po volání PromoSettings - Promo.promoTyp: '" & PromoObj.promoTyp & "'"

Debug.Print "Po volání PromoSettings - Promo.promoTyp: " & PromoObj.promoTyp
    SetupPromoByListBoxValue_Shared = True
    Debug.Print "=== SetupPromoByListBoxValue_Shared END ==="
    Exit Function
    
ErrorHandler:
    Debug.Print "CHYBA v SetupPromoByListBoxValue_Shared: " & Err.Description
    SetupPromoByListBoxValue_Shared = False
End Function

Private Function LoadPromoConfig(TargetWorkbook As Workbook, promoName As String) As Object
    On Error GoTo ErrorHandler
    
    ' Zkusit najít list PromoConfig
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = TargetWorkbook.Sheets("PromoConfig")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Debug.Print "List 'PromoConfig' nebyl nalezen!"
        Set LoadPromoConfig = Nothing
        Exit Function
    End If
    
    ' Najít øádek s danou promocí
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    
    Dim i As Long
    For i = 2 To lastRow  ' Øádek 1 = hlavièka
        ' Hledat podle PromoName (sloupec A) NEBO PromoTyp (sloupec F)
        If Trim(ws.Cells(i, 1).value) = promoName Or Trim(ws.Cells(i, 6).value) = promoName Then
            ' Našli jsme! Vytvoøit Dictionary s daty
            Dim config As Object
            Set config = CreateObject("Scripting.Dictionary")
            
            config.Add "PromoName", ws.Cells(i, 1).value
            config.Add "R", CLng(ws.Cells(i, 2).value)
            config.Add "G", CLng(ws.Cells(i, 3).value)
            config.Add "B", CLng(ws.Cells(i, 4).value)
            config.Add "TypAkce", CStr(ws.Cells(i, 5).value)
            config.Add "PromoTyp", CStr(ws.Cells(i, 6).value)
            config.Add "StartWeekOffset", CLng(ws.Cells(i, 7).value)
            config.Add "EndWeekOffset", CLng(ws.Cells(i, 8).value)
            config.Add "StartPurchaseOffset", CLng(ws.Cells(i, 9).value)
            config.Add "EndPurchaseOffset", CLng(ws.Cells(i, 10).value)
            config.Add "SortFromOffset", CLng(ws.Cells(i, 11).value)
            config.Add "SortToOffset", CLng(ws.Cells(i, 12).value)
            
            Set LoadPromoConfig = config
            Debug.Print "Naètena konfigurace pro: " & promoName
            Exit Function
        End If
    Next i
    
    ' Nenalezeno
    Debug.Print "Konfigurace pro '" & promoName & "' nebyla nalezena"
    Set LoadPromoConfig = Nothing
    Exit Function
    
ErrorHandler:
    Debug.Print "CHYBA v LoadPromoConfig: " & Err.Description
    Set LoadPromoConfig = Nothing
End Function
'Public Function SetupPromoByListBoxValue_Shared(listBoxValue As String, SelectedRange As Range, PromoObj As Object, TargetWorkbook As Workbook, Optional usePlanColor As Boolean = False) As Boolean
'    On Error GoTo ErrorHandler
'
'    Debug.Print "=== SetupPromoByListBoxValue_Shared START pro: " & listBoxValue & " ==="
'
'    Dim StartWeek As Date, endWeek As Date
'    StartWeek = GetStartWeek()
'    endWeek = GetEndWeek()
'
'    Dim weekInterval As String, weekIntervalT As String
'    Call GetWeekIntervalsFromSelection(SelectedRange, weekInterval, weekIntervalT, TargetWorkbook)
'
'    Dim Color As Variant
'    Dim tempR As Long, tempG As Long, tempB As Long
'
'    Select Case listBoxValue
'        Case "Leták", "L"
'            Color = GetRGBColorShared(100, 230, 100, False)
'            tempR = Color(0): tempG = Color(1): tempB = Color(2)
'            Call PromoObj.PromoSettings("L", "L", weekInterval, weekIntervalT, _
'                StartWeek, endWeek, StartWeek - 14, endWeek, StartWeek, endWeek, _
'                tempR, tempG, tempB, vbWhite)
'
'        Case "Leták + Tichá", "L+T"
'            Color = GetRGBColorShared(100, 230, 100, False)
'            tempR = Color(0): tempG = Color(1): tempB = Color(2)
'            Call PromoObj.PromoSettings("L", "L+T", weekInterval, weekIntervalT, _
'                StartWeek, endWeek, StartWeek - 14, endWeek, StartWeek, endWeek, _
'                tempR, tempG, tempB, vbWhite)
'
'        Case "Tichá promoce", "T"
'            Color = GetRGBColorShared(110, 50, 160, False)
'            tempR = Color(0): tempG = Color(1): tempB = Color(2)
'            Call PromoObj.PromoSettings("T", "T", weekInterval, weekIntervalT, _
'                StartWeek, endWeek, StartWeek - 14, endWeek, StartWeek, endWeek, _
'                tempR, tempG, tempB, vbWhite)
'
'        Case "Titulka", "FP"
'            Color = GetRGBColorShared(255, 190, 0, False)
'            tempR = Color(0): tempG = Color(1): tempB = Color(2)
'            Call PromoObj.PromoSettings("L", "FP", weekInterval, weekIntervalT, _
'                StartWeek, endWeek, StartWeek - 14, endWeek, StartWeek, endWeek, _
'                tempR, tempG, tempB, vbWhite)
'
'        Case "Titulka + Tichá", "FP+T"
'            Color = GetRGBColorShared(255, 190, 0, False)
'            tempR = Color(0): tempG = Color(1): tempB = Color(2)
'            Call PromoObj.PromoSettings("L", "FP+T", weekInterval, weekIntervalT, _
'                StartWeek, endWeek, StartWeek - 14, endWeek, StartWeek, endWeek, _
'                tempR, tempG, tempB, vbWhite)
'
'        Case "WOW Page", "WP"
'            Color = GetRGBColorShared(145, 210, 80, False)
'            tempR = Color(0): tempG = Color(1): tempB = Color(2)
'            Call PromoObj.PromoSettings("L", "WP", weekInterval, weekIntervalT, _
'                StartWeek, endWeek, StartWeek - 14, endWeek, StartWeek, endWeek, _
'                tempR, tempG, tempB, vbWhite)
'
'        Case "WOW Page + Tichá", "WP+T"
'            Color = GetRGBColorShared(145, 210, 80, False)
'            tempR = Color(0): tempG = Color(1): tempB = Color(2)
'            Call PromoObj.PromoSettings("L", "WP+T", weekInterval, weekIntervalT, _
'                StartWeek, endWeek, StartWeek - 14, endWeek, StartWeek, endWeek, _
'                tempR, tempG, tempB, vbWhite)
'
'        Case "WOW okno", "WO"
'            Color = GetRGBColorShared(255, 80, 80, False)
'            tempR = Color(0): tempG = Color(1): tempB = Color(2)
'            Call PromoObj.PromoSettings("L", "WO", weekInterval, weekIntervalT, _
'                StartWeek, endWeek, StartWeek - 14, endWeek, StartWeek, endWeek, _
'                tempR, tempG, tempB, vbWhite)
'
'        Case "WOW okno + Tichá", "WO+T"
'            Color = GetRGBColorShared(255, 80, 80, False)
'            tempR = Color(0): tempG = Color(1): tempB = Color(2)
'            Call PromoObj.PromoSettings("L", "WO+T", weekInterval, weekIntervalT, _
'                StartWeek, endWeek, StartWeek - 14, endWeek, StartWeek, endWeek, _
'                tempR, tempG, tempB, vbWhite)
'
'        Case "1denní", "1D"
'            Color = GetRGBColorShared(0, 175, 240, False)
'            tempR = Color(0): tempG = Color(1): tempB = Color(2)
'            Call PromoObj.PromoSettings("T", "1D", weekInterval, weekIntervalT, _
'                StartWeek, endWeek - 6, StartWeek - 14, endWeek - 6, StartWeek, endWeek - 6, _
'                tempR, tempG, tempB, vbWhite)
'
'        Case "Víkendová", "Vík"
'            Color = GetRGBColorShared(0, 175, 240, False)
'            tempR = Color(0): tempG = Color(1): tempB = Color(2)
'            Call PromoObj.PromoSettings("T", "Vík", weekInterval, weekIntervalT, _
'                StartWeek + 3, endWeek - 2, StartWeek - 12, endWeek - 2, StartWeek + 3, endWeek - 2, _
'                tempR, tempG, tempB, vbWhite)
'
'        Case "Vklad"
'            Color = GetRGBColorShared(255, 180, 0, False)
'            tempR = Color(0): tempG = Color(1): tempB = Color(2)
'            Call PromoObj.PromoSettings("L", "Vklad", weekInterval, weekIntervalT, _
'                StartWeek, endWeek, StartWeek - 14, endWeek, StartWeek, endWeek, _
'                tempR, tempG, tempB, vbWhite)
'
'        Case Else
'            Debug.Print "VAROVÁNÍ: Neznámá promoce '" & listBoxValue & "', použito výchozí nastavení"
'            Color = GetRGBColorShared(200, 200, 200, False)
'            tempR = Color(0): tempG = Color(1): tempB = Color(2)
'            Call PromoObj.PromoSettings("L", listBoxValue, weekInterval, weekIntervalT, _
'                StartWeek, endWeek, StartWeek - 14, endWeek, StartWeek, endWeek, _
'                tempR, tempG, tempB, vbWhite)
'    End Select
'
'    SetupPromoByListBoxValue_Shared = True
'    Debug.Print "=== SetupPromoByListBoxValue_Shared END ==="
'    Exit Function
'
'ErrorHandler:
'    Debug.Print "CHYBA v SetupPromoByListBoxValue_Shared: " & Err.Description
'    SetupPromoByListBoxValue_Shared = False
'End Function
'
'' Pomocná funkce pro barvy (bez CheckBoxu)
'Private Function GetRGBColorShared(r As Long, g As Long, b As Long, usePlanColor As Boolean) As Variant
'    If usePlanColor Then
'        GetRGBColorShared = Array(180, 180, 180)
'    Else
'        GetRGBColorShared = Array(r, g, b)
'    End If
'End Function

Public Sub GetWeekIntervalsFromSelection(SelectedRange As Range, ByRef weekInterval As String, ByRef weekIntervalT As String, TargetWorkbook As Workbook)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = SelectedRange.Worksheet

    Dim weekRowNumber As Long, weekRowNumberT As Long
    weekRowNumber = FindWeekRow(ws)

    If weekRowNumber = 0 Then
        Err.Raise vbObjectError + 2, "GetWeekIntervalsFromSelection", "Øádek s komentáøem 'WeekRow' nebyl nalezen!"
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

