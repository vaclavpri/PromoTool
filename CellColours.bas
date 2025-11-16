Attribute VB_Name = "CellColours"
' ===================================================================
' FormatPromoCells - S kontrolou priorit
' ===================================================================
Public Sub FormatPromoCells(TargetWorkbook As Workbook, SelectedRange As Range, PromoObj As Object, promoID As String, usePlanColor As Boolean)
    On Error GoTo ErrorHandler
    
    Debug.Print "=== FormatPromoCells START ==="
    Debug.Print "PromoID: " & promoID
    Debug.Print "Range: " & SelectedRange.Address
    
    ' KONTROLA MERGED BUNÃK - PÿID¡NO
    Debug.Print "=== KONTROLA MERGED BUNÃK ==="
    Dim cell As Range
    For Each cell In SelectedRange.Cells
        Debug.Print "  BuÚka " & cell.Address & ":"
        Debug.Print "    MergeCells: " & cell.MergeCells
        If cell.MergeCells Then
            Debug.Print "    MergeArea: " & cell.MergeArea.Address
        End If
    Next cell
    Debug.Print "==========================="
    
    
    Dim r As Long, g As Long, b As Long
    
    ' ZÌsk·nÌ barev
    If usePlanColor Then
        r = 180: g = 180: b = 180
    Else
        r = PromoObj.r
        g = PromoObj.g
        b = PromoObj.b
    End If
    
    ' Priorita novÈ promoce
    Dim newPriority As Long
    newPriority = GetPromoPriority(PromoObj.typAkce)
    Debug.Print "Priorita novÈ promoce: " & newPriority
    
    ' KROK 1: ProjÌt vöechny buÚky a rozhodnout, kterÈ se majÌ p¯epsat
    Dim cellsToOverwrite As Collection
    Set cellsToOverwrite = New Collection
    
    'Dim cell As Range
    Dim cellCounter As Long
    cellCounter = 0
    
    For Each cell In SelectedRange.Cells
        cellCounter = cellCounter + 1
        Debug.Print "  Kontrola buÚky #" & cellCounter & ": " & cell.Address
        
        Dim shouldOverwrite As Boolean
        shouldOverwrite = True
        
        If Not cell.comment Is Nothing Then
            Dim existingPromoID As String
            existingPromoID = Left(cell.comment.Text, 8)
            Debug.Print "    ExistujÌcÌ PromoID: " & existingPromoID
            
            ' P¯eskoËit, pokud je to stejnÈ ID
            If existingPromoID = promoID Then
                Debug.Print "    õ PÿESKAKUJI (stejnÈ ID)"
                shouldOverwrite = False
            Else
                ' Zjistit prioritu
                Dim existingPriority As Long
                existingPriority = GetExistingPromoPriority(TargetWorkbook, existingPromoID)
                Debug.Print "    Priorita: existujÌcÌ " & existingPriority & " vs nov· " & newPriority
                
                If existingPriority < newPriority Then
                    shouldOverwrite = False
                    Debug.Print "    õ NEPÿEPISUJI (vyööÌ priorita)"
                Else
                    shouldOverwrite = True
                    Debug.Print "    õ PÿEPISUJI"
                End If
            End If
        Else
            Debug.Print "    Pr·zdn· õ ZAPISUJI"
        End If
        
        ' P¯idat do kolekce, pokud se m· p¯epsat
        If shouldOverwrite Then
            cellsToOverwrite.Add cell
        End If
    Next cell
    
    Debug.Print "Celkem bunÏk k z·pisu: " & cellsToOverwrite.Count
    
' KROK 2: Teprve TEœ zapisovat koment·¯e a form·tov·nÌ
Application.ScreenUpdating = False
Application.EnableEvents = False

Dim isFirstCell As Boolean
isFirstCell = True

Dim cellToWrite As Range
Dim writeCounter As Long
writeCounter = 0

    Dim i As Long
    For i = cellsToOverwrite.Count To 1 Step -1
        Set cellToWrite = cellsToOverwrite(i)
        writeCounter = writeCounter + 1
        Debug.Print "  --- Z·pis #" & writeCounter & " do buÚky: " & cellToWrite.Address & " ---"
        
        ' Smazat star˝ koment·¯
        On Error Resume Next
        If Not cellToWrite.comment Is Nothing Then
            cellToWrite.comment.Delete
            Application.Wait Now + TimeValue("00:00:01")  ' PÿID¡NO: Ëek·nÌ 1s
            Debug.Print "    Star˝ koment·¯ smaz·n"
        End If
        On Error GoTo ErrorHandler
        
        ' P¯idat nov˝ koment·¯
        Debug.Print "    P¯id·v·m koment·¯: " & CStr(promoID) & " do " & cellToWrite.Address
        
        ' ZKUSIT AKTIVOVAT BU“KU PÿED PÿID¡NÕM KOMENT¡ÿE
        cellToWrite.Activate
        ActiveCell.AddComment CStr(promoID)
        ActiveCell.comment.Visible = False
        
        ' Kontrola z·pisu
        If Not cellToWrite.comment Is Nothing Then
            Debug.Print "    Koment·¯ p¯id·n: " & Left(cellToWrite.comment.Text, 8)
        Else
            Debug.Print "    CHYBA: Koment·¯ nebyl p¯id·n!"
        End If
        
        ' Form·tov·nÌ
        cellToWrite.Interior.Color = RGB(r, g, b)
        cellToWrite.Font.Color = PromoObj.cFont
        Debug.Print "    Barva nastavena"
        
        ' Hodnota jen do prvnÌ buÚky
        If cellToWrite.Column = SelectedRange.Cells(1).Column And cellToWrite.row = SelectedRange.Cells(1).row Then
            cellToWrite.value = PromoObj.promoTyp
            Debug.Print "    Hodnota zaps·na: " & PromoObj.promoTyp
        End If
        
        Debug.Print "    >> BuÚka dokonËena"
    Next i

Application.EnableEvents = True
Application.ScreenUpdating = True
    
    ' KONTROLA PO Z¡PISU
    Debug.Print "=== KONTROLA KOMENT¡ÿŸ PO Z¡PISU ==="
    For Each cell In SelectedRange.Cells
        If Not cell.comment Is Nothing Then
            Debug.Print "  BuÚka " & cell.Address & " m· koment·¯: " & Left(cell.comment.Text, 8)
        Else
            Debug.Print "  BuÚka " & cell.Address & " NEM¡ koment·¯"
        End If
    Next cell
    
    Debug.Print "=== FormatPromoCells END ==="
    Exit Sub
    
ErrorHandler:
    Debug.Print "CHYBA: " & Err.Description
    MsgBox "Chyba v FormatPromoCells: " & Err.Description, vbCritical
End Sub

' ===================================================================
' ZÌsk·nÌ priority podle typAkce
' ===================================================================
Private Function GetPromoPriority(typAkce As String) As Long
    Select Case UCase(Trim(typAkce))
        Case "TV": GetPromoPriority = 1
        Case "L": GetPromoPriority = 2
        Case "ONLINE": GetPromoPriority = 3
        Case "K": GetPromoPriority = 4
        Case "T": GetPromoPriority = 5
        Case "EDLP": GetPromoPriority = 6
        Case "DU": GetPromoPriority = 7
        Case "OTHERS": GetPromoPriority = 8
        Case "LOYALTY": GetPromoPriority = 9
        Case Else
            Debug.Print "VAROV¡NÕ: Nezn·m˝ typAkce '" & typAkce & "', pouûita nejniûöÌ priorita (99)"
            GetPromoPriority = 99  ' NejniûöÌ priorita pro nezn·mÈ typy
    End Select
End Function

' ===================================================================
' ZÌsk·nÌ priority existujÌcÌ promoce z listu Text
' ===================================================================
Private Function GetExistingPromoPriority(TargetWorkbook As Workbook, promoID As String) As Long
    On Error GoTo ErrorHandler
    
    Dim textList As Worksheet
    Set textList = TargetWorkbook.Sheets("Text")
    
    ' NajÌt ¯·dek s tÌmto PromoID
    Dim lastRow As Long
    lastRow = textList.Cells(textList.rows.Count, GetColumnSafe(textList, "tPromoID")).End(xlUp).row
    
    Dim promoIDCol As Long
    promoIDCol = GetColumnSafe(textList, "tPromoID")
    
    Dim typAkceCol As Long
    typAkceCol = GetColumnSafe(textList, "tTypAkce")
    
    Dim i As Long
    For i = 2 To lastRow
        If CStr(textList.Cells(i, promoIDCol).value) = CStr(promoID) Then
            ' Naöli jsme ¯·dek, zjistit typAkce
            Dim existingTypAkce As String
            existingTypAkce = Trim(textList.Cells(i, typAkceCol).value)
            
            ' Vr·tit prioritu
            GetExistingPromoPriority = GetPromoPriority(existingTypAkce)
            Debug.Print "      Nalezen typAkce pro ID " & promoID & ": " & existingTypAkce & " (priorita: " & GetExistingPromoPriority & ")"
            Exit Function
        End If
    Next i
    
    ' Pokud nenalezeno, vr·tit nejniûöÌ prioritu
    Debug.Print "      VAROV¡NÕ: PromoID " & promoID & " nenalezeno v listu Text"
    GetExistingPromoPriority = 99
    Exit Function
    
ErrorHandler:
    Debug.Print "      CHYBA p¯i zjiöùov·nÌ priority pro ID " & promoID & ": " & Err.Description
    GetExistingPromoPriority = 99
End Function
