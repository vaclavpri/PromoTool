Attribute VB_Name = "Clear"
' ===================================================================
' Clear1 - maz�n� promoc�
' ===================================================================
Public Sub Clear1(TargetWorkbook As Workbook)
    ' Z�sk�n� PromoID z ozna�en�ch bun�k
    Dim promoIDsToDelete As Collection
    Set promoIDsToDelete = New Collection
    
    Dim cell As Range
    Dim promoID As String
    
    ' Z�sk�n� v�ech PromoID z koment��� ozna�en�ch bun�k
    On Error Resume Next
    For Each cell In TargetWorkbook.Application.Selection
        If Not cell.comment Is Nothing Then
            ' Z�skat prvn�ch 8 znak� z koment��e (PromoID)
            promoID = Left(cell.comment.Text, 8)
            
            If Len(promoID) = 8 Then  ' Kontrola d�lky 8 znak�
                promoIDsToDelete.Add promoID, promoID
            End If
        End If
    Next cell
    On Error GoTo 0
    
    If promoIDsToDelete.Count = 0 Then
        MsgBox "V ozna�en�ch bu�k�ch nebyla nalezena ��dn� PromoID v koment���ch.", vbInformation
        Exit Sub
    End If
    
    ' POTVRZEN� P�ED SMAZ�N�M
    Dim response As VbMsgBoxResult
    response = MsgBox("Opravdu chcete smazat " & promoIDsToDelete.Count & " promoc�?", _
                      vbYesNo + vbQuestion, "Potvrzen� smaz�n�")
    
    If response = vbNo Then Exit Sub
    
    ' Smaz�n� v�ech promoc� s dan�mi PromoID
    Dim searchValue As String
    Dim i As Long
    
    For i = 1 To promoIDsToDelete.Count
        searchValue = promoIDsToDelete(i)
        Call DeletePromoByID(TargetWorkbook, searchValue)
    Next i
    
    MsgBox "Bylo odstran�no " & promoIDsToDelete.Count & " promoc�.", vbInformation
    
End Sub

Private Sub DeletePromoByID(TargetWorkbook As Workbook, searchValue As String)
    Call UnlockText(TargetWorkbook)
    Call RemoveFilterIfApplied(TargetWorkbook)
    
    Dim textList As Worksheet
    Dim CrmList As Worksheet
    Set textList = TargetWorkbook.Sheets("Text")
    Set CrmList = TargetWorkbook.Sheets("CRM")
    
    ' Na�ten� dat z Text listu do pole
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
    
    ' Zm�na statusu na CRM listu
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
    
    ' Smaz�n� ��dk�
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
    
    ' Vy�i�t�n� V�ECH bun�k se stejn�m PromoID v koment��i
    Call ClearAllCellsWithPromoID(TargetWorkbook, searchValue)
    
End Sub

Private Sub ClearAllCellsWithPromoID(TargetWorkbook As Workbook, searchValue As String)
    On Error Resume Next

    Dim ws As Worksheet
    Dim cell As Range
    Dim promoIDFromComment As String

    ' Nastavit list Promoplan
    Set ws = TargetWorkbook.Sheets("Promoplan")

    ' Projít všechny buňky s komentářem na listu Promoplan
    For Each cell In ws.Cells.SpecialCells(xlCellTypeComments)
        If Not cell.comment Is Nothing Then
            ' Získat prvních 8 znaků z komentáře
            promoIDFromComment = Left(cell.comment.Text, 8)

            ' Pokud se shoduje s hledanou hodnotou
            If promoIDFromComment = searchValue Then
                cell.ClearContents
                cell.Interior.ColorIndex = 0
                cell.ClearComments
            End If
        End If
    Next cell

    On Error GoTo 0
End Sub

' Volat na konci po zpracov�n� v�ech PromoID
Public Sub FinalizeAfterDelete(TargetWorkbook As Workbook)
    Call SortIt(TargetWorkbook)
    Call rColor(TargetWorkbook)
    Call LockText(TargetWorkbook)
End Sub

