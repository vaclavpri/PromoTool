Attribute VB_Name = "RowsColor"
Public Sub rColor(TargetWorkbook As Workbook)
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = TargetWorkbook.Worksheets("Text")
    
    If ws Is Nothing Then
        Debug.Print "List 'Text' nebyl nalezen"
        Exit Sub
    End If
    
    ' Najít sloupec tCom
    Dim tComColumn As Long
    tComColumn = ws.Range("tCom").Column
    
    ' Najít poslední øádek
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, tComColumn).End(xlUp).row
    
    If lastRow < 2 Then Exit Sub ' Žádná data
    
    ' Definovat rozsah dat
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(1, tComColumn), ws.Cells(lastRow + 1, tComColumn))
    
    ' Naèíst data do pole
    Dim DataValues() As Variant
    DataValues = dataRange.value
    
    ' Promìnné pro sledování blokù
    Dim iStart As Long
    iStart = 1
    Dim BlockValue As Variant
    Dim IsEven As Boolean
    IsEven = False
    Dim EvenBlocks As Range
    Dim OddBlocks As Range
    Dim CurrentBlock As Range
    Dim iRow As Long
    
    ' Projít všechna data a najít bloky
    For iRow = LBound(DataValues) + 1 To UBound(DataValues)
        If BlockValue <> DataValues(iRow, 1) Then
            If iRow - iStart > 0 Then
                Set CurrentBlock = dataRange.Cells(iStart, 1).Resize(RowSize:=iRow - iStart)
                
                If IsEven Then
                    If EvenBlocks Is Nothing Then
                        Set EvenBlocks = CurrentBlock
                    Else
                        Set EvenBlocks = Union(EvenBlocks, CurrentBlock)
                    End If
                Else
                    If OddBlocks Is Nothing Then
                        Set OddBlocks = CurrentBlock
                    Else
                        Set OddBlocks = Union(OddBlocks, CurrentBlock)
                    End If
                End If
                
                IsEven = Not IsEven
            End If
            
            iStart = iRow
            BlockValue = DataValues(iRow, 1)
        End If
    Next iRow
    
    ' Obarvit všechny sudé a liché bloky støídavì
    If Not EvenBlocks Is Nothing Then
        EvenBlocks.EntireRow.Interior.ColorIndex = 0
    End If
    
    If Not OddBlocks Is Nothing Then
        OddBlocks.EntireRow.Interior.Color = RGB(255, 255, 0)
    End If
    
    ' Nastavit barvu písma na èernou pro všechny øádky
    Dim i As Long
    For i = 1 To dataRange.rows.Count
        ws.rows(i).Font.Color = RGB(0, 0, 0)
    Next i
    
    On Error GoTo 0
End Sub
