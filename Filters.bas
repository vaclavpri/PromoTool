Attribute VB_Name = "Filters"
Public Sub RemoveFilterIfApplied(TargetWorkbook As Workbook)
    Dim ws As Worksheet
    Set ws = TargetWorkbook.Sheets("Text")
    
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
End Sub

Public Sub ApplyFilterToRow2(TargetWorkbook As Workbook)
    Dim ws As Worksheet
    Set ws = TargetWorkbook.Sheets("Text")
    
    On Error Resume Next
    
    ' Vypnout filtr (pokud je)
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    ' Najít poslední sloupec na øádku 2
    Dim lastCol As Long
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    
    ' Zapnout AutoFilter na øádku 2
    If lastCol > 0 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(2, lastCol)).AutoFilter
    End If
    
    On Error GoTo 0
    
End Sub
