Attribute VB_Name = "Sort"
Public Sub SortIt(TargetWorkbook As Workbook)
    On Error Resume Next
    
    Dim textList As Worksheet
    Set textList = TargetWorkbook.Sheets("Text")
    
    If textList Is Nothing Then
        Debug.Print "List 'Text' nebyl nalezen"
        Exit Sub
    End If
    
    ' Kontrola, zda je aktivní AutoFilter - pokud ne, aplikuj ho
    If Not textList.AutoFilterMode Then
        Call ApplyFilterToRow2(TargetWorkbook)
    End If
    
    ' Pokraèovat pouze pokud je nyní AutoFilter aktivní
    If textList.AutoFilterMode Then
        With textList.AutoFilter.Sort
            .SortFields.Clear
            
            ' Øazení podle tSortFrom (AN2)
            .SortFields.Add Key:=textList.Range("tSortFrom").Cells(1), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal
            
            ' Øazení podle tSortTo (AO2)
            .SortFields.Add Key:=textList.Range("tSortTo").Cells(1), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal
            
            ' Øazení podle tøetího sloupce (AK)
            .SortFields.Add Key:=textList.Range("tCom").Cells(1), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal
            
            .SortFields.Add Key:=textList.Range("tFamily").Cells(1), _
                           SortOn:=xlSortOnValues, _
                           Order:=xlAscending, _
                           DataOption:=xlSortNormal
            
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Else
        Debug.Print "AutoFilter se nepodaøilo aktivovat"
    End If
    
    On Error GoTo 0
End Sub

