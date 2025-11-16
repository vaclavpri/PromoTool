Attribute VB_Name = "PriceList_Array"
Option Explicit

Public ProductsDict As Object ' Dictionary for access by column names
Public ProductsCollection As Collection ' Collection of all rows as Dictionary
Public pLastRow As Long
Public PriceList As Worksheet

Public Sub ProductsArray(TargetWorkbook As Workbook)
    Dim dataArray As Variant
    Dim lastColumn As Long
    Dim i As Long, j As Long
    Dim rowDict As Object
    Dim columnName As String
    
    ' Set worksheet
    Set PriceList = TargetWorkbook.Sheets("PriceList")
    
    ' Find data range
    pLastRow = PriceList.Cells(PriceList.rows.Count, 3).End(xlUp).row
    lastColumn = PriceList.Cells(4, PriceList.Columns.Count).End(xlToLeft).Column
    
    ' Load all data at once (from row 4 - column names)
    dataArray = PriceList.Range(PriceList.Cells(4, 1), _
                                 PriceList.Cells(pLastRow, lastColumn)).value
    
    ' Initialize collection
    Set ProductsCollection = New Collection
    
    ' Create Dictionary for each data row
    For i = 2 To UBound(dataArray, 1)
        Set rowDict = CreateObject("Scripting.Dictionary")
        
        ' Add only columns with valid names (skip "0" and empty)
        For j = 1 To UBound(dataArray, 2)
            columnName = Trim(CStr(dataArray(1, j)))
            
            ' Only process valid column names
            If columnName <> "" And columnName <> "0" Then
                ' Handle duplicates by adding suffix
                If Not rowDict.Exists(columnName) Then
                    rowDict.Add columnName, dataArray(i, j)
                Else
                    Dim suffix As Long
                    suffix = 2
                    Do While rowDict.Exists(columnName & "_" & suffix)
                        suffix = suffix + 1
                    Loop
                    rowDict.Add columnName & "_" & suffix, dataArray(i, j)
                End If
            End If
        Next j
        
        ' Add row number for reference
        rowDict.Add "RowNumber", i + 3
        
        ProductsCollection.Add rowDict
    Next i
    
    'Debug.Print "Loaded " & ProductsCollection.Count & " product rows with valid columns only."
    
End Sub

' Helper functions for data access

' Get value by row number and column name
Public Function GetProduct(row As Long, columnName As String) As Variant
    Dim index As Long
    index = row - 4 ' Offset because data starts at row 5
    
    If index >= 1 And index <= ProductsCollection.Count Then
        Dim rowData As Object
        Set rowData = ProductsCollection(index)
        
        If rowData.Exists(columnName) Then
            GetProduct = rowData(columnName)
        Else
            GetProduct = Empty
        End If
    Else
        GetProduct = Empty
    End If
End Function

' Get entire row as Dictionary by row number
Public Function GetProductRow(row As Long) As Object
    Dim index As Long
    index = row - 4
    
    If index >= 1 And index <= ProductsCollection.Count Then
        Set GetProductRow = ProductsCollection(index)
    Else
        Set GetProductRow = Nothing
    End If
End Function

' Get value by collection index and column name
Public Function GetProductByIndex(index As Long, columnName As String) As Variant
    If index >= 1 And index <= ProductsCollection.Count Then
        Dim rowData As Object
        Set rowData = ProductsCollection(index)
        
        If rowData.Exists(columnName) Then
            GetProductByIndex = rowData(columnName)
        Else
            GetProductByIndex = Empty
        End If
    Else
        GetProductByIndex = Empty
    End If
End Function

' Get entire collection
Public Function GetProductsCollection() As Collection
    Set GetProductsCollection = ProductsCollection
End Function

' Get last row number
Public Function GetLastRow() As Long
    GetLastRow = pLastRow
End Function

' Get product count in collection
Public Function GetProductCount() As Long
    GetProductCount = ProductsCollection.Count
End Function

' Find product by value in column
Public Function FindProduct(columnName As String, searchValue As Variant) As Object
    Dim rowData As Object
    
    For Each rowData In ProductsCollection
        If rowData.Exists(columnName) Then
            If rowData(columnName) = searchValue Then
                Set FindProduct = rowData
                Exit Function
            End If
        End If
    Next rowData
    
    Set FindProduct = Nothing
End Function


