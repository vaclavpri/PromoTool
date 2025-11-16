Attribute VB_Name = "Family_Module"
' ==================================================
' FAMILYLIST - Kolekce a vyhledávání
' ==================================================

Public FamilyCollection As Collection
Public fLastRow As Long
Public FamilyList As Worksheet

' Dictionary pro rychlé vyhledávání podle EAN
Public fmEANLookup As Object

Public Sub FamilyArray(TargetWorkbook As Workbook)
    Dim dataArray As Variant
    Dim lastColumn As Long
    Dim i As Long, j As Long
    Dim rowDict As Object
    Dim columnName As String
    Dim eanValue As Variant
    
    ' Set worksheet
    Set FamilyList = TargetWorkbook.Sheets("FamilyList")
    
    ' Find data range - názvy sloupcù jsou na øádku 1
    fLastRow = FamilyList.Cells(FamilyList.rows.Count, 1).End(xlUp).row
    lastColumn = FamilyList.Cells(1, FamilyList.Columns.Count).End(xlToLeft).Column
    
    ' Load all data at once (from row 1 - column names)
    dataArray = FamilyList.Range(FamilyList.Cells(1, 1), _
                                  FamilyList.Cells(fLastRow, lastColumn)).value
    
    ' Initialize collections and lookup dictionary
    Set FamilyCollection = New Collection
    Set fmEANLookup = CreateObject("Scripting.Dictionary")
    
    ' Create Dictionary for each data row (od 2. øádku pole = 2. øádek listu)
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
        rowDict.Add "RowNumber", i
        
        FamilyCollection.Add rowDict
        
        ' Pøidat do lookup dictionary pro rychlé vyhledávání podle EAN
        If rowDict.Exists("EAN") Then
            eanValue = Trim(CStr(rowDict("EAN")))
            If eanValue <> "" And Not fmEANLookup.Exists(eanValue) Then
                fmEANLookup.Add eanValue, rowDict
            End If
        ElseIf rowDict.Exists("ean") Then
            eanValue = Trim(CStr(rowDict("ean")))
            If eanValue <> "" And Not fmEANLookup.Exists(eanValue) Then
                fmEANLookup.Add eanValue, rowDict
            End If
        End If
    Next i
    
    Debug.Print "Loaded " & FamilyCollection.Count & " family rows."
    Debug.Print "fmEANLookup count: " & fmEANLookup.Count
    
End Sub

' Získání celé kolekce FamilyList
Public Function GetFamilyCollection() As Collection
    Set GetFamilyCollection = FamilyCollection
End Function

' Vyhledání v FamilyList podle EAN (rychlé vyhledávání pomocí Dictionary)
Public Function FindInFamilyList(eanValue As String) As Object
    Dim cleanEAN As String
    cleanEAN = Trim(CStr(eanValue))
    
    If fmEANLookup Is Nothing Then
        Set FindInFamilyList = Nothing
        Exit Function
    End If
    
    If fmEANLookup.Exists(cleanEAN) Then
        Set FindInFamilyList = fmEANLookup(cleanEAN)
    Else
        Set FindInFamilyList = Nothing
    End If
End Function

Public Function GetFamilyByEAN(eanValue As Variant) As Variant
    Dim familyRow As Object
    Set familyRow = FindInFamilyList(CStr(eanValue))
    
    If Not familyRow Is Nothing Then
        If familyRow.Exists("Family") Then              ' S velkým F ve FamilyList
            GetFamilyByEAN = familyRow("Family")
        Else
            GetFamilyByEAN = ""
        End If
    Else
        GetFamilyByEAN = ""
    End If
End Function

Public Function GetCustomerIDByEAN(eanValue As Variant) As Variant
    Dim familyRow As Object
    Set familyRow = FindInFamilyList(CStr(eanValue))
    
    If Not familyRow Is Nothing Then
        If familyRow.Exists("CustomerID") Then          ' Bez mezery (po pøejmenování)
            GetCustomerIDByEAN = familyRow("CustomerID")
        Else
            GetCustomerIDByEAN = ""
        End If
    Else
        GetCustomerIDByEAN = ""
    End If
End Function

Public Function GetBrandByEAN(eanValue As Variant) As Variant
    Dim familyRow As Object
    Set familyRow = FindInFamilyList(CStr(eanValue))
    
    If Not familyRow Is Nothing Then
        If familyRow.Exists("Brand") Then
            GetBrandByEAN = familyRow("Brand")
        Else
            GetBrandByEAN = ""
        End If
    Else
        GetBrandByEAN = ""
    End If
End Function

' Získání poètu øádkù
Public Function GetFamilyCount() As Long
    If FamilyCollection Is Nothing Then
        GetFamilyCount = 0
    Else
        GetFamilyCount = FamilyCollection.Count
    End If
End Function

