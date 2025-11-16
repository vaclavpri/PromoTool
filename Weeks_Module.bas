Attribute VB_Name = "Weeks_Module"
Option Explicit

Public Weeks() As Variant
Public Promo As Promo
Public startAkce As Date
Public endAkce As Date
Public targetWeekNumber As Long
Public i As Long
Public wLastRow As Long, wsWeeks As Worksheet, StartWeek As Date, endWeek As Date, sPurchase As Date

Public Sub WeeksArray(TargetWorkbook As Workbook, SelectedRange As Range)
    Dim wsWeeks As Worksheet
    Set wsWeeks = TargetWorkbook.Sheets("Weeks")
    
    wLastRow = wsWeeks.Cells(wsWeeks.rows.Count, 1).End(xlUp).row
    
    ' Definuje pole Weeks
    ReDim Weeks(1 To wLastRow - 1, 1 To 3)
    
    ' Naète data z listu "Weeks" do pole Weeks
    Dim i As Long
    For i = 2 To wLastRow
        Weeks(i - 1, 1) = wsWeeks.Cells(i, 1).value ' WeekNumber
        Weeks(i - 1, 2) = wsWeeks.Cells(i, 2).value ' StartWeek
        Weeks(i - 1, 3) = wsWeeks.Cells(i, 3).value ' EndWeek
    Next i
    
    Set Promo = New Promo
    
    Dim o As Variant
    
    With SelectedRange
        o = .Cells(.Count).Column
    End With
    
    Dim endWeekNumber As Integer
    Dim ws As Worksheet
    Set ws = SelectedRange.Worksheet
    
    targetWeekNumber = ws.Cells(5, SelectedRange(1).Column).value
    endWeekNumber = ws.Cells(5, o).value
    
    ' Hledání odpovídajícího záznamu v poli Weeks
    For i = LBound(Weeks, 1) To UBound(Weeks, 1)
        If Weeks(i, 1) = targetWeekNumber Then
            StartWeek = Weeks(i, 2)
        End If
        If Weeks(i, 1) = endWeekNumber Then
            endWeek = CDate(Weeks(i, 3))
            Exit For
        End If
    Next i
End Sub

Public Function GetStartWeek() As Date
    GetStartWeek = StartWeek
End Function

Public Function GetEndWeek() As Date
    GetEndWeek = endWeek
End Function

' Pomocná funkce pro nalezení øádku s komentáøem "WeekRow"
Public Function FindWeekRow(ws As Worksheet) As Long
    Dim cell As Range
    Dim lastRow As Long
    
    On Error Resume Next
    
    ' Najít poslední øádek ve sloupci A
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).row
    
    ' Prohledat sloupec A
    For Each cell In ws.Range("A1:A" & lastRow)
        ' Zkontrolovat, zda buòka má komentáø
        If Not cell.comment Is Nothing Then
            ' Zkontrolovat, zda komentáø obsahuje "WeekRow"
            If InStr(1, cell.comment.Text, "WeekRow", vbTextCompare) > 0 Then
                FindWeekRow = cell.row
                Exit Function
            End If
        End If
    Next cell
    
    ' Pokud nebyl nalezen, vrátit 0
    FindWeekRow = 0
    
    On Error GoTo 0
End Function
