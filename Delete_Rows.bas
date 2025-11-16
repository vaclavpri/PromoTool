Attribute VB_Name = "Delete_Rows"
Public Sub DeleteSelectedRowsText(TargetWorkbook As Workbook, SelectedRange As Range)
    On Error GoTo ErrorHandler
    
    ' Kontrola, zda je nìco oznaèeno
    If SelectedRange Is Nothing Then
        MsgBox "Není oznaèen žádný rozsah.", vbExclamation
        Exit Sub
    End If
    
    ' Potvrzení pøed smazáním
    Dim odpoved As VbMsgBoxResult
    odpoved = MsgBox("Opravdu chceš smazat oznaèené øádky?", vbYesNo + vbQuestion, "Potvrzení")
    
    If odpoved = vbNo Then
        Exit Sub
    End If
    
    ' Odemkne Text list
    Call UnlockText(TargetWorkbook)
    
    ' Smazání øádkù
    Dim i As Long
    For i = SelectedRange.Areas.Count To 1 Step -1
        SelectedRange.Areas(i).EntireRow.Delete
    Next i
    
    ' Zamkne Text list
    Call LockText(TargetWorkbook)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Chyba pøi mazání øádkù: " & Err.Description
    ' Zajistí zamknutí i pøi chybì
    On Error Resume Next
    Call LockText(TargetWorkbook)
End Sub
