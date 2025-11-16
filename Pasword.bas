Attribute VB_Name = "Pasword"
' ===== CENTRÁLNÍ HESLO =====
Public Function GetPassword() As String
    GetPassword = "HeslO"
End Function

Public Sub Pasw(Optional TargetWorkbook As Workbook)
    heslo = GetPassword()
End Sub

'UnlockPriceList
Public Sub UnlockPriceList(TargetWorkbook As Workbook)
    TargetWorkbook.Sheets("PriceList").Unprotect Password:=GetPassword()
End Sub

' UnlockText
Public Sub UnlockText(TargetWorkbook As Workbook)
    TargetWorkbook.Sheets("Text").Unprotect Password:=GetPassword()
End Sub

' UnlockSAP
Public Sub UnlockSAP(TargetWorkbook As Workbook)
    TargetWorkbook.Sheets("SAP").Unprotect Password:=GetPassword()
End Sub

'LockPriceList
Public Sub LockPriceList(TargetWorkbook As Workbook)
    TargetWorkbook.Sheets("PriceList").Protect Password:=GetPassword()
End Sub

' LockText
Public Sub LockText(TargetWorkbook As Workbook)
    TargetWorkbook.Sheets("Text").Protect Password:=GetPassword()
End Sub

' LockSAP
Public Sub LockSAP(TargetWorkbook As Workbook)
    TargetWorkbook.Sheets("SAP").Protect Password:=GetPassword()
End Sub

