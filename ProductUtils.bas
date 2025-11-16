Attribute VB_Name = "ProductUtils"
Option Explicit

' ============================================================================
' Modul: ProductUtils
' Popis: Spole�n� funkce pro pr�ci s produkty
' ============================================================================

' ============================================================================
' Funkce: GetProductName
' Popis: Vytvo�� n�zev produktu podle countryCode
'        - Pro SVK: pouze material_name
'        - Pro ostatn� zem� (CZE): material_name + volume_l
'
' Parametry:
'   rowData      - Dictionary s daty produktu (mus� obsahovat "material_name" a "volume_l")
'   countryCode  - K�d zem� (nap�. "SVK", "CZE")
'
' N�vratov� hodnota: �et�zec s n�zvem produktu
' ============================================================================
Public Function GetProductName(rowData As Object, countryCode As String) As String
    If UCase(Trim(countryCode)) = "SVK" Then
        GetProductName = rowData("material_name")
    Else
        GetProductName = rowData("material_name") & " " & rowData("volume_l")
    End If
End Function

' ============================================================================
' Funkce: GetProductKey
' Popis: Vytvo�� kl�� produktu pro SAP export (Family + ProductName)
'        - Pro SVK: Family + material_name
'        - Pro ostatn� zem� (CZE): Family + material_name + volume_l
'
' Parametry:
'   rowData      - Dictionary s daty produktu (mus� obsahovat "Family", "material_name" a "volume_l")
'   countryCode  - K�d zem� (nap�. "SVK", "CZE")
'
' N�vratov� hodnota: �et�zec s kl��em produktu
' ============================================================================
Public Function GetProductKey(rowData As Object, countryCode As String) As String
    Dim productName As String
    productName = GetProductName(rowData, countryCode)

    If rowData.Exists("Family") Then
        GetProductKey = rowData("Family") & productName
    Else
        GetProductKey = productName
    End If
End Function
