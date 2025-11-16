Attribute VB_Name = "ProductUtils"
Option Explicit

' ============================================================================
' Modul: ProductUtils
' Popis: Společné funkce pro práci s produkty
' ============================================================================

' ============================================================================
' Funkce: GetProductName
' Popis: Vytvoří název produktu podle countryCode
'        - Pro SVK: pouze material_name
'        - Pro ostatní země (CZE): material_name + volume_l
'
' Parametry:
'   rowData      - Dictionary s daty produktu (musí obsahovat "material_name" a "volume_l")
'   countryCode  - Kód země (např. "SVK", "CZE")
'
' Návratová hodnota: Řetězec s názvem produktu
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
' Popis: Vytvoří klíč produktu pro SAP export (Family + ProductName)
'        - Pro SVK: Family + material_name
'        - Pro ostatní země (CZE): Family + material_name + volume_l
'
' Parametry:
'   rowData      - Dictionary s daty produktu (musí obsahovat "Family", "material_name" a "volume_l")
'   countryCode  - Kód země (např. "SVK", "CZE")
'
' Návratová hodnota: Řetězec s klíčem produktu
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
