VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Promo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public typAkce As String
Public promoTyp As String
Public WeekRange As String, weekRangeT As String
Public startAkce As Date
Public endAkce As Date
Public StartPurchase As Date
Public EndPurchase As Date
Public weekInterval As String, weekIntervalT As String ', sPurchase As Date, ePurchase As Date
Public sortFrom As Date, sortTo As Date ', sFrom As Date, sTo As Date
Public r As Integer, g As Integer, b As Integer, numR As Integer, numG As Integer, numB As Integer
Public cFont As String

Public Sub PromoSettings(typ As String, pTyp As String, weekInterval As String, weekIntervalT As String, StartWeek As Date, endWeek As Date, sPurchase As Date, ePurchase As Date, sFrom, sTo, numR As Long, numG As Long, numB As Long, FontColor)
    typAkce = typ
    promoTyp = pTyp
    WeekRange = weekInterval
    weekRangeT = weekIntervalT
    startAkce = StartWeek
    endAkce = endWeek
    StartPurchase = sPurchase
    EndPurchase = ePurchase
    sortFrom = sFrom
    sortTo = sTo
    r = numR
    g = numG
    b = numB
    cFont = FontColor
End Sub

