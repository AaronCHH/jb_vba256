Attribute VB_Name = "Module1"
Option Explicit

Sub 設定各種框線()
    Range("B3:C6").BorderAround LineStyle:=xlContinuous, Weight:=xlThick
    Range("B3:C3").Borders(xlEdgeBottom).LineStyle = xlDouble
    Range("B3:C3").Borders(xlEdgeBottom).Weight = xlMedium
End Sub

