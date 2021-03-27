Attribute VB_Name = "Module1"
Option Explicit

Sub 設定框線()
    With Range("B3:C6")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlDouble
        .Borders(xlInsideHorizontal).LineStyle = xlDash
    End With
End Sub

Sub 設定框線2()
    Range("B3:C6").Borders.LineStyle = xlContinuous
End Sub

Sub 設定框線3()
    Range("B3:C6").Borders.LineStyle = xlNone
End Sub

Sub 設定框線4()
    Range("B3:C6").BorderAround LineStyle:=xlContinuous
End Sub
