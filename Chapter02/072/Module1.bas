Attribute VB_Name = "Module1"
Option Explicit

Sub �]�w�ؽu()
    With Range("B3:C6")
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlDouble
        .Borders(xlInsideHorizontal).LineStyle = xlDash
    End With
End Sub

Sub �]�w�ؽu2()
    Range("B3:C6").Borders.LineStyle = xlContinuous
End Sub

Sub �]�w�ؽu3()
    Range("B3:C6").Borders.LineStyle = xlNone
End Sub

Sub �]�w�ؽu4()
    Range("B3:C6").BorderAround LineStyle:=xlContinuous
End Sub
