Attribute VB_Name = "Module1"
Option Explicit

Sub �������()
    Range("A1").End(xlDown).Select
    ActiveCell.CurrentRegion.Borders.LineStyle = xlContinuous
    ActiveCell.End(xlDown).CurrentRegion.Borders.LineStyle = xlContinuous
End Sub

