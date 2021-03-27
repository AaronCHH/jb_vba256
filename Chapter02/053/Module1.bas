Attribute VB_Name = "Module1"
Option Explicit

Sub 做成表格()
    Range("A1").End(xlDown).Select
    ActiveCell.CurrentRegion.Borders.LineStyle = xlContinuous
    ActiveCell.End(xlDown).CurrentRegion.Borders.LineStyle = xlContinuous
End Sub

