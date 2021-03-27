Attribute VB_Name = "Module1"
Option Explicit

Sub ¿ï¾ÜÀx¦s®æ()
    Range("A3:B11").Select
    Selection.Borders.LineStyle = xlContinuous
    Range("B11").Activate
    ActiveCell.Formula = "=SUM(B4:B10)"
End Sub

