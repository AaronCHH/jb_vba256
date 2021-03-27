Attribute VB_Name = "Module1"
Option Explicit

Sub 插入儲存格()
    Dim i As Integer, xRange As Range
    Set xRange = Range("A4").CurrentRegion
    xRange.Columns(3).Insert xlToRight, xlFormatFromLeftOrAbove
    xRange.Columns(3).FormulaR1C1 = "=if(RC[4]>=180, ""合格"",""不合格"")"
    xRange.Cells(1, 3) = "是否合格"
    Set xRange = Nothing
End Sub

