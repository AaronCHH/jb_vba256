Attribute VB_Name = "Module2"
Option Explicit

Sub 結合儲存格範圍()
    Dim Range1 As Range, Range2 As Range
    Dim Range3 As Range, allRange As Range
    
    Set Range1 = Range("A1").CurrentRegion
    Set Range2 = Range("A8").CurrentRegion
    Set Range3 = Range("A13").CurrentRegion
    Set allRange = Union(Range1, Range2, Range3)
    
    allRange.SpecialCells(xlCellTypeFormulas). _
                Interior.Color = RGB(255, 255, 0)
    Set Range1 = Nothing: Set Range2 = Nothing
    Set Range3 = Nothing: Set Range1 = Nothing
End Sub


