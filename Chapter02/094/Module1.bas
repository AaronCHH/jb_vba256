Attribute VB_Name = "Module1"
Option Explicit

Sub ���J�x�s��()
    Dim i As Integer, xRange As Range
    Set xRange = Range("A4").CurrentRegion
    xRange.Columns(3).Insert xlToRight, xlFormatFromLeftOrAbove
    xRange.Columns(3).FormulaR1C1 = "=if(RC[4]>=180, ""�X��"",""���X��"")"
    xRange.Cells(1, 3) = "�O�_�X��"
    Set xRange = Nothing
End Sub

