Attribute VB_Name = "Module3"
Option Explicit

Sub �۹��x�s��()
    Range("A3").Offset(7, 0).Select
    Selection.Value = "�X�p"
    Selection.Offset(0, 1).Formula = "=SUM(B4:B9)"
    Selection.Offset(0, 2).Formula = "=B10/E1"
    Range(Selection, Selection.Offset(0, 2)).Interior.ColorIndex = 36
End Sub

