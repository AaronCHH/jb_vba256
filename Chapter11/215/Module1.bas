Attribute VB_Name = "Module1"
Option Explicit

Sub �u�@����()
    Dim i As Integer, avg As Double, cr As Range
    Set cr = Range("A3").CurrentRegion
    avg = WorksheetFunction.Average(Range("F4:F13"))
    For i = 2 To cr.Rows.Count
        cr.Cells(i, 6).Font.ColorIndex = _
            IIf(cr.Cells(i, 6).Value >= avg, 3, xlAutomatic)
    Next
    MsgBox "3�Ь쪺��������: " & avg & " ��"
End Sub





