Attribute VB_Name = "Module1"
Option Explicit

Sub �X�p()
    With ActiveSheet.ListObjects("Table1")
        .ShowTotals = True
        .ListColumns("���").TotalsCalculation = xlTotalsCalculationCount
        .ListColumns("�ƶq").TotalsCalculation = xlTotalsCalculationAverage
    End With
End Sub



