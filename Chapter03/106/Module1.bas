Attribute VB_Name = "Module1"
Option Explicit

Sub 合計()
    With ActiveSheet.ListObjects("Table1")
        .ShowTotals = True
        .ListColumns("單價").TotalsCalculation = xlTotalsCalculationCount
        .ListColumns("數量").TotalsCalculation = xlTotalsCalculationAverage
    End With
End Sub



