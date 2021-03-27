Attribute VB_Name = "Module1"
Option Explicit

Sub 調整列高欄寬()
    Range("A1").RowHeight = 30
    Range("A3").CurrentRegion.RowHeight = 20
    Range("A3").CurrentRegion.ColumnWidth = 12
End Sub

Sub 調整列高欄寬2()
    Rows(1).RowHeight = 30
    Rows("3:8").RowHeight = 20
    Columns("B:E").ColumnWidth = 12
End Sub

