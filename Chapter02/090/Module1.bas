Attribute VB_Name = "Module1"
Option Explicit

Sub �վ�C����e()
    Range("A1").RowHeight = 30
    Range("A3").CurrentRegion.RowHeight = 20
    Range("A3").CurrentRegion.ColumnWidth = 12
End Sub

Sub �վ�C����e2()
    Rows(1).RowHeight = 30
    Rows("3:8").RowHeight = 20
    Columns("B:E").ColumnWidth = 12
End Sub

