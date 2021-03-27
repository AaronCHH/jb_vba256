Attribute VB_Name = "Module1"
Option Explicit

Sub 恢復標準列高欄寬()
    Range("A3").CurrentRegion.UseStandardHeight = True
    Range("A3").CurrentRegion.UseStandardWidth = True
End Sub

Sub 整個工作表恢復標準列高和欄寬()
    Rows.RowHeight = ActiveSheet.StandardHeight
    Columns.ColumnWidth = ActiveSheet.StandardWidth
End Sub

