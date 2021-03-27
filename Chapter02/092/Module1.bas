Attribute VB_Name = "Module1"
Option Explicit

Sub 自動調整列高欄寬()
    Rows(1).AutoFit
    Range("A3").CurrentRegion.Rows.AutoFit
    Range("A3").CurrentRegion.Columns.AutoFit
End Sub

