Attribute VB_Name = "Module1"
Option Explicit

Sub 調整適和表格的視窗尺寸()
    With ActiveWindow
        .WindowState = xlNormal
        .Top = 0
        .Left = 120
        .Width = Range("送貨單").Width + 55
        .Height = Range("送貨單").Height + 60
    End With
End Sub
