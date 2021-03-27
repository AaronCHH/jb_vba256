Attribute VB_Name = "Module1"
Option Explicit

Sub 確認縮放比率設定()
    ActiveSheet.PageSetup.Zoom = 200
    ActiveSheet.PrintPreview
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ActiveSheet.PrintPreview
End Sub


