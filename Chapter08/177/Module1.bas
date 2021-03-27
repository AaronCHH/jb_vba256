Attribute VB_Name = "Module1"
Option Explicit

Sub 列印方向和紙張尺寸()
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperB4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ActiveSheet.PrintPreview
End Sub





