Attribute VB_Name = "Module1"
Option Explicit

Sub �C�L��V�M�ȱi�ؤo()
    With ActiveSheet.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperB4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ActiveSheet.PrintPreview
End Sub





