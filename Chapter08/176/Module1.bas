Attribute VB_Name = "Module1"
Option Explicit

Sub �T�{�Y���v�]�w()
    ActiveSheet.PageSetup.Zoom = 200
    ActiveSheet.PrintPreview
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    ActiveSheet.PrintPreview
End Sub


