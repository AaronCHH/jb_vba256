Attribute VB_Name = "Module1"
Option Explicit

Sub ��ɳ]�w()
    With ActiveSheet.PageSetup
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
        .LeftMargin = Application.CentimetersToPoints(1.5)
        .RightMargin = Application.CentimetersToPoints(1.5)
        .CenterHorizontally = True
    End With
    ActiveSheet.PrintPreview
End Sub

