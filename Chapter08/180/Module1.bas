Attribute VB_Name = "Module1"
Option Explicit

Sub �����]�w()
    With ActiveSheet.PageSetup
        .LeftFooter = "&""�s������""&I ������"
        .CenterFooter = "&P/&N"
        .RightFooterPicture.Filename = "C:\ExcelVBA\test.bmp"
        .RightFooter = "&G"
    End With
    ActiveSheet.PrintPreview
End Sub

