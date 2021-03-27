Attribute VB_Name = "Module1"
Option Explicit

Sub 頁尾設定()
    With ActiveSheet.PageSetup
        .LeftFooter = "&""新明細體""&I 期中考"
        .CenterFooter = "&P/&N"
        .RightFooterPicture.Filename = "C:\ExcelVBA\test.bmp"
        .RightFooter = "&G"
    End With
    ActiveSheet.PrintPreview
End Sub

