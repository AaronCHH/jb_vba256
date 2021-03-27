Attribute VB_Name = "Module1"
Option Explicit

Sub 固定視窗尺寸()
    ActiveWindow.WindowState = xlMinimized
    Workbooks.Open Filename:="C:\ExcelVBA\3教科測試.xls"
    
    With ActiveWindow
        .WindowState = xlNormal
        .Top = 1
        .Left = 1
        .Height = 250
        .Width = 400
    End With
    ActiveWindow.EnableResize = False
End Sub



