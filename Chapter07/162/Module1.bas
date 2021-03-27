Attribute VB_Name = "Module1"
Option Explicit

Sub 視窗排列()
    ActiveWindow.WindowState = xlMinimized
    Workbooks.Open Filename:="C:\ExcelVBA\3教科測試.xls"
    ActiveWorkbook.NewWindow
    Windows.Arrange xlArrangeStyleVertical
End Sub



