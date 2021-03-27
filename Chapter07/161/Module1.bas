Attribute VB_Name = "Module1"
Option Explicit

Sub 參照視窗()
    Workbooks.Open Filename:="C:\ExcelVBA\國語測試.xls"
    Workbooks.Open Filename:="C:\ExcelVBA\英語測試.xls"
    Windows.Arrange xlArrangeStyleCascade
    MsgBox "第1個: " & Windows(1).Caption & Chr(10) & _
           "第2個: " & Windows(2).Caption
End Sub


