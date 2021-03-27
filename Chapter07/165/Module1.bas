Attribute VB_Name = "Module1"
Option Explicit

Sub 隱藏捲軸()
    Workbooks.Open Filename:="C:\ExcelVBA\3教科測試.xls"
    With ActiveWindow
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
    End With
End Sub

Sub 隱藏所有視窗的捲軸()
    Application.DisplayScrollBars = Not Application.DisplayScrollBars
End Sub


