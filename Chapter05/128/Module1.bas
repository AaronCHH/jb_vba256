Attribute VB_Name = "Module1"
Option Explicit

Sub 參照作用中的活頁簿()
    Dim xBook As Variant, i As Integer
    MsgBox "目前作用中的活頁簿: " & ActiveWorkbook.Name & Chr(10) & _
           "執行程式中的活頁簿: " & ThisWorkbook.Name
    Workbooks.Open Filename:="課題1.xls"
    MsgBox "目前作用中的活頁簿: " & ActiveWorkbook.Name & Chr(10) & _
           "執行程式中的活頁簿: " & ThisWorkbook.Name
End Sub





