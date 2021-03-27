Attribute VB_Name = "Module1"
Option Explicit

Sub 作用活頁簿()
    Dim xBook As Variant, i As Integer
    xBook = Array("課題1.xls", "課題2.xls", "課題3.xls")
    For i = 0 To 2
        Workbooks.Open Filename:=xBook(i)
        Workbooks(1).Worksheets(1).Cells(i + 2, "B").Value = xBook(i)
    Next
    Workbooks(1).Activate
End Sub


