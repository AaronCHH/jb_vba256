Attribute VB_Name = "Module1"
Option Explicit

Sub �@�ά���ï()
    Dim xBook As Variant, i As Integer
    xBook = Array("���D1.xls", "���D2.xls", "���D3.xls")
    For i = 0 To 2
        Workbooks.Open Filename:=xBook(i)
        Workbooks(1).Worksheets(1).Cells(i + 2, "B").Value = xBook(i)
    Next
    Workbooks(1).Activate
End Sub


