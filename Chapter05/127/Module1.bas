Attribute VB_Name = "Module1"
Option Explicit

Sub �ѷӬ���ï()
    Dim xBook As Variant, i As Integer
    xBook = Array("���D1.xls", "���D2.xls", "���D3.xls")
    For i = 0 To 2
        Workbooks.Open Filename:=xBook(i)
    Next
    MsgBox Workbooks(1).Name & ":" & Workbooks(2).Name & ":" & _
        Workbooks(3).Name & ":" & Workbooks(4).Name
End Sub






