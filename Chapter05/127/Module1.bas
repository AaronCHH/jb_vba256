Attribute VB_Name = "Module1"
Option Explicit

Sub 參照活頁簿()
    Dim xBook As Variant, i As Integer
    xBook = Array("課題1.xls", "課題2.xls", "課題3.xls")
    For i = 0 To 2
        Workbooks.Open Filename:=xBook(i)
    Next
    MsgBox Workbooks(1).Name & ":" & Workbooks(2).Name & ":" & _
        Workbooks(3).Name & ":" & Workbooks(4).Name
End Sub






