Attribute VB_Name = "Module1"
Option Explicit

Sub 字碼()
    Dim strText As String
    strText = InputBox("今天的日期: " & Date & Chr(10) & _
             "請輸入行程表!!!")
    MsgBox "今日的行程: " & Chr(9) & strText
End Sub

Sub 字碼2()
    MsgBox Asc("Excel")
End Sub

