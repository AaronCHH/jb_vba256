Attribute VB_Name = "Module1"
Option Explicit

Sub 取得時間單位()
    Dim dp As Integer
    dp = DatePart("y", Date)
    MsgBox "今天的日期: " & Date & Chr(10) & _
           "1月1日到現在的經過日數: " & dp & Chr(10) & _
           "1年當中經過了: " & Format(dp / 365, "0.0%")
End Sub
