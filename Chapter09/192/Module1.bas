Attribute VB_Name = "Module1"
Option Explicit

Sub 日期時間轉換()
    Dim strDate As String, strTime As String
    strDate = "民國98年9月1日"
    strTime = "下午3時25分"
    MsgBox "日期: " & DateValue(strDate) & Chr(10) & _
           "時間: " & TimeValue(strTime) & Chr(10) & _
           "1週後: " & DateValue(strDate) + 7 & Chr(10) & _
           "2小時後: " & TimeValue(strTime) + 2 / 24
End Sub
