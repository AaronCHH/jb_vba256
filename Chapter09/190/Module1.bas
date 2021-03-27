Attribute VB_Name = "Module1"
Option Explicit

Sub 取得經過日期時間()
    Dim sDate As Date
    sDate = #11/5/1994#
    MsgBox "出生年月日: " & sDate & Chr(10) & _
           "今天的日期: " & Date & Chr(10) & _
           "經過年數: " & DateDiff("yyyy", sDate, Date)
End Sub
