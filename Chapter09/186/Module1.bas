Attribute VB_Name = "Module1"
Option Explicit

Sub 個別取得年月日()
    MsgBox "西曆: " & Year(Range("B1")) & Chr(10) & _
           "  月: " & Month("民國83年11月5日") & Chr(10) & _
           "  日: " & Day(#11/5/1994#)
End Sub

Sub 指定年月日測試()
    MsgBox Year("民國98-09-01") & Chr(10) & _
            Month("2009/09/01") & Chr(10) & _
            Day("２００９年９月１日")
End Sub




