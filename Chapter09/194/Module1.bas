Attribute VB_Name = "Module1"
Option Explicit

Sub 組合成時間()
    Dim myTime As Date, myTime2 As Date, myTime3 As Date
    myTime = TimeSerial(Range("A2"), Range("A3"), 0)
    myTime2 = TimeSerial(Range("A2"), Range("A3") - 10, 0)
    myTime3 = TimeSerial(Range("A2") + 1, Range("A3"), 0)
    MsgBox "預約時間: " & myTime & Chr(10) & Chr(10) & _
           "請於 " & myTime2 & " 前到達!! " & Chr(10) & _
           "超過 " & myTime3 & " 則無效!! "
End Sub


