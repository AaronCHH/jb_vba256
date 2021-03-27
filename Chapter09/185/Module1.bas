Attribute VB_Name = "Module1"
Option Explicit

Sub 顯示現在的日期和時間()
    MsgBox "現在的日期和時間" & Chr(10) & _
           "日期: " & Date & Chr(10) & "時間: " & Time, _
           , "確認日期和時間: " & Now
End Sub



