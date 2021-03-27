Attribute VB_Name = "Module1"
Option Explicit

Sub 開始()
    Range("F3").Activate
    MsgBox "出現僅10秒!! 請輸入相同數字!!"
    Range("B3:D5").Interior.Color = xlNone
    Application.OnTime Now + TimeValue("00:00:10"), "結束"
End Sub

Sub 結束()
    Range("B3:D5").Interior.Color = RGB(0, 0, 0)
    MsgBox "已經過10秒!! 請問輸入多少數字???"
End Sub

