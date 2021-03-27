Attribute VB_Name = "Module1"
Option Explicit

Sub 取得星期()
    MsgBox "您的生日: " & Range("B1") & Chr(10) & _
           "您的出生日期: " & _
            WeekdayName(Weekday(Range("B1")))
End Sub

Sub 取得星期2()
    MsgBox "您的生日: " & Range("B1") & Chr(10) & _
           "您的出生日期: " & _
            Format(Range("B1"), "dddd")
End Sub


