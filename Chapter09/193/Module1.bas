Attribute VB_Name = "Module1"
Option Explicit

Sub 做成日期()
    Dim myDate As Date, myDate2 As Date
    myDate = DateSerial(Range("A2"), Range("A3"), Range("A4"))
    myDate2 = DateSerial(Year(Date), Range("A3"), Range("A4"))
    MsgBox "您的誕生日: " & myDate & Chr(10) & _
           "今年的誕生日是 " & myDate2 & Format(myDate2, "(aaaa)") _
           & " 沒錯吧!!"
End Sub


