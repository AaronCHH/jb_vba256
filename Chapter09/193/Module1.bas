Attribute VB_Name = "Module1"
Option Explicit

Sub �������()
    Dim myDate As Date, myDate2 As Date
    myDate = DateSerial(Range("A2"), Range("A3"), Range("A4"))
    myDate2 = DateSerial(Year(Date), Range("A3"), Range("A4"))
    MsgBox "�z���ϥͤ�: " & myDate & Chr(10) & _
           "���~���ϥͤ�O " & myDate2 & Format(myDate2, "(aaaa)") _
           & " �S���a!!"
End Sub


