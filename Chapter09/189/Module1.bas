Attribute VB_Name = "Module1"
Option Explicit

Sub ���o�P��()
    MsgBox "�z���ͤ�: " & Range("B1") & Chr(10) & _
           "�z���X�ͤ��: " & _
            WeekdayName(Weekday(Range("B1")))
End Sub

Sub ���o�P��2()
    MsgBox "�z���ͤ�: " & Range("B1") & Chr(10) & _
           "�z���X�ͤ��: " & _
            Format(Range("B1"), "dddd")
End Sub


