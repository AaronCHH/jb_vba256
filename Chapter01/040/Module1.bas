Attribute VB_Name = "Module1"
Option Explicit

Sub �}�l()
    Range("F3").Activate
    MsgBox "�X�{��10��!! �п�J�ۦP�Ʀr!!"
    Range("B3:D5").Interior.Color = xlNone
    Application.OnTime Now + TimeValue("00:00:10"), "����"
End Sub

Sub ����()
    Range("B3:D5").Interior.Color = RGB(0, 0, 0)
    MsgBox "�w�g�L10��!! �аݿ�J�h�ּƦr???"
End Sub

