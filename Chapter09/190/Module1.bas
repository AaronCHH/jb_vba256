Attribute VB_Name = "Module1"
Option Explicit

Sub ���o�g�L����ɶ�()
    Dim sDate As Date
    sDate = #11/5/1994#
    MsgBox "�X�ͦ~���: " & sDate & Chr(10) & _
           "���Ѫ����: " & Date & Chr(10) & _
           "�g�L�~��: " & DateDiff("yyyy", sDate, Date)
End Sub
