Attribute VB_Name = "Module1"
Option Explicit

Sub ����ɶ��ഫ()
    Dim strDate As String, strTime As String
    strDate = "����98�~9��1��"
    strTime = "�U��3��25��"
    MsgBox "���: " & DateValue(strDate) & Chr(10) & _
           "�ɶ�: " & TimeValue(strTime) & Chr(10) & _
           "1�g��: " & DateValue(strDate) + 7 & Chr(10) & _
           "2�p�ɫ�: " & TimeValue(strTime) + 2 / 24
End Sub
