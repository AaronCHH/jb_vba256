Attribute VB_Name = "Module1"
Option Explicit

Sub �^�����w�ɶ�()
    Dim s1 As Single, s2 As Double, qText As String, aText As String
    Dim ansText As String, msgText As String
    
    qText = "�^�ꪺ�����W�٬O?? (�b�έ^��r��)"
    aText = "UNITED KINGDOM"
    s1 = Timer
    ansText = InputBox(qText, "���D")
    s2 = Timer - s1
    msgText = "���T����: " & aText & Chr(10) & "�^���ɶ�: " & s2 & "��"
    If StrComp(aText, ansText, 1) = 0 Then
       MsgBox "���T!!" & Chr(10) & msgText
    Else
       MsgBox "���~!!" & Chr(10) & msgText
    End If
End Sub


