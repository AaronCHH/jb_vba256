Attribute VB_Name = "Module1"
Option Explicit

Sub �ƭȽT�{()
    Dim tokuten As String
    tokuten = InputBox("�п�J�o��!!")
    If IsNumeric(tokuten) Then
       MsgBox tokuten & "��!!"
    Else
       MsgBox "�п�J�ƭ�!!"
    End If
End Sub


