Attribute VB_Name = "Module1"
Option Explicit

Sub �x�s�檺�e�M��()
    Range("A1").Select
    MsgBox "�x�s��A1������: " & Selection.Height & "�I" & Chr(10) & _
           "�x�s��A1���e��: " & Selection.Width & "�I"
    ActiveSheet.UsedRange.Select
    MsgBox "�ϥ��x�s��d�򪺰���: " & Selection.Height & "�I" & Chr(10) & _
           "�ϥ��x�s��d�򪺼e��: " & Selection.Width & "�I"
End Sub

Sub �w�νd�򪺼e�M��()
    ActiveSheet.UsedRange.Select
    MsgBox "�ϥ��x�s��d�򪺰���: " & Selection.Height * 0.035 & "cm" & Chr(10) & _
           "�ϥ��x�s��d�򪺼e��: " & Selection.Width * 0.035 & "cm"
End Sub
