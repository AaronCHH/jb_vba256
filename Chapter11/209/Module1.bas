Attribute VB_Name = "Module1"
Option Explicit

Sub ����T�{()
    Dim hizuke As String
    hizuke = InputBox("�п�J���~�����!!")
    If IsDate(hizuke) Then
       MsgBox Format(CDate(hizuke), "yyyy�~mm��dd��")
    Else
       MsgBox "�п�J���T���!!"
    End If
End Sub

Sub ����T�{2()
    Dim hizuke As String
    hizuke = InputBox("�п�J���~�����!!")
    If IsDate(hizuke) Then
       If Year(CDate(hizuke)) = Year(Date) Then
          MsgBox Format(CDate(hizuke), "yyyy�~mm��dd��")
       Else
          MsgBox "�п�J���~�����!!"
       End If
    Else
       MsgBox "�п�J���T���!!"
    End If
End Sub
