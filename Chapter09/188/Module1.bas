Attribute VB_Name = "Module1"
Option Explicit

Sub ���o�ɶ����()
    Dim dp As Integer
    dp = DatePart("y", Date)
    MsgBox "���Ѫ����: " & Date & Chr(10) & _
           "1��1���{�b���g�L���: " & dp & Chr(10) & _
           "1�~���g�L�F: " & Format(dp / 365, "0.0%")
End Sub
