Attribute VB_Name = "Module1"
Option Explicit

Sub 確認清除()
    Dim ans As Integer
    ans = MsgBox(Range("A4") & "的週負責人，" & _
               Chr(10) & "可以清除嗎?", _
               vbYesNo + vbQuestion, "清除確認")
    If ans = vbYes Then
       Range("C4:D10").ClearContents
    End If
End Sub


