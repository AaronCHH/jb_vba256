Attribute VB_Name = "Module1"
Option Explicit

Sub �T�{�M��()
    Dim ans As Integer
    ans = MsgBox(Range("A4") & "���g�t�d�H�A" & _
               Chr(10) & "�i�H�M����?", _
               vbYesNo + vbQuestion, "�M���T�{")
    If ans = vbYes Then
       Range("C4:D10").ClearContents
    End If
End Sub


