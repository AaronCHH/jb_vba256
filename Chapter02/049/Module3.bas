Attribute VB_Name = "Module3"
Option Explicit

Sub �@���x�s��()
    Range("A1").Select
    ActiveSheet.Shapes("Picture 1").Select
    MsgBox "�@���x�s��: " & ActiveCell.Address _
    & Chr(10) & "���: " & TypeName(Selection)
End Sub


