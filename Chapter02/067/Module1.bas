Attribute VB_Name = "Module1"
Option Explicit

Sub ��x�s��()
    Dim myRange As Range
    For Each myRange In Range("C4:D10")
        If myRange.Value = "��" Then
           myRange.Interior.Color = RGB(150, 255, 100)
        End If
    Next
End Sub

Sub �]�w�x�s�檺�C��()
    Range("C4:D10").Font.ColorIndex = xlColorIndexAutomatic
    Range("C4:D10").Interior.ColorIndex = xlColorIndexNone
End Sub

Sub �]�w�x�s�檺�����D�D()
    Range("A1").Interior.ThemeColor = xlThemeColorAccent6
End Sub

