Attribute VB_Name = "Module1"
Option Explicit

Sub �վ�A�M��檺�����ؤo()
    With ActiveWindow
        .WindowState = xlNormal
        .Top = 0
        .Left = 120
        .Width = Range("�e�f��").Width + 55
        .Height = Range("�e�f��").Height + 60
    End With
End Sub
