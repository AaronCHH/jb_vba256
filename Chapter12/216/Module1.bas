Attribute VB_Name = "Module1"
Option Explicit

Sub ���J�Ϫ�()
    With Charts.Add(after:=ActiveSheet)
        .Name = "��XG"
        .SetSourceData Sheets("��X").Range("B3:E13")
    End With
End Sub




