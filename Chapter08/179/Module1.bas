Attribute VB_Name = "Module1"
Option Explicit

Sub �����]�w()
    With ActiveSheet.PageSetup
        .LeftHeader = "&18&B" & Range("A3")
        .CenterHeader = "&A"
        .RightHeader = "�C�L��: " & "&D"
    End With
    ActiveSheet.PrintPreview
End Sub

