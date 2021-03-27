Attribute VB_Name = "Module1"
Option Explicit

Sub 頁首設定()
    With ActiveSheet.PageSetup
        .LeftHeader = "&18&B" & Range("A3")
        .CenterHeader = "&A"
        .RightHeader = "列印日: " & "&D"
    End With
    ActiveSheet.PrintPreview
End Sub

