Attribute VB_Name = "Module2"
Option Explicit

Sub 參照字型()
    With Range("B1:C6").Font
        .Name = "MS PGothic"
        .FontStyle = "斜體"
        .Size = 16
        .Color = RGB(0, 115, 190)
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
End Sub

