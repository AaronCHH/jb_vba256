Attribute VB_Name = "Module2"
Option Explicit

Sub 設定字型的尺寸種類()
    With Range("B1").Font
        .Name = "MS UI Gothic"
        .Size = 18
    End With
    With Range("B3:C6").Font
        .Name = Application.StandardFont
        .Size = Application.StandardFontSize
    End With
End Sub

