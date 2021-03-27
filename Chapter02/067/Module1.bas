Attribute VB_Name = "Module1"
Option Explicit

Sub 塗滿儲存格()
    Dim myRange As Range
    For Each myRange In Range("C4:D10")
        If myRange.Value = "陳" Then
           myRange.Interior.Color = RGB(150, 255, 100)
        End If
    Next
End Sub

Sub 設定儲存格的顏色()
    Range("C4:D10").Font.ColorIndex = xlColorIndexAutomatic
    Range("C4:D10").Interior.ColorIndex = xlColorIndexNone
End Sub

Sub 設定儲存格的布景主題()
    Range("A1").Interior.ThemeColor = xlThemeColorAccent6
End Sub

