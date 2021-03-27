Attribute VB_Name = "Module1"
Option Explicit

Sub 儲存格的寬和高()
    Range("A1").Select
    MsgBox "儲存格A1的高度: " & Selection.Height & "點" & Chr(10) & _
           "儲存格A1的寬度: " & Selection.Width & "點"
    ActiveSheet.UsedRange.Select
    MsgBox "使用儲存格範圍的高度: " & Selection.Height & "點" & Chr(10) & _
           "使用儲存格範圍的寬度: " & Selection.Width & "點"
End Sub

Sub 已用範圍的寬和高()
    ActiveSheet.UsedRange.Select
    MsgBox "使用儲存格範圍的高度: " & Selection.Height * 0.035 & "cm" & Chr(10) & _
           "使用儲存格範圍的寬度: " & Selection.Width * 0.035 & "cm"
End Sub
