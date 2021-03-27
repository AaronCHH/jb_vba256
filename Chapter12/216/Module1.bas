Attribute VB_Name = "Module1"
Option Explicit

Sub 插入圖表()
    With Charts.Add(after:=ActiveSheet)
        .Name = "綜合G"
        .SetSourceData Sheets("綜合").Range("B3:E13")
    End With
End Sub




