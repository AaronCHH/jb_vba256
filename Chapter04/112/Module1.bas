Attribute VB_Name = "Module1"
Option Explicit

Sub 參照工作表()
    Dim i As Integer
    With Worksheets("Sheet1")
        .Range("B1").Value = Sheets.Count
        .Range("B2").Value = Worksheets.Count
        .Range("B3").Value = Charts.Count
    End With
    MsgBox ActiveSheet.Name
End Sub

