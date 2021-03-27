Attribute VB_Name = "Module1"
Option Explicit

Sub 新增工作表()
    Dim i As Integer
    
    Do Until Worksheets.Count = 12
        i = Worksheets.Count
        Worksheets.Add After:=Worksheets(i)
        ActiveSheet.Name = i + 1 & "月"
    Loop
End Sub



