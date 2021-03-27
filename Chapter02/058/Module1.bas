Attribute VB_Name = "Module1"
Option Explicit

Sub 列欄編號()
    Dim allRange As Range
    Set allRange = Range("A3").CurrentRegion
    
    allRange.Columns("A:B").HorizontalAlignment = xlCenter
    allRange.Rows(1).Select
    Selection.HorizontalAlignment = xlCenter
    Selection.Interior.ColorIndex = 38
    Set allRange = Nothing
End Sub

Sub 日期格式()
    Cells.ClearFormats
    Columns(1).NumberFormatLocal = "yyyy/mm/dd"
End Sub

