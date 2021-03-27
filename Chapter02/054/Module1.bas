Attribute VB_Name = "Module1"
Option Explicit

Sub 變更儲存格範圍()
    Dim myRange As Range
    Dim myRow As Integer, myCol As Integer
    
    Set myRange = Range("A3").CurrentRegion
    myRow = myRange.Rows.Count
    myCol = myRange.Columns.Count
    Set myRange = myRange.Resize(myRow - 1, myCol - 1)
    
    ActiveSheet.Shapes.AddChart(xlPie, 10, 200).Select
    ActiveChart.SetSourceData Source:=myRange
    Set myRange = Nothing
End Sub

