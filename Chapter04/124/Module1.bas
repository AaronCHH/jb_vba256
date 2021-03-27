Attribute VB_Name = "Module1"
Option Explicit

Sub 工作表的排序()
    Dim i As Integer, sRange As Range, sName As String
    Worksheets.Add(before:=Worksheets(1)).Name = "temp"
    For i = 2 To Worksheets.Count
        Cells(i, 1) = Worksheets(i).Name
    Next
    Set sRange = Cells(2, 1).CurrentRegion
    
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A2"), SortOn:=xlSortOnValues, _
                        Order:=xlAscending
        .SetRange sRange
        .Apply
    End With
    For i = 1 To sRange.Rows.Count
        sName = Worksheets("temp").Cells(i + 1, 1)
        Worksheets(sName).Move after:=ActiveSheet
    Next
    Application.DisplayAlerts = False
    Worksheets("temp").Delete
    Application.DisplayAlerts = True
End Sub


