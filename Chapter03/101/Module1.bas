Attribute VB_Name = "Module1"
Option Explicit

Sub 排序()
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("C2"), SortOn:=xlSortOnValues, _
         Order:=xlAscending
        .SortFields.Add Key:=Range("H2"), SortOn:=xlSortOnValues, _
         Order:=xlDescending
        .SetRange Range("A2:H27")
        .Header = xlYes
        .Apply
    End With
End Sub

Sub 使用者自定排序()
    With ActiveSheet.Sort
        .SortFields.Add Key:=Range("C2"), SortOn:=xlSortOnValues, _
            Order:=xlAscending, CustomOrder:="台北分店, 新竹分店, 台中分店, 高雄分店"
        .SetRange Range("A2:H27")
        .Header = xlYes
        .Apply
    End With
End Sub

