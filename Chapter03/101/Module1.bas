Attribute VB_Name = "Module1"
Option Explicit

Sub �Ƨ�()
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

Sub �ϥΪ̦۩w�Ƨ�()
    With ActiveSheet.Sort
        .SortFields.Add Key:=Range("C2"), SortOn:=xlSortOnValues, _
            Order:=xlAscending, CustomOrder:="�x�_����, �s�ˤ���, �x������, ��������"
        .SetRange Range("A2:H27")
        .Header = xlYes
        .Apply
    End With
End Sub

