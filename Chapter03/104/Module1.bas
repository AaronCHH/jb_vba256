Attribute VB_Name = "Module1"
Option Explicit

Sub �إߪ��()
    If ActiveSheet.ListObjects.Count <> 0 Then
       MsgBox "�w�g���n���!!"
    Else
       ActiveSheet.ListObjects.Add(xlSrcRange, _
           Range("A2").CurrentRegion, , xlYes).Name = "Table1"
       'Range("Table1").Select
    End If
End Sub

Sub �R�����()
    If ActiveSheet.ListObjects.Count <> 0 Then
        'ActiveSheet.ListObjects("Table1").TableStyle = ""
        ActiveSheet.ListObjects("Table1").Unlist
    End If
End Sub

Sub �ƧǩM���()
    'With ActiveSheet.ListObjects("Table1").Sort
    '   .SortFields.Clear
    '  .SortFields.Add Key:=Range("H2"), _
    '        SortOn:=xlSortOnValues, Order:=xlDescending
    '    .Header = xlYes
    '    .Apply
    'End With
     ActiveSheet.ListObjects("Table1").Range.AutoFilter _
     Field:=3, Criteria1:="�x�_����"
End Sub





