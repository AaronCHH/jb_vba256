Attribute VB_Name = "Module1"
Option Explicit

Sub ����d��()
    Dim srcRange As Range, fndRange As Range
    Worksheets("�d��").Activate
    Set srcRange = Worksheets("Data").Range("A2").CurrentRegion.Columns(5)
    Set fndRange = srcRange.Find(what:=Range("B1").Value)
    If Not fndRange Is Nothing Then
       Cells(5, 1).Value = fndRange.Offset(, -4).Value
       Cells(5, 2).Value = fndRange.Offset(, -3).Value
       Cells(5, 3).Value = fndRange.Offset(, -2).Value
    Else
       MsgBox "�S���Ӱӫ~!!"
    End If
End Sub

