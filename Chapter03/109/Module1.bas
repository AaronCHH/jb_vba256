Attribute VB_Name = "Module1"
Option Explicit

Sub 執行查詢()
    Dim srcRange As Range, fndRange As Range
    Worksheets("查詢").Activate
    Set srcRange = Worksheets("Data").Range("A2").CurrentRegion.Columns(5)
    Set fndRange = srcRange.Find(what:=Range("B1").Value)
    If Not fndRange Is Nothing Then
       Cells(5, 1).Value = fndRange.Offset(, -4).Value
       Cells(5, 2).Value = fndRange.Offset(, -3).Value
       Cells(5, 3).Value = fndRange.Offset(, -2).Value
    Else
       MsgBox "沒有該商品!!"
    End If
End Sub

