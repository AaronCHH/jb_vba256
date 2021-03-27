Attribute VB_Name = "Module1"
Option Explicit

Sub 持續執行查詢()
    Dim srcRange As Range, fndRange As Range
    Dim fstAddress As String, i As Integer
    Worksheets("查詢").Activate
    Set srcRange = Worksheets("Data").Range("A2").CurrentRegion.Columns(5)
    Set fndRange = srcRange.Find(what:=Range("B1").Value)
    If Not fndRange Is Nothing Then
        fstAddress = fndRange.Address
        i = 5
        Do
            Cells(i, 1).Value = fndRange.Offset(, -4).Value
            Cells(i, 2).Value = fndRange.Offset(, -3).Value
            Cells(i, 3).Value = fndRange.Offset(, -2).Value
            Set fndRange = srcRange.FindNext(after:=fndRange)
            i = i + 1
        Loop Until fndRange.Address = fstAddress
    Else
       MsgBox "沒有該商品!!"
    End If
End Sub

