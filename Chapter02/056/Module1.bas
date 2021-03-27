Attribute VB_Name = "Module1"
Option Explicit

Sub 取得位址()
    Dim myRange As Range
    Set myRange = Range("A3").CurrentRegion
    MsgBox "儲存格A3的目前區域作用領域: " & myRange.Address(RowAbsolute:=False, _
            ColumnAbsolute:=False, ReferenceStyle:=xlA1, External:=True)
    Set myRange = Nothing
End Sub


