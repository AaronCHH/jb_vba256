Attribute VB_Name = "Module1"
Option Explicit

Sub 參照空白儲存格()
    Dim myRange As Range
    Set myRange = Range("A2").CurrentRegion.SpecialCells(xlCellTypeBlanks)
    myRange.Value = 0
    MsgBox "空白儲存格: " & myRange.Address(False, False, xlA1)
    Set myRange = Nothing
End Sub

Sub 將儲存格轉為空白()
    
    Range("A2").CurrentRegion.Replace _
            What:=0, Replacement:="", Lookat:=xlWhole

End Sub

