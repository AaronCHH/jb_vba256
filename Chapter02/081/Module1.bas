Attribute VB_Name = "Module1"
Option Explicit

Sub �]�w����Ʈ榡()
    Dim myRange As Range, myFC As FormatCondition
    Set myRange = Worksheets(1).Range("E3:E14")
    If myRange.FormatConditions.Count > 0 Then myRange.FormatConditions.Delete
    Set myFC = myRange.FormatConditions.Add _
        (Type:=xlCellValue, Operator:=xlGreater, Formula1:="=1")
    myFC.Font.Bold = True
    myFC.Interior.Color = RGB(140, 180, 230)
    Set myRange = Nothing: Set myFC = Nothing
End Sub

Sub �]�w����Ʈ榡����ƾ��()
    Dim myRange As Range, myDB As DataBar
    Set myRange = Worksheets(1).Range("E3:E14")
    Set myDB = myRange.FormatConditions.Add(Type:=xlDatabar)
    myDB.BarColor.Color = RGB(150, 255, 100)
    Set myRange = Nothing: Set myDB = Nothing
End Sub


