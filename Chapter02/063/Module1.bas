Attribute VB_Name = "Module1"
Option Explicit

Sub �ѷӪť��x�s��()
    Dim myRange As Range
    Set myRange = Range("A2").CurrentRegion.SpecialCells(xlCellTypeBlanks)
    myRange.Value = 0
    MsgBox "�ť��x�s��: " & myRange.Address(False, False, xlA1)
    Set myRange = Nothing
End Sub

Sub �N�x�s���ର�ť�()
    
    Range("A2").CurrentRegion.Replace _
            What:=0, Replacement:="", Lookat:=xlWhole

End Sub

