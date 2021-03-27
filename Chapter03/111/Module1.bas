Attribute VB_Name = "Module1"
Option Explicit

Sub °õ¦æ¸m´«()
    Range("B2:B8").Replace What:="2003", Replacement:="2007"
    Range("B2:B8").Replace What:=".NET", Replacement:="2005"
End Sub

Sub ¸m´«®æ¦¡()
    Application.FindFormat.Interior.Color = RGB(153, 255, 153)
    Application.ReplaceFormat.Interior.Color = RGB(255, 255, 102)
    ActiveSheet.UsedRange.Replace What:="", Replacement:="", _
            SearchFormat:=True, ReplaceFormat:=True
End Sub


