Attribute VB_Name = "Module1"
Option Explicit

Sub ¸m´«¦r¦ê()
    Range("B2").Value = Replace(Range("B1").Value, " ", "")
    Range("B3").Value = Replace(Range("B1").Value, " ", Chr(10))
End Sub

Sub ¸m´«¦r¦ê2()
    Range("A2:A4").Replace what:=" ", Replacement:=Chr(10)
    Rows.AutoFit
End Sub


