Attribute VB_Name = "Module2"
Option Explicit

Sub �ѷӦX���x�s��()
    Range("A3").Value = 2006
    Range("A4").Value = "12��"
    Range("A5").MergeArea.Value = 2007
    Range("A6").MergeArea.Value = "1��"
    Range("B3").MergeArea.ClearContents
    Range("B5").ClearContents
End Sub

