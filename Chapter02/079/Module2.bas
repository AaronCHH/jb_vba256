Attribute VB_Name = "Module2"
Option Explicit

Sub 參照合併儲存格()
    Range("A3").Value = 2006
    Range("A4").Value = "12月"
    Range("A5").MergeArea.Value = 2007
    Range("A6").MergeArea.Value = "1月"
    Range("B3").MergeArea.ClearContents
    Range("B5").ClearContents
End Sub

