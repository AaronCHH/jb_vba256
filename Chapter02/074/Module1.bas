Attribute VB_Name = "Module1"
Option Explicit

Sub 設定儲存格樣式()
    Range("B1").Style = "標題"
    Range("B3:C3").Style = "60% - 輔色3"
    Range("B4:B9").Style = "40% - 輔色3"
    Range("C4:C9").Style = "20% - 輔色3"
    Range("C9").Style = "百分比"
End Sub

