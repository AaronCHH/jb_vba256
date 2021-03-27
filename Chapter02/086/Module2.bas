Attribute VB_Name = "Module2"
Option Explicit

Sub 設定或取得值()
    Range("A5").Value = #5/7/2007#
    Range("B5").Value = #10:30:00 AM#
    Range("C5").Value = "王大董"
    
    Range("A13").Value = Range("B13")
    Range("B13").Value = 16
End Sub

