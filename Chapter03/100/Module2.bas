Attribute VB_Name = "Module2"
Option Explicit

Sub AutoFill()
    Range("A3").AutoFill Destination:=Range("A3:A33"), Type:=xlFillSeries
    Range("B2").AutoFill Destination:=Range("B2:M2")
    Range("A3:A33").AutoFill Destination:=Range("A3:M33"), Type:=xlFillFormats
End Sub

