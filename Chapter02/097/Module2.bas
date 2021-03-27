Attribute VB_Name = "Module2"
Option Explicit

Sub 直接複製()
    Range("A5:F9").Copy Destination:=Range("A13")
End Sub

Sub 移動()
    Range("A5:D9").Cut Destination:=Range("A13")
End Sub

