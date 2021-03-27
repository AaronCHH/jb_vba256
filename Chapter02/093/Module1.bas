Attribute VB_Name = "Module1"
Option Explicit

Sub 變更對齊方式()
    Range("A1:B1").HorizontalAlignment = xlCenterAcrossSelection
    Range("A3:B6").HorizontalAlignment = xlCenter
    Range("A3:B6").VerticalAlignment = xlCenter
End Sub

Sub 變更對齊方式2()
    Range("A3:B3").Orientation = 30
    Range("A4:A6").Orientation = xlVertical
End Sub
