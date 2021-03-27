Attribute VB_Name = "Module1"
Option Explicit

Sub ¦C¦L½d³ò³]©w()
    With ActiveSheet
        .PageSetup.PrintArea = "A1:F30"
        .PrintPreview
        .PageSetup.PrintArea = Range("°ê»y").Address
        .PrintPreview
    End With
End Sub
