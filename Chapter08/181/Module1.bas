Attribute VB_Name = "Module1"
Option Explicit

Sub �C�L�d��]�w()
    With ActiveSheet
        .PageSetup.PrintArea = "A1:F30"
        .PrintPreview
        .PageSetup.PrintArea = Range("��y").Address
        .PrintPreview
    End With
End Sub
