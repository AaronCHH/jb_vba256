Attribute VB_Name = "Module2"
Option Explicit

Sub �X���x�s��()
    Application.DisplayAlerts = False
    Range("A4:A9").Merge
    Range("B10:D13").Merge Across:=True
    Application.DisplayAlerts = True
End Sub

