Attribute VB_Name = "Module1"
Option Explicit

Sub �s�W�u�@��()
    Dim i As Integer
    
    Do Until Worksheets.Count = 12
        i = Worksheets.Count
        Worksheets.Add After:=Worksheets(i)
        ActiveSheet.Name = i + 1 & "��"
    Loop
End Sub



