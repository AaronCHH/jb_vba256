Attribute VB_Name = "Module1"
Option Explicit

Sub ¤j¼g¤p¼gÂà´«()
    Dim i As Integer
    
    For i = 2 To 4
        Cells(i, 2) = UCase(Cells(i, 1))
        Cells(i, 3) = LCase(Cells(i, 1))
    Next
End Sub
