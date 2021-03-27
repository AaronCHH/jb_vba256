Attribute VB_Name = "Module1"
Option Explicit

Sub Clear()
    Range("A1:D1").ClearFormats
    Range("A5:C7").ClearContents
    Range("D3").ClearComments
    Range("A10:D10").Clear
End Sub

