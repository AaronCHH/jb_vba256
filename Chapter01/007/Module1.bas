Attribute VB_Name = "Module1"
Option Explicit

Sub ObjectTest()

    Dim myRange As Range
    Set myRange = Range("A1:C3")
    
    myRange.Select
    myRange.Borders.LineStyle = xlContinuous
    
    Set myRange = Nothing
    
End Sub

