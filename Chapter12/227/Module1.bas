Attribute VB_Name = "Module1"
Option Explicit

Sub �s�W��r���()
    Dim r As Range
    Set r = Range("A8:F9")
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        r.Left, r.Top, r.Width, r.Height)
        .Name = "Text1"
        .TextFrame.Characters.Text = "HDD���v���~�Z��!!!"
    End With
    
End Sub




