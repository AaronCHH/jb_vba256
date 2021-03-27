Attribute VB_Name = "Module1"
Option Explicit

Sub 新增文字方塊()
    Dim r As Range
    Set r = Range("A8:F9")
    With ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        r.Left, r.Top, r.Width, r.Height)
        .Name = "Text1"
        .TextFrame.Characters.Text = "HDD錄影機業績佳!!!"
    End With
    
End Sub




