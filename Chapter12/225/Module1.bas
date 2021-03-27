Attribute VB_Name = "Module1"
Option Explicit

Sub 將圖插入最大值的儲存格()
    Dim max As Long, r As Range
    max = Application.WorksheetFunction.max(Range("B2:E5"))
    Set r = Range("B2:E5").Find(max)
    With ActiveSheet.Shapes.AddShape _
            (msoShapeOval, r.Left, r.Top, r.Width, r.Height)
        .Name = "MAX"
        .Fill.Visible = msoFalse
        .Line.ForeColor.RGB = RGB(150, 205, 0)
    End With
End Sub

Sub 設定圖形輪廓()
    ActiveSheet.Shapes(1).Line.ForeColor.RGB = RGB(255, 0, 0)
    ActiveSheet.Shapes(1).Line.Weight = 4
End Sub

Sub 設定圖形填色()
    ActiveSheet.Shapes(1).Fill.ForeColor.RGB = RGB(255, 255, 0)
    ActiveSheet.Shapes(2).Fill.PresetTextured msoTextureBouquet
    ActiveSheet.Shapes(3).Fill.OneColorGradient msoGradientFromCenter, 2, 1
End Sub



