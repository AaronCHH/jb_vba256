Attribute VB_Name = "Module1"
Option Explicit

Sub 設定複數圖形的格式()
    ActiveSheet.Shapes.SelectAll
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset10
    ActiveSheet.Shapes.Range(Array("橢圓 3", "七角星形 2")).Fill.Visible = False
End Sub
