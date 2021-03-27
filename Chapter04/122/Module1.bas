Attribute VB_Name = "Module1"
Option Explicit

Sub 指定儲存格移動範圍()
    Worksheets("英語測試").ScrollArea = Range("A3").CurrentRegion.Address
End Sub

Sub 解除儲存格移動範圍()
    Worksheets("英語測試").ScrollArea = ""
End Sub



