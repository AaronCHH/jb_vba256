Attribute VB_Name = "Module1"
Option Explicit

Sub 設定視窗尺寸()
    Dim maxWidth As Double
    Dim maxHeight As Double
    Dim xWidth

    maxWidth = Application.UsableWidth
    maxHeight = Application.UsableHeight
    xWidth = 545

    Worksheets("送貨單").Activate
    With ActiveWindow
        .WindowState = xlNormal
        .Top = 0
        .Left = 0
        .Width = xWidth
        .Height = maxHeight
    End With
    
    ActiveWindow.NewWindow
    Worksheets("內容").Activate
    With ActiveWindow
        .Top = 0
        .Left = xWidth
        .Width = maxWidth - xWidth
        .Height = maxHeight
    End With
End Sub

Sub 切換檢視()
    Dim v As Integer
    v = Application.InputBox _
    (Prompt:="1:標準, 2:整頁, 3:分頁預覽", Type:=2)
    Select Case v
        Case 1: ActiveWindow.View = xlNormalView
        Case 2: ActiveWindow.View = xlPageLayoutView
        Case 3: ActiveWindow.View = xlPageBreakPreview
    End Select
End Sub


