Attribute VB_Name = "Module1"
Option Explicit

Sub �]�w�����ؤo()
    Dim maxWidth As Double
    Dim maxHeight As Double
    Dim xWidth

    maxWidth = Application.UsableWidth
    maxHeight = Application.UsableHeight
    xWidth = 545

    Worksheets("�e�f��").Activate
    With ActiveWindow
        .WindowState = xlNormal
        .Top = 0
        .Left = 0
        .Width = xWidth
        .Height = maxHeight
    End With
    
    ActiveWindow.NewWindow
    Worksheets("���e").Activate
    With ActiveWindow
        .Top = 0
        .Left = xWidth
        .Width = maxWidth - xWidth
        .Height = maxHeight
    End With
End Sub

Sub �����˵�()
    Dim v As Integer
    v = Application.InputBox _
    (Prompt:="1:�з�, 2:�㭶, 3:�����w��", Type:=2)
    Select Case v
        Case 1: ActiveWindow.View = xlNormalView
        Case 2: ActiveWindow.View = xlPageLayoutView
        Case 3: ActiveWindow.View = xlPageBreakPreview
    End Select
End Sub


