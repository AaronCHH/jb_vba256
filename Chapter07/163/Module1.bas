Attribute VB_Name = "Module1"
Option Explicit

Sub 變更視窗尺寸()
    ActiveWindow.WindowState = xlNormal
    MsgBox "作用視窗恢復為原有尺寸!!"
    ActiveWindow.WindowState = xlMaximized
    MsgBox "作用視窗最大化!!"
End Sub



