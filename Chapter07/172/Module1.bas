Attribute VB_Name = "Module1"
Option Explicit

Sub 視窗的顯示和隱藏()
    Dim myWindows As String
    myWindows = ActiveWindow.Caption
    MsgBox "隱藏作用視窗!!"
    ActiveWindow.Visible = False
    MsgBox "再顯示作用視窗!!"
    Windows(myWindows).Visible = True
End Sub
