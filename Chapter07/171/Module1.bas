Attribute VB_Name = "Module1"
Option Explicit

Sub ���ù����()
    Application.DisplayFullScreen = True
    Range("�e�f��").Select
    With ActiveWindow
        .DisplayHeadings = False
        .DisplayGridlines = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .Zoom = True
    End With
    ActiveSheet.ScrollArea = "A1:E21"
    Range("A1").Select
End Sub

