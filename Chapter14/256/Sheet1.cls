VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub 結束Excel()
    Dim ans As Integer
    ans = MsgBox("結束可以嗎??", vbOKCancel, "結束確認")
    If ans = vbOK Then
       ActiveWorkbook.Save
       Application.Quit
    End If
End Sub


