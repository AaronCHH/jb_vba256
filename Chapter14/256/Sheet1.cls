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

Sub ����Excel()
    Dim ans As Integer
    ans = MsgBox("�����i�H��??", vbOKCancel, "�����T�{")
    If ans = vbOK Then
       ActiveWorkbook.Save
       Application.Quit
    End If
End Sub

