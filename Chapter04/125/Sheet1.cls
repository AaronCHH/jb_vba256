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

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo errhandler
    If Target.Column <> 2 Then Exit Sub
    If Target.Offset(, -1) <> "" Then
       Target.Offset(, 1) = Date
    End If
    Exit Sub
errhandler:
End Sub

Private Sub Worksheet_Activate()
    Dim ans As Integer
    ans = MsgBox("�i�H�R���@�~���e�M�@�~���!!", _
        vbYesNo, "�R���T�{")
    If ans = vbYes Then
        Range("B4:C7").ClearContents
    End If
End Sub


