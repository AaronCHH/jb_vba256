Attribute VB_Name = "Module1"
Option Explicit

Sub �w���C�L()
    Dim x As Integer
    x = Application.InputBox(Prompt:="�п�ܹw���C�L�d��" & _
        Chr(10) & _
        "1: 3�Ь�u�@��w��" & Chr(10) & _
        "2: �U�����w��" & Chr(10) & _
        "3: ��اO�u�@��w��", Type:=1)
    Select Case x
        Case 1: ActiveSheet.PrintPreview False
        Case 2: ActiveSheet.Range("A1:F14,A17:F30,A33:G46").PrintPreview
        Case 3: Worksheets(Array(2, 3, 4)).PrintPreview False
        Case Else: Exit Sub
    End Select
End Sub

