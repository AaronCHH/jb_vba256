Attribute VB_Name = "Module1"
Option Explicit

Sub �C�L()
    Dim x As Integer
    x = Application.InputBox(Prompt:="�п�ܹw���C�L�d��" & _
        Chr(10) & _
        "1: [3�Ь�]�u�@��: ����" & Chr(10) & _
        "2: [3�Ь�]�u�@��: �����" & Chr(10) & _
        "3: [3�Ь�]�u�@��: ���Ϫ�" & Chr(10) & _
        "4: ����ï���Ҧ��u�@��", Type:=1)
    Select Case x
        Case 1: ActiveSheet.PrintOut Preview:=True
        Case 2: ActiveSheet.Range("A1:F14").PrintOut Preview:=True
        Case 3: ActiveSheet.ChartObjects(1).Chart.PrintOut Preview:=True
        Case 4: ActiveWorkbook.PrintOut Preview:=True
        Case Else: Exit Sub
    End Select
End Sub



