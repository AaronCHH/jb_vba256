Attribute VB_Name = "Module1"
Option Explicit

Sub �s�W����ï()
    Dim ans As Integer
    ans = MsgBox("�аݷs����ï���u�@��w�]1�i��???", vbYesNo)
    If ans = vbYes Then
       Application.SheetsInNewWorkbook = 1
    Else
       Application.SheetsInNewWorkbook = 3
    End If
    Workbooks.Add
End Sub


