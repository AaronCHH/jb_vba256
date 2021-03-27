Attribute VB_Name = "Module1"
Option Explicit

Sub 新增活頁簿()
    Dim ans As Integer
    ans = MsgBox("請問新活頁簿的工作表預設1張嗎???", vbYesNo)
    If ans = vbYes Then
       Application.SheetsInNewWorkbook = 1
    Else
       Application.SheetsInNewWorkbook = 3
    End If
    Workbooks.Add
End Sub


