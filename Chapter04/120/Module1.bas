Attribute VB_Name = "Module1"
Option Explicit

Sub 確認工作表名()
    Dim myWS As Worksheet, myName As String
    
    myName = Format(Date, "yyyy-mm")
    For Each myWS In Worksheets
        If myWS.Name = myName Then
           MsgBox "同名的工作表存在!!"
           Exit Sub
        End If
    Next
    Worksheets("Template").Copy Before:=Worksheets(1)
    ActiveSheet.Name = myName
End Sub



