Attribute VB_Name = "Module1"
Option Explicit

Sub �T�{�u�@��W()
    Dim myWS As Worksheet, myName As String
    
    myName = Format(Date, "yyyy-mm")
    For Each myWS In Worksheets
        If myWS.Name = myName Then
           MsgBox "�P�W���u�@��s�b!!"
           Exit Sub
        End If
    Next
    Worksheets("Template").Copy Before:=Worksheets(1)
    ActiveSheet.Name = myName
End Sub



