Attribute VB_Name = "Module1"
Option Explicit

Sub ���ʤu�@��()
    Dim myWS As Worksheet
    For Each myWS In Worksheets
        If Left(myWS.Name, 4) = "2006" Then
           myWS.Move before:=Workbooks("2006�~.xls").Worksheets(1)
        End If
    Next
End Sub

