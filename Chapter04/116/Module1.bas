Attribute VB_Name = "Module1"
Option Explicit

Sub �R���u�@��()
    Dim myWS As Worksheet
    Application.DisplayAlerts = False
    For Each myWS In Worksheets
        If myWS.Name Like "2006*" Then
           myWS.Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub



