Attribute VB_Name = "Module1"
Option Explicit

Sub 複製工作表()
    Dim myMonth, myYear
    myMonth = Right(Worksheets(1).Name, 2)
    myYear = Left(Worksheets(1).Name, 4)
    
    Worksheets("Template").Copy Before:=Worksheets(1)
    If myMonth = 12 Then
       ActiveSheet.Name = myYear + 1 & "-01"
    Else
       ActiveSheet.Name = myYear & "-" & Format(myMonth + 1, "00")
    End If
End Sub



