Attribute VB_Name = "Module1"
Option Explicit

Sub �s�W�˦�()
    On Error GoTo errHandler
    With ActiveWorkbook.Styles.Add(Name:="myTitle")
         .Interior.ColorIndex = 38
         .HorizontalAlignment = xlHAlignCenter
         .Font.Size = 16
         .Font.ColorIndex = 56
    End With
    Range("B1:C1").Style = "myTitle"
    Exit Sub
errHandler:
   MsgBox "�˦��W����!!"
End Sub

Sub �R���˦�()
    On Error Resume Next
    ActiveWorkbook.Styles("myTitle").Delete
End Sub
