Attribute VB_Name = "Module1"
Option Explicit

Sub 新增樣式()
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
   MsgBox "樣式名重複!!"
End Sub

Sub 刪除樣式()
    On Error Resume Next
    ActiveWorkbook.Styles("myTitle").Delete
End Sub
