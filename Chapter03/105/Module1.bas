Attribute VB_Name = "Module1"
Option Explicit

Sub 變更樣式()
    ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium5"
    MsgBox "變更為" & ActiveSheet.ListObjects("Table1").TableStyle.NameLocal _
           & "!!"
End Sub


