Attribute VB_Name = "Module1"
Option Explicit

Sub �ܧ�˦�()
    ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium5"
    MsgBox "�ܧ�" & ActiveSheet.ListObjects("Table1").TableStyle.NameLocal _
           & "!!"
End Sub


