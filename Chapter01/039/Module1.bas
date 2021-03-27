Attribute VB_Name = "Module1"
Option Explicit

Sub 呼叫其他活頁簿的程序()
    Workbooks.Open "C:\ExcelVBA\Book1.xls"
    Application.Run "Book1.xls!Sample"
End Sub

Sub 呼叫其他活頁簿的程序2()
    Application.Run "'C:\ExcelVBA\Book1.xls'!Sample"
End Sub

