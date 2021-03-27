Attribute VB_Name = "Module1"
Option Explicit

Sub 切換印表機()
    Dim myPrinter As String
    myPrinter = Application.ActivePrinter
    MsgBox "現在的印表機: " & myPrinter & Chr(10) & _
           "切換到送貨單用印表機!!"
    ActiveSheet.PrintOut preview:=True, ActivePrinter:="Printer101"
    Application.ActivePrinter = myPrinter
End Sub



