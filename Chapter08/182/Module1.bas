Attribute VB_Name = "Module1"
Option Explicit

Sub �����L���()
    Dim myPrinter As String
    myPrinter = Application.ActivePrinter
    MsgBox "�{�b���L���: " & myPrinter & Chr(10) & _
           "������e�f��ΦL���!!"
    ActiveSheet.PrintOut preview:=True, ActivePrinter:="Printer101"
    Application.ActivePrinter = myPrinter
End Sub



