Attribute VB_Name = "Module1"
Option Explicit

Sub 選擇印表機對話方塊()
    Dim myPrinter As String
    myPrinter = Application.ActivePrinter
    If Application.Dialogs(xlDialogPrinterSetup).Show Then
       ActiveSheet.PrintPreview
       Application.ActivePrinter = myPrinter
    End If
End Sub

