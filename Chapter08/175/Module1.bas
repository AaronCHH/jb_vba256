Attribute VB_Name = "Module1"
Option Explicit

Sub �C�L�]�w()
    With ActiveSheet.PageSetup
        .PrintArea = "$A$1:$G$46"
        .CenterFooter = "�� &P ��/�@ &N ��"
        .CenterHorizontally = True
    End With
    ActiveSheet.PrintOut
End Sub

