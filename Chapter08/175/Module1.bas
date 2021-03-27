Attribute VB_Name = "Module1"
Option Explicit

Sub 列印設定()
    With ActiveSheet.PageSetup
        .PrintArea = "$A$1:$G$46"
        .CenterFooter = "第 &P 頁/共 &N 頁"
        .CenterHorizontally = True
    End With
    ActiveSheet.PrintOut
End Sub

