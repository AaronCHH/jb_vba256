Attribute VB_Name = "Module1"
Option Explicit

Sub 商品別分頁()
    Dim i As Integer
    ActiveSheet.ResetAllPageBreaks
    i = 3
    Do While Cells(i, 4) <> ""
       If Cells(i, 4).Value <> Cells(i + 1, 4).Value Then
          ActiveSheet.HPageBreaks.Add before:=Cells(i + 1, 4)
       End If
       i = i + 1
    Loop
    ActiveSheet.PageSetup.PrintTitleRows = "$1:$2"
    ActiveSheet.PrintPreview
End Sub

