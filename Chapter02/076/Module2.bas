Attribute VB_Name = "Module2"
Option Explicit

Sub �]�w�榡()
   Range("A5:A9").NumberFormatLocal = "mm/dd"
   Range("B5:B9").NumberFormatLocal = "h:mm AM/PM"
   Range("C3").NumberFormatLocal = """���z:""@"
   Range("A13:C13").NumberFormatLocal = "#,##0;[����]-#,##0"
End Sub


Sub �]�w�榡2()
    Cells.NumberFormat = "general"
    Range("A4:A9").NumberFormatLocal = "yyyy/m/d"
    Range("B5:B9").NumberFormatLocal = "h:mm"
End Sub
