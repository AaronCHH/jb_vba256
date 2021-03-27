Attribute VB_Name = "Module2"
Option Explicit

Sub 設定格式()
   Range("A5:A9").NumberFormatLocal = "mm/dd"
   Range("B5:B9").NumberFormatLocal = "h:mm AM/PM"
   Range("C3").NumberFormatLocal = """受理:""@"
   Range("A13:C13").NumberFormatLocal = "#,##0;[紅色]-#,##0"
End Sub


Sub 設定格式2()
    Cells.NumberFormat = "general"
    Range("A4:A9").NumberFormatLocal = "yyyy/m/d"
    Range("B5:B9").NumberFormatLocal = "h:mm"
End Sub
