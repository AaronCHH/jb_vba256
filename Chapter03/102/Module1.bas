Attribute VB_Name = "Module1"
Option Explicit

Sub 自動篩選()
   Dim Joken1 As String, Joken2 As String
   
   Joken1 = "台北分店"
   Joken2 = "DVD*"
   Range("A2").AutoFilter Field:=3, Criteria1:=Joken1
   Range("A2").AutoFilter Field:=5, Criteria1:=Joken2
   
   MsgBox "分店:" & Joken1 & ", 商品名稱:" & Joken2 & " 滿足條件資料: " _
   & Range("A2").CurrentRegion.Columns(1). _
   SpecialCells(xlCellTypeVisible).Count - 1 & "筆 "
   
   Range("A2").AutoFilter
End Sub


