Attribute VB_Name = "Module1"
Option Explicit

Sub �۰ʿz��()
   Dim Joken1 As String, Joken2 As String
   
   Joken1 = "�x�_����"
   Joken2 = "DVD*"
   Range("A2").AutoFilter Field:=3, Criteria1:=Joken1
   Range("A2").AutoFilter Field:=5, Criteria1:=Joken2
   
   MsgBox "����:" & Joken1 & ", �ӫ~�W��:" & Joken2 & " ����������: " _
   & Range("A2").CurrentRegion.Columns(1). _
   SpecialCells(xlCellTypeVisible).Count - 1 & "�� "
   
   Range("A2").AutoFilter
End Sub


