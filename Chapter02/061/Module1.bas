Attribute VB_Name = "Module1"
Option Explicit

Sub 定義名稱()
    Range("A2").CurrentRegion.Name = "年間目標"
    Range("年間目標").Select
End Sub

Sub 定義名稱2()
    ActiveWorkbook.Names.Add Name:="年間目標", RefersTo:=Range("A2").CurrentRegion
    Range("年間目標").Select
End Sub


