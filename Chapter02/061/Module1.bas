Attribute VB_Name = "Module1"
Option Explicit

Sub �w�q�W��()
    Range("A2").CurrentRegion.Name = "�~���ؼ�"
    Range("�~���ؼ�").Select
End Sub

Sub �w�q�W��2()
    ActiveWorkbook.Names.Add Name:="�~���ؼ�", RefersTo:=Range("A2").CurrentRegion
    Range("�~���ؼ�").Select
End Sub


