Attribute VB_Name = "Module1"
Option Explicit

Sub 連結Word()
    Dim myWord As New Word.Application
    Range("行程表").Copy
    With myWord
        .Visible = True
        .Documents.Open ThisWorkbook.Path & "\月行程表.doc"
        .Activate
        .Selection.MoveDown Unit:=wdLine, Count:=6
        .Selection.PasteExcelTable False, False, False
        .ActiveDocument.PrintOut
    End With
    Application.CutCopyMode = False
    Set myWord = Nothing
End Sub

Sub 連結Word2()
    Dim myWord As Object
    Set myWord = CreateObject("word.application")
    Range("行程表").Copy
    With myWord
        .Visible = True
        .Documents.Open ThisWorkbook.Path & "\月行程表.doc"
        .Activate
        .Selection.MoveDown Unit:=5, Count:=6
        .Selection.PasteExcelTable False, False, False
        .ActiveDocument.PrintOut
    End With
    Application.CutCopyMode = False
    Set myWord = Nothing
End Sub
