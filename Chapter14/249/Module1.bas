Attribute VB_Name = "Module1"
Option Explicit

Sub �s��Word()
    Dim myWord As New Word.Application
    Range("��{��").Copy
    With myWord
        .Visible = True
        .Documents.Open ThisWorkbook.Path & "\���{��.doc"
        .Activate
        .Selection.MoveDown Unit:=wdLine, Count:=6
        .Selection.PasteExcelTable False, False, False
        .ActiveDocument.PrintOut
    End With
    Application.CutCopyMode = False
    Set myWord = Nothing
End Sub

Sub �s��Word2()
    Dim myWord As Object
    Set myWord = CreateObject("word.application")
    Range("��{��").Copy
    With myWord
        .Visible = True
        .Documents.Open ThisWorkbook.Path & "\���{��.doc"
        .Activate
        .Selection.MoveDown Unit:=5, Count:=6
        .Selection.PasteExcelTable False, False, False
        .ActiveDocument.PrintOut
    End With
    Application.CutCopyMode = False
    Set myWord = Nothing
End Sub
