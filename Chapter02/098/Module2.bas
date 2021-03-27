Attribute VB_Name = "Module2"
Option Explicit

Sub 選擇格式貼上()
    Range("A5:D9").Copy
    Range("F5").PasteSpecial Paste:=xlPasteColumnWidths
    Range("F5").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub

