VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub BookOpen()
    Dim xBook As Integer, i As Integer
    xBook = Array("���D1.xls", "���D2.xls", "���D3.xls")
    For i = 0 To 2
        Worksheet.Open Filename:=xBook(i)
    Next
    MsgBox Worksheet(1).Name & ":" & Worksheet(2).Name & ":" & _
        Worksheet(3).Name & ":" & Worksheet(4).Name & ":"
End Sub


