VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    Cells(Rows.Count, 1).End(xlUp).Offset(1).Value = Now
    MsgBox "¥´¶}[ " & Me.Name & " ]¬¡­¶Ã¯!!"
End Sub


