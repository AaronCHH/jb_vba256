VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim myRange As Range, myflag As Boolean
    myflag = False
    For Each myRange In Range("A7:C7")
        If myRange.Value = "" Then
           myflag = True
        End If
    Next
    If myflag Then
       MsgBox "¦³º|°O!!!"
       Cancel = True
    Else
       ActiveWorkbook.SaveAs Format(Date, "mmdd") & Range("C2").Value
    End If
End Sub

