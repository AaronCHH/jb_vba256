Attribute VB_Name = "Module1"
Option Explicit

Sub 眔计㎝逆计()
    Dim rcnt As Long, ccnt As Long
    
    rcnt = Rows.Count
    ccnt = Columns.Count
    MsgBox "计: " & rcnt & Chr(10) & _
           "逆计: " & ccnt
    rcnt = Range("A3:C10").Rows.Count
    ccnt = Range("A3:C10").Columns.Count
    MsgBox "计: " & rcnt & Chr(10) & _
           "逆计: " & ccnt

End Sub

