Attribute VB_Name = "Module1"
Option Explicit

Sub Loop5()
    Dim i As Integer
    i = 3
    
    Do While Cells(i, 2).Value <> ""
      'myWeight = Cells(i, "B").Value
       If Cells(i, 2).Value <= Cells(1, 2).Value Then
          Cells(i, 3).Value = "達成!"
          Exit Do
       End If
       i = i + 1
    Loop
    
End Sub

Sub Loop6()
    Dim i As Integer
    For i = 3 To 11
        Select Case Cells(i, 2).Value
            Case ""
                MsgBox Cells(i, 1).Text & "沒有數字!"
                Exit Sub
            Case Is <= Cells(1, 2).Value
                Cells(i, 3).Value = "達成!"
                Exit For
        End Select
    Next
    MsgBox "結束!!!"
End Sub

