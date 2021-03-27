Attribute VB_Name = "Module1"
Option Explicit

Sub 圖表範圍變更()
    Dim sd As Variant, gRange As Range
    Charts("圖表").Activate
    sd = InputBox("指定圖表化的科目:  國語: 1, 英語 2, 數學 3")
    Select Case sd
        Case 1: Set gRange = Worksheets("3教科").Range("B3:D13")
        Case 2: Set gRange = Worksheets("3教科").Range("B19:D29")
        Case 3: Set gRange = Worksheets("3教科").Range("B35:E45")
        Case Else
            MsgBox "指定不正確!!"
            Exit Sub
    End Select
    Charts("圖表").SetSourceData gRange
    Set gRange = Nothing
End Sub




