Attribute VB_Name = "Module1"
Option Explicit

Sub 改變工作表的索引標籤顏色()
    Dim mySheet As Worksheet
    For Each mySheet In Worksheets
        Select Case Left(mySheet.Name, 4)
            Case "2006"
                mySheet.Tab.Color = RGB(80, 255, 255)
            Case "2007"
                mySheet.Tab.Color = RGB(255, 255, 80)
        End Select
    Next
End Sub






