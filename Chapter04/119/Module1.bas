Attribute VB_Name = "Module1"

Sub 工作表的顯示或隱藏()

    With Worksheets("Template")
      If .Visible = True Then
         .Visible = xlSheetVeryHidden
      Else
         .Visible = True
      End If
    End With
End Sub


