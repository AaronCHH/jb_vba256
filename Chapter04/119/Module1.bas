Attribute VB_Name = "Module1"

Sub �u�@����ܩ�����()

    With Worksheets("Template")
      If .Visible = True Then
         .Visible = xlSheetVeryHidden
      Else
         .Visible = True
      End If
    End With
End Sub


