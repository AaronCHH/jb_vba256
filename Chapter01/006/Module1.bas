Attribute VB_Name = "Module1"
Public pubID As String
Dim mojNickName As String

Sub text1()
    Dim proString As String
    
    proString = "���p��"
    pubID = "wang_sho_fa"
    mojNickName = "�p��"
    
    MsgBox "�m�W:" & proString & Chr(10) & _
           "ID:" & pubID & Chr(10) & _
           "�p�W:" & mojNickName & Chr(10)
    
End Sub

Sub text2()

    MsgBox mojNickName & "�p�j" & _
           "�Ȧw", , pubID

End Sub


