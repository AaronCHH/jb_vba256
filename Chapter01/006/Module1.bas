Attribute VB_Name = "Module1"
Public pubID As String
Dim mojNickName As String

Sub text1()
    Dim proString As String
    
    proString = "王小花"
    pubID = "wang_sho_fa"
    mojNickName = "小花"
    
    MsgBox "姓名:" & proString & Chr(10) & _
           "ID:" & pubID & Chr(10) & _
           "小名:" & mojNickName & Chr(10)
    
End Sub

Sub text2()

    MsgBox mojNickName & "小姐" & _
           "午安", , pubID

End Sub


