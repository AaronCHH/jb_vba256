Attribute VB_Name = "Module1"
Option Explicit

Sub 刪除字串空白()
    Dim strText As String
    strText = "     台北市士林區重慶北路四段1500號     "
    
    MsgBox "字串        :  [ " & strText & " ]" & Chr(10) & _
           "刪除前後空白:  [ " & Trim(strText) & " ]" & Chr(10) & _
           "刪除前空白  :  [ " & LTrim(strText) & " ]" & Chr(10) & _
           "刪除後空白  :  [ " & RTrim(strText) & " ]  "
End Sub




