VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub �R�����()
    Dim myCn As New ADODB.Connection
    Dim myRs As New ADODB.Recordset
    Dim mySQL As String
    myCn.Provider = "Microsoft.Jet.OLEDB.4.0"
    myCn.Open ThisWorkbook.Path & "\�Ȥ�޲z.mdb"
    mySQL = "SELECT * FROM T_�Ȥ� WHERE �Ȥ�W�� Like '%�u�ȩ�'"
    myRs.Open mySQL, myCn, adOpenDynamic, adLockOptimistic
    If myRs.EOF Then
       MsgBox "�S���O��!!"
    Else
       Do While myRs.EOF = False
           myRs.Delete
           myRs.MoveNext
       Loop
    End If
    myRs.Close: myCn.Close
    Set myRs = Nothing: Set myCn = Nothing
End Sub



