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

Sub 刪除資料()
    Dim myCn As New ADODB.Connection
    Dim myRs As New ADODB.Recordset
    Dim mySQL As String
    myCn.Provider = "Microsoft.Jet.OLEDB.4.0"
    myCn.Open ThisWorkbook.Path & "\客戶管理.mdb"
    mySQL = "SELECT * FROM T_客戶 WHERE 客戶名稱 Like '%工務所'"
    myRs.Open mySQL, myCn, adOpenDynamic, adLockOptimistic
    If myRs.EOF Then
       MsgBox "沒有記錄!!"
    Else
       Do While myRs.EOF = False
           myRs.Delete
           myRs.MoveNext
       Loop
    End If
    myRs.Close: myCn.Close
    Set myRs = Nothing: Set myCn = Nothing
End Sub




