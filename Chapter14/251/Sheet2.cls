VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub 複製資料表()
    Dim myCn As New ADODB.Connection
    Dim myRs As New ADODB.Recordset
    myCn.Provider = "Microsoft.Jet.OLEDB.4.0"
    myCn.Open ThisWorkbook.Path & "\客戶管理.mdb"
    myRs.Open "T_客戶", myCn
    Range("A2").CopyFromRecordset myRs
    myRs.Close: myCn.Close
    Set myRs = Nothing: Set myCn = Nothing
End Sub

Private Sub CommandButton1_Click()
    Call 複製資料表
End Sub
