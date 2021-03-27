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

Sub 追加資料()
    Dim myCn As New ADODB.Connection
    Dim myRs As New ADODB.Recordset
    myCn.Provider = "Microsoft.Jet.OLEDB.4.0"
    myCn.Open ThisWorkbook.Path & "\客戶管理.mdb"
    myRs.Open "T_客戶", myCn, adOpenDynamic, adLockOptimistic
    myRs.AddNew
    
    myRs!客戶編號 = Range("A2").Value
    myRs!客戶名稱 = Range("B2").Value
    myRs!負責人 = Range("C2").Value
    myRs!郵遞區號 = Range("D2").Value
    myRs!住址 = Range("B2").Value
    myRs!電話 = Range("F2").Value
    myRs.Update
    
    myRs.Close: myCn.Close
    Set myRs = Nothing: Set myCn = Nothing
End Sub

