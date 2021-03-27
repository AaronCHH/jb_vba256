VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "讀取影像"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6795
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dRange As Range '宣告模組等級的變數dRange

Sub readrecord(no As Long)
    Dim rw As Long
   
   '查詢表格第1欄"NO"的儲存格,接著指定給dRange
    Set dRange = Range("A1").CurrentRegion.Columns(1).Find(What:=no, LookAt:=xlWhole)
    If dRange Is Nothing Then
        MsgBox "沒有資料!!"
        Exit Sub
    End If
   
   '取得NO欄位的列編號,並指定給rw
    rw = dRange.Row
    Label1.Caption = Cells(rw, 1).Value
    TextBox1.Text = Cells(rw, 2).Value
    If Cells(rw, 3).Value = "男" Then
       OptionButton1.Value = True
    Else
       OptionButton2.Value = True
    End If
    
   '將對應的列的每個值,讀進各個控制項中
    CheckBox1.Value = Cells(rw, 4).Value
    CheckBox2.Value = Cells(rw, 5).Value
    ComboBox1.Text = Cells(rw, 6).Value
    ListBox1.Value = Cells(rw, 7).Value
       
End Sub

Private Sub search_Click()
    Dim i As Long, no As Long
    If IsNumeric(TextBox2.Text) = False Then
       MsgBox "請輸入編號!!"
       Exit Sub
    End If
    For i = 2 To Range("A1").CurrentRegion.Rows.Count
        If Cells(i, 1).Text = TextBox2.Text Then
            no = Val(TextBox2.Text)
            Exit For
        End If
    Next
    If no = 0 Then
        MsgBox "指定的編號不存在!!"
        Exit Sub
    Else
        Call readrecord(no)
    End If
End Sub

Private Sub delete_Click()
    Dim ans As Integer, no As Long, srcRange As Range
    ans = MsgBox(Label1.Caption & ":" & TextBox1.Text & _
          " 的記錄可以刪除嗎??", _
          vbOKCancel + vbExclamation, "刪除確認")
    If ans = vbCancel Then Exit Sub
    Set srcRange = Range("A1").CurrentRegion
    If dRange.Row = srcRange.Rows.Count Then
       no = dRange.Offset(-1).Value
    Else
       no = dRange.Offset(1).Value
    End If
    srcRange.Rows(dRange.Row).delete Shift:=xlShiftUp
    Call readrecord(no)
    Set srcRange = Nothing
    
End Sub

Private Sub first_Click()
   '起始:將第1件儲存格A2的值(1)當作參數,呼叫readrecord
    Call readrecord(Range("A2").Value)
End Sub
    
Private Sub prev_Click()
    If dRange.Row = 2 Then
       MsgBox "起始資料錄!!"
    Else
      '前:將第目前儲存格的上1個儲存格的值(1)當作參數,呼叫readrecord
       Call readrecord(dRange.Offset(-1).Value)
    End If
End Sub

Private Sub exit1_Click()
   Unload Me
End Sub

Private Sub register_Click()
    Dim obj As Object, ans As Integer, rw As Long
   
   '查詢每個控制項有沒有空的,如果發現有則轉移到myMessage處理
    For Each obj In UserForm1.Controls
        Select Case TypeName(obj)
            Case "TextBox"
                If Not (obj.Name = "TextBox2") Then
                   If obj.Text = "" Then GoTo myMessage
                End If
            Case "ListBox", "ComboBox"
                If obj.ListIndex = -1 Then GoTo myMessage
        End Select
    Next
    
   '如果2個選項按鈕都是False,則轉移到myMessage處理
    If OptionButton1.Value = False And OptionButton2.Value = False Then
        GoTo myMessage
    End If
   
   '登錄確認
    ans = MsgBox("可以登錄資料嗎??", vbOKCancel, "登錄確認")
    If ans = vbCancel Then Exit Sub
    rw = dRange.Row
    Cells(rw, 1).Value = Label1.Caption
    
   '將各個控制項的值,寫入對應的儲存格中
    Cells(rw, 2).Value = TextBox1.Text
    If OptionButton1.Value Then
        Cells(rw, 3).Value = OptionButton1.Caption
    Else
        Cells(rw, 3).Value = OptionButton2.Caption
    End If
    Cells(rw, 4).Value = CheckBox1.Value
    Cells(rw, 5).Value = CheckBox2.Value
    Cells(rw, 6).Value = ComboBox1.Text
    Cells(rw, 7).Value = ListBox1.List(ListBox1.ListIndex)
    Exit Sub
    
myMessage:
    MsgBox ("遺漏資料!!")
    Exit Sub
End Sub

Private Sub new1_Click()
    Dim obj As Control
     
   '重設每個控制項的值
    For Each obj In UserForm1.Controls
        Select Case TypeName(obj)
            Case "TextBox"
                obj.Text = ""
            Case "ListBox", "ComboBox"
                obj.ListIndex = -1
            Case "OptionButton", "CheckBox"
                obj.Value = False
        End Select
    Next
    With Range("A1").CurrentRegion
        '在標籤上設定新列的編號
         Label1.Caption = .Cells(.Rows.Count, 1) + 1
        '將新列的A欄儲存格,指定給變數dRange
         Set dRange = .Cells(.Rows.Count, 1).Offset(1)
    End With
    
End Sub

Private Sub next1_Click()
    If dRange.Row = Range("A1").CurrentRegion.Rows.Count Then
       MsgBox "終止資料錄!!"
    Else
      '後:將第目前儲存格的下1個儲存格的值(1)當作參數,呼叫readrecord
       Call readrecord(dRange.Offset(1).Value)
    End If
End Sub

Private Sub last_Click()
    Dim no As Long
    With Range("A1").CurrentRegion
        '將表格的A欄的最後列的值,指定給變數no
         no = .Cells(.Rows.Count, 1).Value
    End With
      
   '終止:將第目前儲存格的下1個儲存格的值(1)當作參數,呼叫readrecord
    Call readrecord(no)

End Sub

Private Sub UserForm_Initialize()
    
    ComboBox1.RowSource = Range("N2:N27").Address
    ListBox1.RowSource = Range("O2:O9").Address
   '將第1件儲存格A2的值(1)當作參數,呼叫readrecord
    Call readrecord(Range("A2").Value)
        
End Sub

