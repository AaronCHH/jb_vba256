VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Ū���v��"
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
Dim dRange As Range '�ŧi�Ҳյ��Ū��ܼ�dRange

Sub readrecord(no As Long)
    Dim rw As Long
   
   '�d�ߪ���1��"NO"���x�s��,���۫��w��dRange
    Set dRange = Range("A1").CurrentRegion.Columns(1).Find(What:=no, LookAt:=xlWhole)
    If dRange Is Nothing Then
        MsgBox "�S�����!!"
        Exit Sub
    End If
   
   '���oNO��쪺�C�s��,�ë��w��rw
    rw = dRange.Row
    Label1.Caption = Cells(rw, 1).Value
    TextBox1.Text = Cells(rw, 2).Value
    If Cells(rw, 3).Value = "�k" Then
       OptionButton1.Value = True
    Else
       OptionButton2.Value = True
    End If
    
   '�N�������C���C�ӭ�,Ū�i�U�ӱ����
    CheckBox1.Value = Cells(rw, 4).Value
    CheckBox2.Value = Cells(rw, 5).Value
    ComboBox1.Text = Cells(rw, 6).Value
    ListBox1.Value = Cells(rw, 7).Value
       
End Sub

Private Sub search_Click()
    Dim i As Long, no As Long
    If IsNumeric(TextBox2.Text) = False Then
       MsgBox "�п�J�s��!!"
       Exit Sub
    End If
    For i = 2 To Range("A1").CurrentRegion.Rows.Count
        If Cells(i, 1).Text = TextBox2.Text Then
            no = Val(TextBox2.Text)
            Exit For
        End If
    Next
    If no = 0 Then
        MsgBox "���w���s�����s�b!!"
        Exit Sub
    Else
        Call readrecord(no)
    End If
End Sub

Private Sub delete_Click()
    Dim ans As Integer, no As Long, srcRange As Range
    ans = MsgBox(Label1.Caption & ":" & TextBox1.Text & _
          " ���O���i�H�R����??", _
          vbOKCancel + vbExclamation, "�R���T�{")
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
   '�_�l:�N��1���x�s��A2����(1)��@�Ѽ�,�I�sreadrecord
    Call readrecord(Range("A2").Value)
End Sub
    
Private Sub prev_Click()
    If dRange.Row = 2 Then
       MsgBox "�_�l��ƿ�!!"
    Else
      '�e:�N�ĥثe�x�s�檺�W1���x�s�檺��(1)��@�Ѽ�,�I�sreadrecord
       Call readrecord(dRange.Offset(-1).Value)
    End If
End Sub

Private Sub exit1_Click()
   Unload Me
End Sub

Private Sub register_Click()
    Dim obj As Object, ans As Integer, rw As Long
   
   '�d�ߨC�ӱ�����S���Ū�,�p�G�o�{���h�ಾ��myMessage�B�z
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
    
   '�p�G2�ӿﶵ���s���OFalse,�h�ಾ��myMessage�B�z
    If OptionButton1.Value = False And OptionButton2.Value = False Then
        GoTo myMessage
    End If
   
   '�n���T�{
    ans = MsgBox("�i�H�n����ƶ�??", vbOKCancel, "�n���T�{")
    If ans = vbCancel Then Exit Sub
    rw = dRange.Row
    Cells(rw, 1).Value = Label1.Caption
    
   '�N�U�ӱ������,�g�J�������x�s�椤
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
    MsgBox ("��|���!!")
    Exit Sub
End Sub

Private Sub new1_Click()
    Dim obj As Control
     
   '���]�C�ӱ������
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
        '�b���ҤW�]�w�s�C���s��
         Label1.Caption = .Cells(.Rows.Count, 1) + 1
        '�N�s�C��A���x�s��,���w���ܼ�dRange
         Set dRange = .Cells(.Rows.Count, 1).Offset(1)
    End With
    
End Sub

Private Sub next1_Click()
    If dRange.Row = Range("A1").CurrentRegion.Rows.Count Then
       MsgBox "�פ��ƿ�!!"
    Else
      '��:�N�ĥثe�x�s�檺�U1���x�s�檺��(1)��@�Ѽ�,�I�sreadrecord
       Call readrecord(dRange.Offset(1).Value)
    End If
End Sub

Private Sub last_Click()
    Dim no As Long
    With Range("A1").CurrentRegion
        '�N��檺A�檺�̫�C����,���w���ܼ�no
         no = .Cells(.Rows.Count, 1).Value
    End With
      
   '�פ�:�N�ĥثe�x�s�檺�U1���x�s�檺��(1)��@�Ѽ�,�I�sreadrecord
    Call readrecord(no)

End Sub

Private Sub UserForm_Initialize()
    
    ComboBox1.RowSource = Range("N2:N27").Address
    ListBox1.RowSource = Range("O2:O9").Address
   '�N��1���x�s��A2����(1)��@�Ѽ�,�I�sreadrecord
    Call readrecord(Range("A2").Value)
        
End Sub

