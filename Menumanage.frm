VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menumanage 
   Caption         =   "���޲z"
   ClientHeight    =   4291
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   7735
   OleObjectBlob   =   "Menumanage.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "Menumanage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub But_Bookingback_Click()
    Menumanage.Hide
    Signin.Show
End Sub

Private Sub But_Menucheck_Click()
'�s�W���~��

    '���ˬd�O�_���ť���A���ť��椣�i�H�e�X
    If Txt_Menuname.Text = "" Then
        MsgBox ("�п�J�W��!")
        Exit Sub
    End If
    If Com_Menutype = Empty Then
        MsgBox ("�п�����O!")
        Exit Sub
    End If
    If Txt_Menuprice.Text = "" Then
        MsgBox ("�п�J���!")
        Exit Sub
    ElseIf Txt_Menuprice.Value < 0 Then
        MsgBox ("�����J���~�A�Э��s��J!")
        Exit Sub
    End If
    If Txt_Menucost = Empty Then
        MsgBox ("�п�J����!")
        Exit Sub
    ElseIf Txt_Menucost.Value < 0 Then
        MsgBox ("������J���~�A�Э��s��J!")
        Exit Sub
    End If
    
    '�����������u���J�Ʀr
    If Not IsNumeric(Txt_Menuprice) Then
        MsgBox ("�п�J���T���!")
        Exit Sub
    ElseIf Not IsNumeric(Txt_Menucost) Then
        MsgBox ("�п�J���T����!")
        Exit Sub
    End If
    
    '���������j����
    If Txt_Menuprice.Text < Txt_Menucost.Value Then
        MsgBox ("�����������A�ЦA���T�{")
        Txt_Menucost.Value = Null
        Exit Sub
    End If
    
    '��~�s�W�ܤu�@��
    Sheets("���޲z").Select
    
    Dim menurow As Integer
    menurow = ActiveSheet.UsedRange.Rows.Count + 1
    
    '���T�w�u�@���S�����ƪ��~��
    Dim cnt As Integer
    Dim result As Byte
    
    For cnt = 1 To menurow
        If Cells(cnt, "C").Value = Txt_Menuname.Text Then
            result = MsgBox("���w�g���ۦP���~���F�A�n�ק����?", vbYesNo)
            If result = 6 Then '�s�W��~
                '�۰ʷs�W��Ѥ���ܤu�@��
                    Cells(cnt, "A").Value = Date
    
                '�s�W�n�O�H
                 Cells(cnt, "B").Value = Signin.Txt_Username.Text
    
                '�s�W�W��
                 Cells(cnt, "C").Value = Txt_Menuname.Text
    
                '�s�W�~��
                Cells(cnt, "D").Value = Com_Menutype.Value
    
                '�s�W���
                Cells(cnt, "E").Value = Txt_Menuprice.Text
    
                '�s�W����
                  Cells(cnt, "F").Value = Txt_Menucost.Text
                  MsgBox "�ק粒��"
                  Txt_Menuname.Value = Null
                  Com_Menutype.Value = Null
                  Txt_Menuprice.Value = Null
                  Txt_Menucost.Value = Null
                  Exit For
                  Exit Sub
           ElseIf result = 7 Then
                '�����ק�
                MsgBox "���ק�"
                Exit For
                Exit Sub
            End If
        
        End If
    Next
    
    '�S�����ƫ~�������s�W
    If result = 0 Then
        
        '�۰ʷs�W��Ѥ���ܤu�@��
        Cells(menurow, "A").Value = Date
    
        '�s�W�n�O�H
        Cells(menurow, "B").Value = Signin.Txt_Username.Text
    
        '�s�W�W��
        Cells(menurow, "C").Value = Txt_Menuname.Text
    
        '�s�W�~��
         Cells(menurow, "D").Value = Com_Menutype.Value
    
        '�s�W���
        Cells(menurow, "E").Value = Txt_Menuprice.Text
    
        '�s�W����
         Cells(menurow, "F").Value = Txt_Menucost.Text
         
         MsgBox "�s�W����"
         Txt_Menuname.Value = Null
         Com_Menutype.Value = Null
         Txt_Menuprice.Value = Null
         Txt_Menucost.Value = Null
         Exit Sub
    
    End If
    
End Sub

Private Sub But_Menuclear_Click()

    Txt_Menuname.Value = Null
    Com_Menutype.Value = Null
    Txt_Menuprice.Value = Null
    Txt_Menucost.Value = Null
End Sub

Private Sub But_Menudel_Click()
'�R�����~��
    
    '�W�٪ťդ����J
    If Txt_Menuname.Text = "" Then
        MsgBox ("�п�J�W��!")
        Exit Sub
    End If
    
   '�R���~��
    Sheets("���޲z").Select
    Dim cnt As Integer
    Dim result As Byte
    Dim menurow As Integer
    menurow = ActiveSheet.UsedRange.Rows.Count
    
    For cnt = 1 To menurow
        If Cells(cnt, "C").Value = Txt_Menuname.Text Then
            result = MsgBox("�T�w���R���ӫ~����?", vbYesNo)
            If result = 6 Then '�T�w�R��
                Rows(cnt).EntireRow.Delete
                  MsgBox "�R������"
                  Txt_Menuname.Value = Null
                  Exit For
                  Exit Sub
           ElseIf result = 7 Then
                '�����R��
                Exit For
                Exit Sub
            End If
        End If
    Next
    
    If result = 0 Then
        MsgBox "���S���ӫ~���A�Э��s�T�{"
        Exit Sub
    End If

End Sub

Private Sub UserForm_Initialize()
'���޲z��l�Ƭɭ�

    '���O�U�Ԧ����
    Com_Menutype.AddItem "�ѭ�"
    Com_Menutype.AddItem "����"
    Com_Menutype.AddItem "�I��"
    
End Sub
