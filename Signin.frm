VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Signin 
   Caption         =   "���z�I�\�t��"
   ClientHeight    =   5495
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   7742
   OleObjectBlob   =   "Signin.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "Signin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub But_Signinbook_Click()
    
    '�ˬd�n�O�H���O�_�ť�
    If Txt_Username.Text = Empty Then
        MsgBox "�п�J�W��!"
        Exit Sub
    End If
    
    Signin.Hide
    Bookingseat.Show
    
End Sub

Private Sub But_Signinmeal_Click()
    '�ˬd�n�O�H���O�_�ť�
    If Txt_Username.Text = Empty Then
        MsgBox "�п�J�W��!"
        Exit Sub
    End If
    
    '�ˬd�n�O����O�_���T
    If Com_Signinyear.Value = Empty Then
        MsgBox "�п�J���T�~��!"
        Exit Sub
    End If
    If Com_Signinmonth.Value = Empty Then
        MsgBox "�п�J���!"
        Exit Sub
    End If
    If Com_Signinday.Value = Empty Then
        MsgBox "�п�J���!"
        Exit Sub
    End If
    
    Signin.Hide
    Salechart.Show
    
End Sub

Private Sub But_Signinmenu_Click()

    '�ˬd�n�O�H���O�_�ť�
    If Txt_Username.Text = Empty Then
        MsgBox "�п�J�W��!"
        Exit Sub
    End If
    
    '�ˬd�n�O����O�_���T
    If Com_Signinyear.Value = Empty Then
        MsgBox "�п�J���T�~��!"
        Exit Sub
    End If
    If Com_Signinmonth.Value = Empty Then
        MsgBox "�п�J���!"
        Exit Sub
    End If
    If Com_Signinday.Value = Empty Then
        MsgBox "�п�J���!"
        Exit Sub
    End If
    
    '�i�J����޲z�u�@��
    Signin.Hide
    Menumanage.Show
End Sub

Private Sub But_Signinorder_Click()
    '�ˬd�n�O�H���O�_�ť�
    If Txt_Username.Text = Empty Then
        MsgBox "�п�J�W��!"
        Exit Sub
    End If
    
    Signin.Hide
    Order.Show
    
End Sub

Private Sub UserForm_Initialize()
    '�n�J����l�Ƭɭ�

    '�n�J�ɶ��U�Ԧ����
    For signin_year = 2019 To 2028
        Com_Signinyear.AddItem signin_year
    Next
    
    For signin_month = 1 To 12
        Com_Signinmonth.AddItem signin_month
    Next
    
    For signin_day = 1 To 31
        Com_Signinday.AddItem signin_day
    Next
    
    '�ɶ��w�]���
    Com_Signinyear.Value = year(Date)
    Com_Signinmonth.Value = Month(Date)
    Com_Signinday.Value = Day(Date)
     

End Sub
