VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Bookingseat 
   Caption         =   "�w���y��"
   ClientHeight    =   6216
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   8806.001
   OleObjectBlob   =   "Bookingseat.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "Bookingseat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub But_Bookingcheck_Click()
'�e�X�w��
    
    '�ˬd�O�_���ť�
    If Txt_Bookingname.Value = Empty Then
        MsgBox "�п�J�w���H�m�W!"
        Exit Sub
    End If
    
    '�ˬd�n�O����O�_���T
    If Com_Bookingyear.Value = Empty Then
        MsgBox "�п�J���T�~��!"
        Exit Sub
    End If
    If Com_Bookingmonth.Value = Empty Then
        MsgBox "�п�J���!"
        Exit Sub
    End If
    If Com_Bookingday.Value = Empty Then
        MsgBox "�п�J���!"
        Exit Sub
    End If
    If Com_Bookinghour.Value = Empty Then
        MsgBox "�п�J�ɶ�!"
        Exit Sub
    End If
    If Com_Bookinghour.Value = Empty Then
        MsgBox "�п�J�ɶ�!"
        Exit Sub
    End If
    
    '����O�_�X�z
    Dim bookingdate As String
    Dim bookingtime As String
    Dim mytime
    
    bookingdate = Com_Bookingyear.Value & "/" & Com_Bookingmonth.Value & "/" & Com_Bookingday.Value
    bookingtime = Com_Bookinghour.Value & ":" & Com_Bookingmin.Value & Com_Bookingning.Value
    mytime = TimeValue(bookingtime)
    
    If bookingdate < Date Then
        MsgBox ("�����J���~")
        Exit Sub
    ElseIf bookingdate = Date Then
       
        If mytime < Time Or bookingtime = Time Then
            MsgBox ("�ɶ���J���~")
            Com_Bookinghour.Value = Null
            Com_Bookingmin.Value = Null
            Exit Sub
        End If
    End If

    '2�뤣��j��28��
    If Com_Bookingmonth.Value = 2 Then
        If Com_Bookingday.Value < 28 Then
        MsgBox "2��j��28�ѡA�Э��s��J���T���!"
        Com_Bookingday.Value = Null
        Exit Sub
        End If
    End If
    
    '30���٬O31��
    Select Case Com_Bookingmonth.Value
    Case 2, 4, 6, 9, 11
        If Com_Bookingday.Value = 31 Then
            MsgBox ("�Ӥ���S��31��")
            Com_Bookingday.Value = Null
            Exit Sub
        End If
    End Select
    
    'AMPM���S���g�W
    If Com_Bookingning = Empty Then
        MsgBox ("�ж�g�w���ɬq")
        Exit Sub
    End If
    
    '����H�ƥu���J�Ʀr
    If Not IsNumeric(Txt_Bookingpeople) Then
        MsgBox ("�п�J���T�H��!")
        Exit Sub
    End If
    
    
    '�s�W�ܤu�@��
    
    Sheets("�w���n�O").Select
    
    Dim bookrow As Integer
    Sheets("�w���n�O").Select
    bookrow = ActiveSheet.UsedRange.Rows.Count + 1
    
    Dim cnt As Integer
    Dim result As Byte
    Dim notice As Integer
    
    '���T�w��ѨS���ۦP���H�w��
    
    For cnt = 2 To bookrow
        If Cells(cnt, "C").Value = Txt_Bookingname.Value Then
            If Cells(cnt, "D").Value = bookingdate Then
                result = MsgBox("�w�g���d�ߨ��Ѧ��w���A�T�w�٭n�A�w����?", vbYesNo)
                If result = 6 Then '�s�W�w��
                    Cells(cnt + 1, "A").Value = Date
                    Cells(cnt + 1, "B").Value = Signin.Txt_Username.Value
                    Cells(cnt + 1, "C").Value = Txt_Bookingname.Value
                    Cells(cnt + 1, "D").Value = bookingdate
                    Cells(cnt + 1, "E").Value = bookingtime
                    Cells(cnt + 1, "F").Value = Txt_Bookingpeople.Value
                    Cells(cnt + 1, "G").Value = Txt_Bookingtel.Text
                    Cells(cnt + 1, "H").Value = Txt_Bookingps.Value
                    MsgBox "�w������"
                    Txt_Bookingname.Value = Null
                    Txt_Bookingpeople.Value = Null
                    Txt_Bookingpeople.Value = Null
                    Txt_Bookingtel.Value = Null
                    Txt_Bookingps.Value = Null
                    Com_Bookingyear.Value = Null
                    Com_Bookingmonth.Value = Null
                    Com_Bookingday.Value = Null
                    Com_Bookinghour.Value = Null
                    Com_Bookingmin.Value = Null
                    notice = 1
                    Exit Sub
                ElseIf result = 7 Then
                    MsgBox "�����w��"
                    Exit Sub
                End If
                Exit Sub
            End If
        
        End If
    Next
    
    cnt = bookrow
    
    '�S�����ƴN�����w��
    If notice = 0 Then
        If result = 0 Then
            Cells(bookrow, "A").Value = Date
            Cells(bookrow, "B").Value = Signin.Txt_Username.Value
            Cells(bookrow, "C").Value = Txt_Bookingname.Value
            Cells(bookrow, "D").Value = bookingdate
            Cells(bookrow, "E").Value = bookingtime
            Cells(bookrow, "F").Value = Txt_Bookingpeople.Value
            Cells(bookrow, "G").Value = Txt_Bookingtel.Text
            Cells(bookrow, "H").Value = Txt_Bookingps.Value
            MsgBox "�w������"
            Txt_Bookingname.Value = Null
            Txt_Bookingpeople.Value = Null
            Txt_Bookingpeople.Value = Null
            Txt_Bookingtel.Value = Null
            Txt_Bookingps.Value = Null
            Com_Bookingyear.Value = Null
            Com_Bookingmonth.Value = Null
            Com_Bookingday.Value = Null
            Com_Bookinghour.Value = Null
            Com_Bookingmin.Value = Null
            Exit Sub
        End If
    End If
End Sub

Private Sub But_Bookingclear_Click()
'�����M��
    Txt_Bookingname.Value = Null
    Txt_Bookingpeople.Value = Null
    Txt_Bookingpeople.Value = Null
    Txt_Bookingtel.Value = Null
    Txt_Bookingps.Value = Null
    Com_Bookingyear.Value = Null
    Com_Bookingmonth.Value = Null
    Com_Bookingday.Value = Null
    Com_Bookinghour.Value = Null
    Com_Bookingmin.Value = Null
End Sub

Private Sub But_Bookingback_Click()
'��^����
    Bookingseat.Hide
    Signin.Show
End Sub

Private Sub UserForm_Initialize()

'�w���y����l�Ƭɭ�

    '�w������U�Ԧ����
    Dim booking_year As Integer
    For booking_year = 2019 To 2023
        Com_Bookingyear.AddItem booking_year
    Next
    
    Dim booking_month As Integer
    For booking_month = 1 To 12
        Com_Bookingmonth.AddItem booking_month
    Next
    
    Dim booking_day As Integer
    For booking_day = 1 To 31
        Com_Bookingday.AddItem booking_day
    Next
    
    '�w���ɶ��U�Ԧ����
    Dim booking_hour As Integer
    For booking_hour = 1 To 12
        Com_Bookinghour.AddItem booking_hour
    Next
    
    Dim booking_min As Integer
    For booking_min = 0 To 59
        Com_Bookingmin.AddItem booking_min
        booking_min = booking_min + 14
    Next
    
    Com_Bookingning.AddItem "AM"
    Com_Bookingning.AddItem "PM"
        
    
    '�ɶ��w�]���
    Com_Bookingyear.Value = year(Date)
    Com_Bookingmonth.Value = Month(Date)
    Com_Bookingday.Value = Day(Date)
    
End Sub
