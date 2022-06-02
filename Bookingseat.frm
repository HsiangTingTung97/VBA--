VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Bookingseat 
   Caption         =   "預約座位"
   ClientHeight    =   6216
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   8806.001
   OleObjectBlob   =   "Bookingseat.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "Bookingseat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub But_Bookingcheck_Click()
'送出預約
    
    '檢查是否有空白
    If Txt_Bookingname.Value = Empty Then
        MsgBox "請輸入預約人姓名!"
        Exit Sub
    End If
    
    '檢查登記日期是否正確
    If Com_Bookingyear.Value = Empty Then
        MsgBox "請輸入正確年份!"
        Exit Sub
    End If
    If Com_Bookingmonth.Value = Empty Then
        MsgBox "請輸入月份!"
        Exit Sub
    End If
    If Com_Bookingday.Value = Empty Then
        MsgBox "請輸入日期!"
        Exit Sub
    End If
    If Com_Bookinghour.Value = Empty Then
        MsgBox "請輸入時間!"
        Exit Sub
    End If
    If Com_Bookinghour.Value = Empty Then
        MsgBox "請輸入時間!"
        Exit Sub
    End If
    
    '日期是否合理
    Dim bookingdate As String
    Dim bookingtime As String
    Dim mytime
    
    bookingdate = Com_Bookingyear.Value & "/" & Com_Bookingmonth.Value & "/" & Com_Bookingday.Value
    bookingtime = Com_Bookinghour.Value & ":" & Com_Bookingmin.Value & Com_Bookingning.Value
    mytime = TimeValue(bookingtime)
    
    If bookingdate < Date Then
        MsgBox ("日期輸入錯誤")
        Exit Sub
    ElseIf bookingdate = Date Then
       
        If mytime < Time Or bookingtime = Time Then
            MsgBox ("時間輸入錯誤")
            Com_Bookinghour.Value = Null
            Com_Bookingmin.Value = Null
            Exit Sub
        End If
    End If

    '2月不能大於28天
    If Com_Bookingmonth.Value = 2 Then
        If Com_Bookingday.Value < 28 Then
        MsgBox "2月大於28天，請重新輸入正確日期!"
        Com_Bookingday.Value = Null
        Exit Sub
        End If
    End If
    
    '30天還是31天
    Select Case Com_Bookingmonth.Value
    Case 2, 4, 6, 9, 11
        If Com_Bookingday.Value = 31 Then
            MsgBox ("該月份沒有31天")
            Com_Bookingday.Value = Null
            Exit Sub
        End If
    End Select
    
    'AMPM有沒有寫上
    If Com_Bookingning = Empty Then
        MsgBox ("請填寫預約時段")
        Exit Sub
    End If
    
    '限制人數只能輸入數字
    If Not IsNumeric(Txt_Bookingpeople) Then
        MsgBox ("請輸入正確人數!")
        Exit Sub
    End If
    
    
    '新增至工作表
    
    Sheets("預約登記").Select
    
    Dim bookrow As Integer
    Sheets("預約登記").Select
    bookrow = ActiveSheet.UsedRange.Rows.Count + 1
    
    Dim cnt As Integer
    Dim result As Byte
    Dim notice As Integer
    
    '先確定當天沒有相同的人預約
    
    For cnt = 2 To bookrow
        If Cells(cnt, "C").Value = Txt_Bookingname.Value Then
            If Cells(cnt, "D").Value = bookingdate Then
                result = MsgBox("已經有查詢到當天有預約，確定還要再預約嗎?", vbYesNo)
                If result = 6 Then '新增預約
                    Cells(cnt + 1, "A").Value = Date
                    Cells(cnt + 1, "B").Value = Signin.Txt_Username.Value
                    Cells(cnt + 1, "C").Value = Txt_Bookingname.Value
                    Cells(cnt + 1, "D").Value = bookingdate
                    Cells(cnt + 1, "E").Value = bookingtime
                    Cells(cnt + 1, "F").Value = Txt_Bookingpeople.Value
                    Cells(cnt + 1, "G").Value = Txt_Bookingtel.Text
                    Cells(cnt + 1, "H").Value = Txt_Bookingps.Value
                    MsgBox "預約完成"
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
                    MsgBox "取消預約"
                    Exit Sub
                End If
                Exit Sub
            End If
        
        End If
    Next
    
    cnt = bookrow
    
    '沒有重複就直接預約
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
            MsgBox "預約完成"
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
'全部清空
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
'返回介面
    Bookingseat.Hide
    Signin.Show
End Sub

Private Sub UserForm_Initialize()

'預約座位表初始化界面

    '預約日期下拉式選單
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
    
    '預約時間下拉式選單
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
        
    
    '時間預設當天
    Com_Bookingyear.Value = year(Date)
    Com_Bookingmonth.Value = Month(Date)
    Com_Bookingday.Value = Day(Date)
    
End Sub
