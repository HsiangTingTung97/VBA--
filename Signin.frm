VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Signin 
   Caption         =   "智慧點餐系統"
   ClientHeight    =   5495
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   7742
   OleObjectBlob   =   "Signin.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "Signin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub But_Signinbook_Click()
    
    '檢查登記人欄位是否空白
    If Txt_Username.Text = Empty Then
        MsgBox "請輸入名稱!"
        Exit Sub
    End If
    
    Signin.Hide
    Bookingseat.Show
    
End Sub

Private Sub But_Signinmeal_Click()
    '檢查登記人欄位是否空白
    If Txt_Username.Text = Empty Then
        MsgBox "請輸入名稱!"
        Exit Sub
    End If
    
    '檢查登記日期是否正確
    If Com_Signinyear.Value = Empty Then
        MsgBox "請輸入正確年份!"
        Exit Sub
    End If
    If Com_Signinmonth.Value = Empty Then
        MsgBox "請輸入月份!"
        Exit Sub
    End If
    If Com_Signinday.Value = Empty Then
        MsgBox "請輸入日期!"
        Exit Sub
    End If
    
    Signin.Hide
    Salechart.Show
    
End Sub

Private Sub But_Signinmenu_Click()

    '檢查登記人欄位是否空白
    If Txt_Username.Text = Empty Then
        MsgBox "請輸入名稱!"
        Exit Sub
    End If
    
    '檢查登記日期是否正確
    If Com_Signinyear.Value = Empty Then
        MsgBox "請輸入正確年份!"
        Exit Sub
    End If
    If Com_Signinmonth.Value = Empty Then
        MsgBox "請輸入月份!"
        Exit Sub
    End If
    If Com_Signinday.Value = Empty Then
        MsgBox "請輸入日期!"
        Exit Sub
    End If
    
    '進入其菜單管理工作表
    Signin.Hide
    Menumanage.Show
End Sub

Private Sub But_Signinorder_Click()
    '檢查登記人欄位是否空白
    If Txt_Username.Text = Empty Then
        MsgBox "請輸入名稱!"
        Exit Sub
    End If
    
    Signin.Hide
    Order.Show
    
End Sub

Private Sub UserForm_Initialize()
    '登入表單初始化界面

    '登入時間下拉式選單
    For signin_year = 2019 To 2028
        Com_Signinyear.AddItem signin_year
    Next
    
    For signin_month = 1 To 12
        Com_Signinmonth.AddItem signin_month
    Next
    
    For signin_day = 1 To 31
        Com_Signinday.AddItem signin_day
    Next
    
    '時間預設當天
    Com_Signinyear.Value = year(Date)
    Com_Signinmonth.Value = Month(Date)
    Com_Signinday.Value = Day(Date)
     

End Sub
