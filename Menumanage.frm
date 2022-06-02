VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menumanage 
   Caption         =   "菜單管理"
   ClientHeight    =   4291
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   7735
   OleObjectBlob   =   "Menumanage.frx":0000
   StartUpPosition =   1  '所屬視窗中央
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
'新增菜單品項

    '先檢查是否有空白欄，有空白欄不可以送出
    If Txt_Menuname.Text = "" Then
        MsgBox ("請輸入名稱!")
        Exit Sub
    End If
    If Com_Menutype = Empty Then
        MsgBox ("請選擇類別!")
        Exit Sub
    End If
    If Txt_Menuprice.Text = "" Then
        MsgBox ("請輸入售價!")
        Exit Sub
    ElseIf Txt_Menuprice.Value < 0 Then
        MsgBox ("價格輸入錯誤，請重新輸入!")
        Exit Sub
    End If
    If Txt_Menucost = Empty Then
        MsgBox ("請輸入成本!")
        Exit Sub
    ElseIf Txt_Menucost.Value < 0 Then
        MsgBox ("成本輸入錯誤，請重新輸入!")
        Exit Sub
    End If
    
    '限制售價成本只能輸入數字
    If Not IsNumeric(Txt_Menuprice) Then
        MsgBox ("請輸入正確售價!")
        Exit Sub
    ElseIf Not IsNumeric(Txt_Menucost) Then
        MsgBox ("請輸入正確成本!")
        Exit Sub
    End If
    
    '限制成本不能大於售價
    If Txt_Menuprice.Text < Txt_Menucost.Value Then
        MsgBox ("成本高於售價，請再次確認")
        Txt_Menucost.Value = Null
        Exit Sub
    End If
    
    '菜品新增至工作表
    Sheets("菜單管理").Select
    
    Dim menurow As Integer
    menurow = ActiveSheet.UsedRange.Rows.Count + 1
    
    '先確定工作表內沒有重複的品項
    Dim cnt As Integer
    Dim result As Byte
    
    For cnt = 1 To menurow
        If Cells(cnt, "C").Value = Txt_Menuname.Text Then
            result = MsgBox("菜單已經有相同的品項了，要修改菜單嗎?", vbYesNo)
            If result = 6 Then '新增菜品
                '自動新增當天日期至工作表
                    Cells(cnt, "A").Value = Date
    
                '新增登記人
                 Cells(cnt, "B").Value = Signin.Txt_Username.Text
    
                '新增名稱
                 Cells(cnt, "C").Value = Txt_Menuname.Text
    
                '新增品項
                Cells(cnt, "D").Value = Com_Menutype.Value
    
                '新增售價
                Cells(cnt, "E").Value = Txt_Menuprice.Text
    
                '新增成本
                  Cells(cnt, "F").Value = Txt_Menucost.Text
                  MsgBox "修改完成"
                  Txt_Menuname.Value = Null
                  Com_Menutype.Value = Null
                  Txt_Menuprice.Value = Null
                  Txt_Menucost.Value = Null
                  Exit For
                  Exit Sub
           ElseIf result = 7 Then
                '取消修改
                MsgBox "不修改"
                Exit For
                Exit Sub
            End If
        
        End If
    Next
    
    '沒有重複品項直接新增
    If result = 0 Then
        
        '自動新增當天日期至工作表
        Cells(menurow, "A").Value = Date
    
        '新增登記人
        Cells(menurow, "B").Value = Signin.Txt_Username.Text
    
        '新增名稱
        Cells(menurow, "C").Value = Txt_Menuname.Text
    
        '新增品項
         Cells(menurow, "D").Value = Com_Menutype.Value
    
        '新增售價
        Cells(menurow, "E").Value = Txt_Menuprice.Text
    
        '新增成本
         Cells(menurow, "F").Value = Txt_Menucost.Text
         
         MsgBox "新增完成"
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
'刪除菜單品項
    
    '名稱空白不能輸入
    If Txt_Menuname.Text = "" Then
        MsgBox ("請輸入名稱!")
        Exit Sub
    End If
    
   '刪除品項
    Sheets("菜單管理").Select
    Dim cnt As Integer
    Dim result As Byte
    Dim menurow As Integer
    menurow = ActiveSheet.UsedRange.Rows.Count
    
    For cnt = 1 To menurow
        If Cells(cnt, "C").Value = Txt_Menuname.Text Then
            result = MsgBox("確定有刪除該品項嗎?", vbYesNo)
            If result = 6 Then '確定刪除
                Rows(cnt).EntireRow.Delete
                  MsgBox "刪除完畢"
                  Txt_Menuname.Value = Null
                  Exit For
                  Exit Sub
           ElseIf result = 7 Then
                '取消刪除
                Exit For
                Exit Sub
            End If
        End If
    Next
    
    If result = 0 Then
        MsgBox "菜單沒有該品項，請重新確認"
        Exit Sub
    End If

End Sub

Private Sub UserForm_Initialize()
'菜單管理初始化界面

    '類別下拉式選單
    Com_Menutype.AddItem "麵食"
    Com_Menutype.AddItem "飲料"
    Com_Menutype.AddItem "點心"
    
End Sub
