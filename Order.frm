VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Order 
   Caption         =   "點餐系統"
   ClientHeight    =   6209
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   8547.001
   OleObjectBlob   =   "Order.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "Order"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub But_Orderback_Click()
    Order.Hide
    Signin.Show
End Sub

Private Sub But_Ordercheck_Click()
    '送出訂單
    Dim recordcnt As Integer
    Dim sheetrow As Integer
    Dim cnt As Integer
    Dim textlen As Integer
    Dim ordercnt As Integer
    Dim ordername As String
    Dim list As String
    Dim row As Integer
    Dim i, x As Integer
    Dim menuprice, orderamount As Integer

    recordcnt = List_Orderrecord.ListCount
    
    For cnt = 1 To recordcnt
        Sheets("銷售紀錄").Select
        sheetrow = ActiveSheet.UsedRange.Rows.Count
        list = List_Orderrecord.list(cnt - 1)
        textlen = Len(list)
        ordername = Left(list, textlen - 2)
        ordercnt = CByte(Left(Right(list, 2), 1))
        Cells(sheetrow + 1, "A").Value = Signin.Txt_Username.Text
        Cells(sheetrow + 1, "B").Value = Date
        Cells(sheetrow + 1, "C").Value = ordername
        Cells(sheetrow + 1, "D").Value = ordercnt
        Sheets("銷售紀錄").Select
        sheetrow = ActiveSheet.UsedRange.Rows.Count
        Sheets("菜單管理").Select
        row = ActiveSheet.UsedRange.Rows.Count
        For i = 1 To sheetrow
            For x = 1 To row
                If Worksheets("銷售紀錄").Cells(i, "C").Value = Worksheets("菜單管理").Cells(x, "C").Value Then
                    menuprice = Worksheets("菜單管理").Cells(x, "E").Value
                    orderamount = Worksheets("銷售紀錄").Cells(i, "D").Value
                    Worksheets("銷售紀錄").Cells(i, "E").Value = menuprice * orderamount
                    Worksheets("銷售紀錄").Cells(i, "F").Value = Worksheets("菜單管理").Cells(x, "F").Value * orderamount
                End If
            Next
        Next
    Next
    MsgBox ("已送出訂單")
    
    '將餐點歸納成類別
    Dim menu_cnt, sale_cnt As Integer
    Sheets("菜單管理").Select
    menu_cnt = ActiveSheet.UsedRange.Rows.Count
    Sheets("銷售紀錄").Select
    sale_cnt = ActiveSheet.UsedRange.Rows.Count
    
    Dim menu_typecnt, sale_typecnt As Integer
    
    For menu_typecnt = 2 To menu_cnt
    
        For sale_typecnt = 2 To sale_cnt
            If (Sheets("銷售紀錄").Cells(sale_typecnt, "C").Value = Sheets("菜單管理").Cells(sale_typecnt, "C").Value) Then
                Sheets("銷售紀錄").Cells(sale_typecnt, "G").Value = Sheets("菜單管理").Cells(sale_typecnt, "D").Value
            End If
        Next
    Next
    
End Sub

Private Sub But_Orderclear_Click()
    List_Orderrecord.Clear
    Com_Orderamount.Value = Null
End Sub

Private Sub Com_Ordernew_Click()
    '限制只能數字輸入
    
    If Not IsNumeric(Com_Orderamount) Or Com_Orderamount.Value = Empty Then
        MsgBox ("請輸入正確數量!")
        Exit Sub
    End If
    
    '將菜品新增至點菜內容
    Dim record As String
    record = List_Ordermenu.Value & Com_Orderamount.Value & "份"
    
    List_Orderrecord.AddItem (record)
    Com_Orderamount.Value = Null

End Sub

Private Sub UserForm_Initialize()
    '點餐表單初始化界面
    
    '菜單隨著工作表更新

    Sheets("菜單管理").Select
    Dim cnt As Integer
    Dim noodlecnt As Integer
    Dim snakecnt As Integer
    Dim drinkcnt As Integer
    Dim ordermenurow As Integer
    Dim noodlearray() As String
    Dim snakearray() As String
    Dim drinkarray() As String
    Dim allarray() As String
    ordermenurow = ActiveSheet.UsedRange.Rows.Count
    
    For cnt = 1 To ordermenurow
        Select Case Cells(cnt, "D").Value
        Case "麵食"
            noodlecnt = noodlecnt + 1
        Case "點心"
            snakecnt = snakecnt + 1
        Case "飲料"
            drinkcnt = drinkcnt + 1
        End Select
    Next
    
    noodlecnt = noodlecnt - 1
    snakecnt = snakecnt - 1
    drinkcnt = drinkcnt - 1
    cnt = ordermenurow - 2
    ReDim Preserve noodlearray(noodlecnt)
    ReDim Preserve snakearray(snakecnt)
    ReDim Preserve drinkarray(drinkcnt)
    ReDim Preserve allarray(cnt)
    
    noodlecnt = -1
    snakecnt = -1
    drinkcnt = -1
    
    For cnt = 2 To ordermenurow
        Select Case Cells(cnt, "D").Value
        Case "麵食"
            noodlearray(noodlecnt + 1) = Cells(cnt, "C").Value
            noodlecnt = noodlecnt + 1
        Case "點心"
            snakearray(snakecnt + 1) = Cells(cnt, "C").Value
            snakecnt = snakecnt + 1
        Case "飲料"
            drinkarray(drinkcnt + 1) = Cells(cnt, "C").Value
            drinkcnt = drinkcnt + 1
        End Select
    Next
    
    
    
    '將三個陣列合併成一個，依據麵食點心飲料順序排序
    Dim allcnt As Integer
    Dim ncnt As Integer
    Dim scnt As Integer
    Dim acnt As Integer
    ncnt = noodlecnt
    scnt = snakecnt
    acnt = noodlecnt + snakecnt
    '麵食
    If noodlecnt > 0 Then
        For allcnt = 0 To noodlecnt
            allarray(allcnt) = noodlearray(allcnt)
        Next
    ElseIf noodlecnt = 0 Then
        allarray(0) = noodlearray(0)
    End If
    
    '點心
    If snakecnt > 0 Then
        If noodlecnt > 0 Or noodlecnt = 0 Then
            For allcnt = 0 To snakecnt
                allarray(ncnt + 1) = snakearray(allcnt)
                ncnt = ncnt + 1
            Next
        ElseIf noodlecnt < 0 Then
            For allcnt = 0 To snakecnt
                allarray(allcnt) = snakearray(allcnt)
            Next
        End If
    ElseIf snakecnt = 0 Then
        If noodlecnt > 0 Or noodlecnt = 0 Then
            allarray(ncnt + 1) = snakearray(0)
        End If
    End If
    
    '飲料
    ncnt = noodlecnt
    
    If drinkcnt > 0 Then
        If snakecnt > 0 Then
            If noodlecnt > 0 Then
                For allcnt = 0 To drinkcnt
                    allarray(acnt + 2) = drinkarray(allcnt)
                    acnt = acnt + 1
                Next
            ElseIf noodlecnt < 0 Then
                For allcnt = 0 To drinkcnt
                    allarray(scnt) = drinkarray(allcnt) '要不要加1
                    scnt = scnt + 1
                Next
            End If
        ElseIf snakecnt = 0 Then
            If noodlecnt > 0 Then
                For allcnt = 0 To drinkcnt
                    allarray(ncnt + 1) = drinkarray(allcnt) '要不要加1
                    ncnt = ncnt + 1
                Next
            ElseIf noodlecnt = 0 Then
                For allcnt = 0 To drinkcnt
                    allarray(allcnt) = drinkarray(allcnt)
                Next
            End If
        End If
    ElseIf drinkcnt = 0 Then '表示有一項飲料品項
        If snakecnt > 0 Or snakecnt = 0 Then
            If noodlecnt > 0 Or noodlecnt = 0 Then
                allarray(acnt + 1) = drinkarray(0)
            End If
        End If
    End If

    '依照品項顯示在ListBox
    List_Ordermenu.list() = allarray
    
    '新增數量至下拉式選單
    Dim amount As Integer
    For amount = 1 To 5
        Com_Orderamount.AddItem amount
    Next
End Sub
